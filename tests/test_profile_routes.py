import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app
import profile_store

RULE = {"bagian_cocok": "Kesimpulan", "awalan": "K", "mulai": 1, "lanjut": "per_bagian", "aktif": True}


class ProfileRoutesTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        patcher = patch.object(profile_store, "PROFILES_DIR", root)
        patcher.start()
        self.addCleanup(patcher.stop)
        log_patcher = patch.object(app, "_log_processing_times")  # jangan menulis ke PostgreSQL sungguhan
        self.log = log_patcher.start()
        self.addCleanup(log_patcher.stop)
        path = root / "dokumen-2d"
        path.mkdir()
        captions = [{
            "source_file": "a.html", "section": "2. Bagian > 2.1 Kesimpulan", "layout_id": "x",
            "type": "headline", "text": "K1. Pesan",
            "penomoran": {"awalan": "K", "nomor": 1, "pemisah": "."},
        }]
        files = {
            "meta": {"title": "Dokumen 2d", "sources": []}, "layouts": [], "row_patterns": [],
            "sections": [], "numbering": [], "captions": captions,
        }
        for key, value in files.items():
            (path / f"{key}.json").write_text(json.dumps(value), encoding="utf-8")
        self.client = app.app.test_client()

    def test_list_and_detail(self):
        self.assertEqual(self.client.get("/api/profiles").json["profiles"][0]["name"], "dokumen-2d")
        detail = self.client.get("/api/profiles/dokumen-2d").json
        self.assertEqual(detail["settings"]["penomoran"][0]["awalan"], "K")
        self.assertFalse(detail["settings_saved"])

    def test_unknown_profile_is_404_and_bad_name_is_400(self):
        self.assertEqual(self.client.get("/api/profiles/tidak-ada").status_code, 404)
        self.assertEqual(self.client.get("/api/profiles/Besar").status_code, 400)

    def test_save_then_reset_settings(self):
        saved = self.client.put("/api/profiles/dokumen-2d/settings", json={"penomoran": [{**RULE, "awalan": "P"}]})
        self.assertEqual((saved.status_code, saved.json["settings"]["penomoran"][0]["awalan"]), (200, "P"))
        self.assertTrue(self.client.get("/api/profiles/dokumen-2d").json["settings_saved"])

        reset = self.client.delete("/api/profiles/dokumen-2d/settings")
        self.assertEqual(reset.json["settings"]["penomoran"][0]["awalan"], "K")
        self.assertFalse(self.client.get("/api/profiles/dokumen-2d").json["settings_saved"])

    def write_style_files(self):
        path = profile_store.PROFILES_DIR / "dokumen-2d"
        (path / "style_samples.json").write_text(
            json.dumps([{"source": "a", "kind": "paragraf", "text": "Teks sampel."}]), encoding="utf-8")
        (path / "style_stats.json").write_text(json.dumps({"blok": {}}), encoding="utf-8")

    def test_style_guide_route_generates_and_saves_with_requested_model(self):
        self.write_style_files()
        with patch.object(app, "BedrockDocumentGenerator") as generator_class:
            generator = generator_class.return_value
            generator.model_id = "model-uji"
            generator.usage_totals = {"style_guide": {"input_tokens": 120, "output_tokens": 30}}
            generator.generate_style_guide.return_value = ({"nada": "formal"}, 1)
            response = self.client.post(
                "/api/profiles/dokumen-2d/style-guide", json={"model_id": " model-uji ", "region": ""})
        self.assertEqual(response.status_code, 200)
        generator_class.assert_called_once_with(model_id="model-uji", region=None)
        generator.generate_style_guide.assert_called_once_with(
            [{"source": "a", "kind": "paragraf", "text": "Teks sampel."}], {"blok": {}})
        self.assertEqual(response.json["panduan"]["panduan"], {"nada": "formal"})
        gaya = self.client.get("/api/profiles/dokumen-2d").json["gaya"]
        self.assertEqual((gaya["panduan"]["model_id"], gaya["usang"]), ("model-uji", False))
        name, status, timing, total_seconds, usage = self.log.call_args.args
        self.assertEqual((name, status), ("dokumen-2d [Panduan gaya]", "done"))
        self.assertIsNotNone(timing["writing"]["seconds"])
        self.assertEqual(usage, {"writing": {"input_tokens": 120, "output_tokens": 30}})

    def test_style_guide_route_reports_bedrock_and_profile_errors_as_400(self):
        self.write_style_files()
        with patch.object(app, "BedrockDocumentGenerator") as generator_class:
            generator_class.return_value.usage_totals = {}
            generator_class.return_value.generate_style_guide.side_effect = app.BedrockError("Bedrock gagal.")
            failed = self.client.post("/api/profiles/dokumen-2d/style-guide", json={})
        self.assertEqual((failed.status_code, failed.json["error"]), (400, "Bedrock gagal."))
        self.assertEqual(self.log.call_args.args[:2], ("dokumen-2d [Panduan gaya]", "error"))  # token terpakai tetap dicatat
        self.assertFalse((profile_store.PROFILES_DIR / "dokumen-2d" / "style_guide.json").exists())

        (profile_store.PROFILES_DIR / "dokumen-2d" / "style_samples.json").unlink()
        self.assertEqual(self.client.post("/api/profiles/dokumen-2d/style-guide", json={}).status_code, 400)
        self.assertEqual(self.client.post("/api/profiles/tidak-ada/style-guide", json={}).status_code, 404)

    STRUCTURES = {"bawaan": "Mei 2025", "daftar": [
        {"nama": "Mei 2025", "bab": ["1. Kesimpulan", "2 Ringkasan", "3. Proyeksi"]},
        {"nama": "Maret 2026", "bab": ["2. Pokok Pembahasan", "3. Rancangan"]},
    ]}

    def test_profile_without_saved_chapters_reports_an_empty_structure_list(self):
        self.assertEqual(self.client.get("/api/profiles/dokumen-2d").json["struktur_bab"], {"bawaan": "", "daftar": []})

    def test_chapter_structures_are_saved_per_profile_and_come_back_with_the_profile(self):
        response = self.client.put("/api/profiles/dokumen-2d/chapters", json=self.STRUCTURES)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json["struktur_bab"], self.STRUCTURES)
        self.assertEqual(self.client.get("/api/profiles/dokumen-2d").json["struktur_bab"], self.STRUCTURES)
        self.assertTrue((profile_store.PROFILES_DIR / "dokumen-2d" / "chapters.json").is_file())
        # pengaturan penomoran tidak ikut terpengaruh oleh struktur bab, dan sebaliknya
        self.assertFalse((profile_store.PROFILES_DIR / "dokumen-2d" / "settings.json").exists())
        self.client.put("/api/profiles/dokumen-2d/settings", json={"penomoran": [RULE], "caption_default": "kosong"})
        self.assertEqual(self.client.get("/api/profiles/dokumen-2d").json["struktur_bab"], self.STRUCTURES)

    def test_chapter_structures_are_normalized_and_the_default_falls_back_to_the_first(self):
        messy = {"bawaan": " maret   2026 ", "daftar": [{"nama": "  Maret   2026 ", "bab": ["  1.  Satu  ", "", "2. Dua"]}]}
        saved = self.client.put("/api/profiles/dokumen-2d/chapters", json=messy).json["struktur_bab"]
        self.assertEqual(saved, {"bawaan": "Maret 2026", "daftar": [{"nama": "Maret 2026", "bab": ["1. Satu", "2. Dua"]}]})
        self.assertEqual(self.client.put("/api/profiles/dokumen-2d/chapters", json={"bawaan": "", "daftar": [
            {"nama": "A", "bab": ["x"]}, {"nama": "B", "bab": ["y"]}]}).json["struktur_bab"]["bawaan"], "A")
        emptied = self.client.put("/api/profiles/dokumen-2d/chapters", json={"bawaan": "A", "daftar": []})
        self.assertEqual(emptied.json["struktur_bab"], {"bawaan": "", "daftar": []})  # menghapus semua struktur

    def test_invalid_chapter_structures_return_clear_400_and_save_nothing(self):
        good = {"nama": "A", "bab": ["1. Satu"]}
        cases = [
            ({"daftar": "x"}, "daftar"),
            ({"bawaan": "", "daftar": [{"nama": "", "bab": ["x"]}]}, "nama wajib diisi"),
            ({"bawaan": "", "daftar": [{"nama": "n" * 61, "bab": ["x"]}]}, "nama wajib diisi"),
            ({"bawaan": "", "daftar": [good, {"nama": "a", "bab": ["y"]}]}, "sudah dipakai"),
            ({"bawaan": "", "daftar": [{"nama": "A", "bab": []}]}, "minimal satu judul"),
            ({"bawaan": "", "daftar": [{"nama": "A", "bab": ["  ", ""]}]}, "minimal satu judul"),
            ({"bawaan": "", "daftar": [{"nama": "A", "bab": [str(i) for i in range(21)]}]}, "maksimal 20 bab"),
            ({"bawaan": "", "daftar": [{"nama": "A", "bab": ["x" * 201]}]}, "maksimal 200 karakter"),
            ({"bawaan": "", "daftar": [{"nama": "A", "bab": [1]}]}, "daftar judul"),
            ({"bawaan": "Tidak Ada", "daftar": [good]}, "tidak ada di daftar"),
            ({"bawaan": "", "daftar": [{"nama": f"N{i}", "bab": ["x"]} for i in range(11)]}, "Maksimal 10"),
        ]
        for body, message in cases:
            response = self.client.put("/api/profiles/dokumen-2d/chapters", json=body)
            self.assertEqual(response.status_code, 400, body)
            self.assertIn(message, response.json["error"], body)
        self.assertEqual(self.client.put("/api/profiles/dokumen-2d/chapters", data="bukan json").status_code, 400)
        self.assertEqual(self.client.put("/api/profiles/tidak-ada/chapters", json=self.STRUCTURES).status_code, 404)
        self.assertFalse((profile_store.PROFILES_DIR / "dokumen-2d" / "chapters.json").exists())

    def test_corrupt_chapter_file_is_reported_instead_of_silently_ignored(self):
        (profile_store.PROFILES_DIR / "dokumen-2d" / "chapters.json").write_text('{"daftar": 3}', encoding="utf-8")
        response = self.client.get("/api/profiles/dokumen-2d")
        self.assertEqual(response.status_code, 400)
        self.assertIn("Berkas struktur bab profil rusak", response.json["error"])

    def test_invalid_settings_return_clear_400(self):
        response = self.client.put("/api/profiles/dokumen-2d/settings", json={"penomoran": [{**RULE, "mulai": "1"}]})
        self.assertEqual(response.status_code, 400)
        self.assertIn("nomor awal", response.json["error"])
        self.assertEqual(self.client.put("/api/profiles/dokumen-2d/settings", data="bukan json").status_code, 400)


if __name__ == "__main__":
    unittest.main()
