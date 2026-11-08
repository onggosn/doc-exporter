import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import profile_store
from profile_store import ProfileError, ProfileNotFound


def caption(section, awalan, nomor, source="a.html", text=None):
    return {
        "source_file": source, "section": section, "layout_id": "x", "type": "headline",
        "text": text or f"{awalan}{nomor}. Pesan utama",
        "penomoran": {"awalan": awalan, "nomor": nomor, "pemisah": "."},
    }


CAPTIONS = [
    caption("2. Pembahasan > 2.1 Kesimpulan Moneter", "K", 1),
    caption("2. Pembahasan > 2.1 Kesimpulan Moneter", "K", 2),
    caption("2. Pembahasan > 2.2 Lampiran Data", "Grafik", 1),
    caption("2. Pembahasan > 2.3 Lampiran Data Lanjutan", "Grafik", 2),
    {"source_file": "a.html", "section": "2. Pembahasan > 2.1 Kesimpulan Moneter",
     "layout_id": "x", "type": "kosong", "text": ""},
]


class ProfileStoreTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.root = Path(self._tmp.name)
        patcher = patch.object(profile_store, "PROFILES_DIR", self.root)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.make_profile("dokumen-2d", "Dokumen 2d")

    def make_profile(self, name, title):
        path = self.root / name
        path.mkdir()
        files = {
            "meta": {"title": title, "created_at": "2026-01-01", "sources": [{"file": "a.html"}]},
            "layouts": [{"id": "l1"}, {"id": "l2"}],
            "row_patterns": [], "sections": [], "numbering": [], "captions": CAPTIONS,
        }
        for key, value in files.items():
            (path / f"{key}.json").write_text(json.dumps(value), encoding="utf-8")

    def rule(self, **overrides):
        base = {"bagian_cocok": "Kesimpulan", "awalan": "K", "mulai": 1, "lanjut": "per_bagian", "aktif": True}
        return {**base, **overrides}

    def test_list_profiles_reports_title_and_layout_count(self):
        (self.root / "bukan-profil").mkdir()
        self.assertEqual(
            [(p["name"], p["title"], p["layouts"]) for p in profile_store.list_profiles()],
            [("dokumen-2d", "Dokumen 2d", 2)],
        )

    def test_list_profiles_without_directory_is_empty(self):
        with patch.object(profile_store, "PROFILES_DIR", self.root / "tidak-ada"):
            self.assertEqual(profile_store.list_profiles(), [])

    def test_rejects_path_traversal_and_unknown_names(self):
        for name in ("../dokumen-2d", "Dokumen", "", "a/b"):
            with self.assertRaises(ProfileError):
                profile_store.load_profile(name)
        with self.assertRaises(ProfileNotFound):
            profile_store.load_profile("tidak-ada")

    def test_page_properties_are_optional(self):
        self.assertIsNone(profile_store.load_profile("dokumen-2d")["page"])
        (self.root / "dokumen-2d" / "page.json").write_text(json.dumps({"pdf": []}), encoding="utf-8")
        self.assertEqual(profile_store.load_profile("dokumen-2d")["page"], {"pdf": []})

    def add_style_files(self, samples=None):
        path = self.root / "dokumen-2d"
        (path / "style_samples.json").write_text(
            json.dumps(samples or [{"source": "a", "kind": "paragraf", "text": "Contoh teks sampel."}]),
            encoding="utf-8",
        )
        (path / "style_stats.json").write_text(json.dumps({"blok": {"paragraf": 1}}), encoding="utf-8")

    def test_style_summary_without_style_files_is_empty(self):
        gaya = profile_store.load_profile("dokumen-2d")["gaya"]
        self.assertEqual((gaya["statistik"], gaya["sampel"], gaya["panduan"], gaya["usang"]), (None, None, None, False))
        with self.assertRaisesRegex(ProfileError, "sampel gaya"):
            profile_store.style_inputs("dokumen-2d")

    def test_style_guide_is_saved_and_marked_stale_when_samples_change(self):
        self.add_style_files()
        samples, statistics = profile_store.style_inputs("dokumen-2d")
        self.assertEqual((len(samples), statistics), (1, {"blok": {"paragraf": 1}}))

        profile_store.save_style_guide("dokumen-2d", {"nada": "formal"}, "model-x", 2)
        gaya = profile_store.load_profile("dokumen-2d")["gaya"]
        self.assertEqual((gaya["panduan"]["panduan"], gaya["panduan"]["contoh_dibuang"]), ({"nada": "formal"}, 2))
        self.assertEqual((gaya["sampel"]["blok"], gaya["usang"]), (1, False))
        self.assertFalse((self.root / "dokumen-2d" / "style_guide.json.tmp").exists())

        self.add_style_files([{"source": "a", "kind": "poin", "text": "Sampel yang sudah berbeda."}])
        self.assertTrue(profile_store.load_profile("dokumen-2d")["gaya"]["usang"])

    def test_default_settings_come_from_reference_numbering(self):
        profile = profile_store.load_profile("dokumen-2d")
        self.assertFalse(profile["settings_saved"])
        rules = {r["awalan"]: r for r in profile["settings"]["penomoran"]}
        self.assertEqual(rules["K"]["bagian_cocok"], "Kesimpulan Moneter")
        self.assertEqual((rules["K"]["mulai"], rules["K"]["lanjut"]), (1, "per_bagian"))

    def test_numbering_that_continues_across_sections_defaults_to_berlanjut(self):
        rules = {r["awalan"]: r for r in profile_store.default_settings(CAPTIONS)["penomoran"]}
        self.assertEqual(rules["Grafik"]["lanjut"], "berlanjut")

    def test_evidence_groups_by_section_keyword_and_prefix(self):
        evidence = {(e["bagian"], e["awalan"]): e for e in profile_store.numbering_evidence(CAPTIONS)}
        self.assertEqual(evidence[("Kesimpulan Moneter", "K")]["jumlah"], 2)
        self.assertEqual(evidence[("Kesimpulan Moneter", "K")]["terakhir"], 2)

    def test_saved_settings_are_loaded_and_survive_reset_only_until_reset(self):
        saved = profile_store.save_settings("dokumen-2d", {"penomoran": [self.rule(awalan=" P ", mulai=3)]})
        self.assertEqual(saved["penomoran"][0]["awalan"], "P")
        loaded = profile_store.load_profile("dokumen-2d")
        self.assertTrue(loaded["settings_saved"])
        self.assertEqual(loaded["settings"], saved)
        self.assertFalse((self.root / "dokumen-2d" / "settings.json.tmp").exists())

        defaults = profile_store.reset_settings("dokumen-2d")
        self.assertFalse(profile_store.load_profile("dokumen-2d")["settings_saved"])
        self.assertEqual(defaults, profile_store.load_profile("dokumen-2d")["settings"])

    def test_validation_rejects_bad_rules(self):
        bad_rules = [
            self.rule(bagian_cocok="  "),
            self.rule(awalan=""),
            self.rule(awalan="X" * 21),
            self.rule(mulai=True),
            self.rule(mulai=-1),
            self.rule(mulai="1"),
            self.rule(lanjut="acak"),
            self.rule(aktif="ya"),
            "bukan objek",
        ]
        for rule in bad_rules:
            with self.assertRaises(ProfileError, msg=repr(rule)):
                profile_store.validate_settings({"penomoran": [rule]})
        for data in (None, [], {"penomoran": "x"}, {"penomoran": [self.rule()] * 51}):
            with self.assertRaises(ProfileError):
                profile_store.validate_settings(data)

    def test_caption_default_is_validated_defaulted_and_saved(self):
        self.assertEqual(profile_store.load_profile("dokumen-2d")["settings"]["caption_default"], "kosong")
        self.assertEqual(profile_store.validate_settings({"penomoran": []})["caption_default"], "kosong")
        saved = profile_store.save_settings("dokumen-2d", {"penomoran": [], "caption_default": "headline"})
        self.assertEqual(saved["caption_default"], "headline")
        self.assertEqual(profile_store.load_profile("dokumen-2d")["settings"]["caption_default"], "headline")
        self.assertEqual(profile_store.reset_settings("dokumen-2d")["caption_default"], "kosong")
        for bad in ("acak", None, 1):
            with self.assertRaisesRegex(ProfileError, "Caption bawaan"):
                profile_store.validate_settings({"penomoran": [], "caption_default": bad})

    def test_corrupt_settings_file_gives_clear_error(self):
        (self.root / "dokumen-2d" / "settings.json").write_text("{rusak", encoding="utf-8")
        with self.assertRaisesRegex(ProfileError, "rusak"):
            profile_store.load_profile("dokumen-2d")


if __name__ == "__main__":
    unittest.main()
