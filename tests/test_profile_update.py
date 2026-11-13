import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import profile_builder as pb
import profile_store


def html_doc(period_tag: str, children=("Moneter", "Makroprudensial", "Sistem Pembayaran"), prefix="Kesimpulan dan Arahan",
             grafik_start=1):
    kids = "".join(
        f'<div class=\\"title-level-2\\"><p><strong>1.{i} {prefix} {name}</strong></p></div>'
        for i, name in enumerate(children, start=1))
    return (
        r'<div class=\"title-level-1\"><p><strong>1. Kesimpulan Diskusi dan Arahan ADG</strong></p></div>' + kids +
        r'<div class=\"title-level-1\"><p><strong>2. Ringkasan Asesmen</strong></p></div>'
        r'<div class=\"title-level-2\"><p><strong>2.1 Asesmen Moneter</strong></p></div>'
        r'<p>Inflasi tetap terkendali dalam kisaran sasaran sepanjang periode laporan ini berlangsung, ' + period_tag +
        r' dan nilai tukar relatif stabil.</p>'
        r'<ul><li>Poin pertama yang cukup panjang untuk menjadi sampel gaya penulisan referensi dokumen.</li></ul>'
        r'<figure class=\"image\"><img src=\"a.png\"><figcaption>Grafik ' + str(grafik_start) + r'. Kredit UMKM</figcaption></figure>'
    )


class ProfileTestBase(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        self.root = self.tmp / "profiles"
        self.refs = self.tmp / "ref"
        self.refs.mkdir()

    def ref(self, name, content):
        path = self.refs / name
        path.write_text(content, encoding="utf-8")
        return path

    def run_update(self, htmls, **kwargs):
        return pb.update_profile("dokumen-x", htmls, kwargs.pop("pdfs", []), root=self.root, **kwargs)

    def files(self):
        return sorted(p.name for p in (self.root / "dokumen-x").iterdir())


class UpdateProfileTests(ProfileTestBase):
    def test_unknown_type_is_created_and_known_type_is_updated(self):
        self.assertFalse(pb.document_type_exists("dokumen-x", self.root))
        self.assertEqual(pb.existing_document_types(self.root), [])
        first = self.ref("DKEM Maret 2026.html", html_doc("Maret"))
        report = self.run_update([first], title="Dokumen X")
        self.assertEqual((report["mode"], report["ditambah"]), ("baru", ["DKEM Maret 2026.html"]))
        self.assertTrue(pb.document_type_exists("dokumen-x", self.root))
        self.assertEqual(pb.existing_document_types(self.root), ["dokumen-x"])
        second = self.ref("DKEM Mei 2025.html", html_doc("Mei", grafik_start=3))
        report = self.run_update([second])
        self.assertEqual((report["mode"], report["ditambah"], report["dipertahankan"]),
                         ("update", ["DKEM Mei 2025.html"], ["DKEM Maret 2026.html"]))
        meta = json.loads((self.root / "dokumen-x" / "meta.json").read_text())
        self.assertEqual([s["file"] for s in meta["sources"]], ["DKEM Maret 2026.html", "DKEM Mei 2025.html"])  # urut nama berkas
        self.assertEqual(meta["title"], "Dokumen X")  # judul lama dipertahankan
        self.assertEqual(len(meta["riwayat"]), 2)

    def test_invalid_type_name_is_rejected_with_the_list_of_existing_types(self):
        self.run_update([self.ref("a.html", html_doc("x"))])
        with self.assertRaisesRegex(pb.ProfileBuildError, "dokumen-x"):
            pb.update_profile("Dokumen 2D!", [], [], root=self.root)

    def test_update_equals_building_from_all_documents_at_once(self):
        a, b = self.ref("DKEM Maret 2026.html", html_doc("Maret")), self.ref("DKEM Mei 2025.html", html_doc("Mei", grafik_start=3))
        self.run_update([a])
        self.run_update([b])
        incremental = {f: json.loads((self.root / "dokumen-x" / f"{f}.json").read_text()) for f in pb.PROFILE_FILES}
        scratch = pb.assemble_profile(
            sorted([pb.analyze_html_source(a), pb.analyze_html_source(b)], key=lambda s: s["file"]), [], "Dokumen X")
        for key in pb.PROFILE_FILES:
            if key != "meta":
                self.assertEqual(incremental[key], scratch[key], key)

    def test_unchanged_documents_are_skipped_and_nothing_is_rewritten(self):
        a = self.ref("DKEM Maret 2026.html", html_doc("Maret"))
        self.run_update([a])
        stamp = (self.root / "dokumen-x" / "layouts.json").stat().st_mtime_ns
        report = self.run_update([a])
        self.assertEqual((report["tetap"], report["ada_perubahan"]), (["DKEM Maret 2026.html"], False))
        self.assertEqual((self.root / "dokumen-x" / "layouts.json").stat().st_mtime_ns, stamp)
        self.assertEqual(self.run_update([a], force=True)["tetap"], ["DKEM Maret 2026.html"])

    def test_changed_document_with_the_same_name_replaces_the_old_analysis(self):
        a = self.ref("DKEM Maret 2026.html", html_doc("Maret"))
        self.run_update([a])
        a.write_text(html_doc("Maret", grafik_start=7), encoding="utf-8")
        report = self.run_update([a])
        self.assertEqual(report["diganti"], ["DKEM Maret 2026.html"])
        numbering = json.loads((self.root / "dokumen-x" / "numbering.json").read_text())
        self.assertEqual(numbering[0]["first"], 7)
        self.assertEqual(len(list((self.root / "dokumen-x" / "sumber").glob("*.json"))), 1)

    def test_remove_and_replace_all(self):
        a, b = self.ref("DKEM Maret 2026.html", html_doc("Maret")), self.ref("DKEM Mei 2025.html", html_doc("Mei"))
        self.run_update([a, b])
        report = self.run_update([], remove=["DKEM Mei 2025.html"])
        self.assertEqual(report["dihapus"], ["DKEM Mei 2025.html"])
        meta = json.loads((self.root / "dokumen-x" / "meta.json").read_text())
        self.assertEqual([s["file"] for s in meta["sources"]], ["DKEM Maret 2026.html"])
        with self.assertRaisesRegex(pb.ProfileBuildError, "tidak ada di profil"):
            self.run_update([], remove=["tidak-ada.html"])
        report = self.run_update([b], replace_all=True)
        self.assertEqual(report["dihapus"], ["DKEM Maret 2026.html"])
        with self.assertRaisesRegex(pb.ProfileBuildError, "sedikitnya satu"):
            self.run_update([], remove=["DKEM Mei 2025.html"])

    def test_legacy_profile_requires_its_documents_once_and_loses_nothing(self):
        a, b = self.ref("DKEM Maret 2026.html", html_doc("Maret")), self.ref("DKEM Mei 2025.html", html_doc("Mei"))
        self.run_update([a, b])
        for cache in (self.root / "dokumen-x" / "sumber").glob("*.json"):  # meniru profil lama tanpa cache sumber
            cache.unlink()
        c = self.ref("DKEM Juni 2026.html", html_doc("Juni"))
        with self.assertRaisesRegex(pb.ProfileBuildError, "DKEM Maret 2026.html, DKEM Mei 2025.html"):
            self.run_update([c])
        before = (self.root / "dokumen-x" / "layouts.json").read_text()
        report = self.run_update([a, b])  # hanya mencatat cache; isi profil tidak berubah
        self.assertEqual(sorted(report["dicatat_ke_cache"]), ["DKEM Maret 2026.html", "DKEM Mei 2025.html"])
        self.assertEqual((self.root / "dokumen-x" / "layouts.json").read_text(), before)
        report = self.run_update([c])
        self.assertEqual((report["ditambah"], sorted(report["dipertahankan"])), (["DKEM Juni 2026.html"], ["DKEM Maret 2026.html", "DKEM Mei 2025.html"]))

    def test_default_section_and_settings_files_made_by_hand_are_not_overwritten(self):
        a = self.ref("DKEM Maret 2026.html", html_doc("Maret"))
        self.run_update([a])
        directory = self.root / "dokumen-x"
        manual = {"bagian_cocok": "Default Section", "judul_bab": "Buatan saya", "kolom": [{"cocok": "X", "nama": "X"}]}
        (directory / "default_section.json").write_text(json.dumps(manual), encoding="utf-8")  # tanpa dibuat_oleh
        (directory / "settings.json").write_text(json.dumps(
            {"penomoran": [{"bagian_cocok": "Lain", "awalan": "Z", "mulai": 3, "lanjut": "per_bagian", "aktif": True}],
             "caption_default": "headline"}), encoding="utf-8")
        report = self.run_update([self.ref("DKEM Mei 2025.html", html_doc("Mei", grafik_start=2))])
        self.assertIn("dibiarkan", report["berkas"]["default_section.json"])
        self.assertEqual(json.loads((directory / "default_section.json").read_text())["judul_bab"], "Buatan saya")
        rules = json.loads((directory / "settings.json").read_text())["penomoran"]
        self.assertEqual(rules[0]["awalan"], "Z")  # aturan lama tetap di depan
        self.assertTrue(any(r["awalan"] == "Grafik" for r in rules[1:]))  # aturan baru dari referensi ditambah
        self.assertIn("ditambahkan", report["berkas"]["settings.json"])

    def test_generated_default_section_is_refreshed_on_update(self):
        self.run_update([self.ref("DKEM Maret 2026.html", html_doc("Maret", children=("Moneter", "Makroprudensial")))])
        path = self.root / "dokumen-x" / "default_section.json"
        self.assertEqual(len(json.loads(path.read_text())["kolom"]), 2)
        report = self.run_update([self.ref("DKEM Juni 2026.html", html_doc("Juni"))])  # periode terbaru: tiga kolom
        self.assertEqual(report["berkas"]["default_section.json"], "diperbarui")
        self.assertEqual(len(json.loads(path.read_text())["kolom"]), 3)

    def test_dry_run_writes_nothing(self):
        report = self.run_update([self.ref("a.html", html_doc("x"))], dry_run=True)
        self.assertTrue(report["dry_run"])
        self.assertFalse(pb.document_type_exists("dokumen-x", self.root))

    def test_all_profile_files_are_produced_by_the_script_except_the_model_made_style_guide(self):
        self.run_update([self.ref("DKEM Maret 2026.html", html_doc("Maret"))])
        produced = set(self.files())
        for expected in ("meta.json", "layouts.json", "row_patterns.json", "sections.json", "captions.json",
                         "numbering.json", "style_stats.json", "style_samples.json", "chapters.json",
                         "default_section.json", "sumber"):
            self.assertIn(expected, produced)
        self.assertNotIn("style_guide.json", produced)  # butuh model: hanya lewat --panduan-gaya atau tombol di menu
        self.assertNotIn("settings.json", produced)  # diturunkan otomatis dari caption saat profil dimuat
        with patch.object(profile_store, "PROFILES_DIR", self.root):
            loaded = profile_store.load_profile("dokumen-x")
        self.assertEqual(loaded["default_section"]["judul_kolom"], "1.{n} Kesimpulan dan Arahan {nama}")


class CommandLineAndStyleGuideTests(ProfileTestBase):
    def cli(self, *args):
        import contextlib
        import io
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            code = pb.main(["--name", "dokumen-x", "--out", str(self.root), *map(str, args)])
        return code, buffer.getvalue()

    def test_cli_creates_then_updates_and_reports_what_changed(self):
        a = self.ref("DKEM Maret 2026.html", html_doc("Maret"))
        code, out = self.cli(a, "--title", "Dokumen X")
        self.assertEqual(code, 0)
        self.assertIn("belum ada", out)
        self.assertIn("TIPE DOKUMEN BARU DIBUAT: dokumen-x", out)
        b = self.ref("DKEM Mei 2025.html", html_doc("Mei", grafik_start=3))
        code, out = self.cli(b)
        self.assertIn("Tipe dokumen yang sudah ada", out)
        self.assertIn("dokumen-x", out.splitlines()[0])
        self.assertIn("TIPE DOKUMEN SUDAH ADA, DIPERBARUI", out)
        self.assertIn("dokumen ditambah: DKEM Mei 2025.html", out)
        self.assertIn("style_guide.json (butuh model", out)

    def test_cli_check_mode_writes_nothing_and_errors_return_code_1(self):
        a = self.ref("a.html", html_doc("x"))
        code, out = self.cli(a, "--cek")
        self.assertEqual(code, 0)
        self.assertIn("tidak ada berkas yang ditulis", out)
        self.assertFalse(pb.document_type_exists("dokumen-x", self.root))
        code, out = self.cli("--hapus", "a.html")
        self.assertEqual(code, 1)
        self.assertIn("GAGAL", out)

    def test_style_guide_is_only_made_on_request_and_marked_stale_when_samples_change(self):
        a = self.ref("DKEM Maret 2026.html", html_doc("Maret"))
        self.run_update([a])
        calls = []

        class FakeGenerator:
            model_id = "model-uji"

            def __init__(self, **kwargs):
                calls.append(kwargs)

            def generate_style_guide(self, samples, statistics):
                return {"nada": "Formal"}, 1

        with patch("bedrock_service.BedrockDocumentGenerator", FakeGenerator):
            code, out = self.cli(a, "--panduan-gaya", "--paksa", "--model-id", "m", "--region", "r")
        self.assertEqual(code, 0)
        self.assertEqual(calls, [{"model_id": "m", "region": "r"}])
        saved = json.loads((self.root / "dokumen-x" / "style_guide.json").read_text())
        self.assertEqual((saved["panduan"], saved["model_id"], saved["contoh_dibuang"]), ({"nada": "Formal"}, "model-uji", 1))
        self.assertIn("dibuat ulang", out)
        report = self.run_update([self.ref("DKEM Mei 2025.html", html_doc("Mei", grafik_start=3))])
        self.assertIn("usang", report["panduan_gaya"])
        calls.clear()
        self.cli(a)  # tanpa --panduan-gaya model tidak dipanggil
        self.assertEqual(calls, [])


class DerivedFilesTests(unittest.TestCase):
    def source(self, name, children, prefix="Kesimpulan dan Arahan"):
        outline = [{"level": 1, "title": "1. Kesimpulan Diskusi Hari 1 dan Arahan ADG"}]
        outline += [{"level": 2, "title": f"1.{i} {prefix} {c}".replace("  ", " ")} for i, c in enumerate(children, start=1)]
        outline += [{"level": 1, "title": "2. Lain"}]
        return {"file": name, "outline": outline}

    def test_default_section_comes_from_the_latest_reference_outline(self):
        old = self.source("DKEM Mei 2025.html", ["Moneter", "Makroprudensial", "SP"])
        new = self.source("DKEM Maret 2026.html", ["Moneter", "Makroprudensial", "Sistem Pembayaran"])
        config = pb.derive_default_section([old, new])
        self.assertEqual(config["judul_bab"], "1. Kesimpulan Diskusi Hari 1 dan Arahan ADG")
        self.assertEqual(config["judul_kolom"], "1.{n} Kesimpulan dan Arahan {nama}")
        self.assertEqual([(k["cocok"], k["nama"]) for k in config["kolom"]],
                         [("MONETER", "Moneter"), ("MAKROPRUDENSIAL", "Makroprudensial"), ("SP", "Sistem Pembayaran")])
        self.assertEqual(config["dibuat_oleh"], "profile_builder.py")
        profile_store.validate_default_section(config)  # lolos validasi pembaca

    def test_default_section_is_none_without_columns_or_with_clashing_keys(self):
        self.assertIsNone(pb.derive_default_section([{"file": "a.html", "outline": [{"level": 1, "title": "1. Satu"}]}]))
        clash = self.source("a.html", ["Sistem Pembayaran", "Sarana Penunjang"])
        self.assertIsNone(pb.derive_default_section([clash]))

    def test_chapter_structures_are_added_replaced_and_user_made_ones_are_kept(self):
        sources = [self.source("DKEM Maret 2026.html", ["Moneter", "Makro"]), self.source("DKEM Mei 2025.html", ["Moneter", "Makro"])]
        derived = pb.derive_chapter_structures(sources)
        existing = {"bawaan": "Buatan saya", "daftar": [{"nama": "Buatan saya", "bab": ["A"]}, {"nama": "Maret 2026", "bab": ["usang"]}]}
        merged = pb.merge_chapters(existing, derived)
        by_name = {item["nama"]: item["bab"] for item in merged["daftar"]}
        self.assertEqual(by_name["Buatan saya"], ["A"])
        self.assertEqual(by_name["Maret 2026"][0], "1. Kesimpulan Diskusi Hari 1 dan Arahan ADG")
        self.assertIn("Mei 2025", by_name)
        self.assertEqual(merged["bawaan"], "Buatan saya")
        self.assertEqual(pb.merge_chapters(None, derived)["bawaan"], "Maret 2026")  # terbaru
        profile_store.validate_chapters(merged)

    def test_settings_merge_adds_only_new_rules_and_keeps_edited_ones(self):
        def caption(prefix, section):
            return {"type": "label_bernomor", "section": section, "source_file": "a.html", "layout_id": "x",
                    "text": f"{prefix} 1. X", "penomoran": {"awalan": prefix, "nomor": 1}}

        captions = [caption("Grafik", "2. Ringkasan > 2.1 A")]
        derived = profile_store.default_settings(captions)["penomoran"]
        self.assertEqual(len(derived), 1)
        edited = {"penomoran": [{**derived[0], "mulai": 5, "aktif": False}], "caption_default": "headline"}
        merged, added = pb.merge_settings(edited, captions)
        self.assertEqual((merged["penomoran"][0]["mulai"], merged["penomoran"][0]["aktif"], added), (5, False, []))
        merged, added = pb.merge_settings(edited, [*captions, caption("Tabel", "2. Ringkasan > 2.1 A")])
        self.assertEqual(len(merged["penomoran"]), 2)
        self.assertEqual(len(added), 1)
        self.assertEqual(merged["caption_default"], "headline")


if __name__ == "__main__":
    unittest.main()
