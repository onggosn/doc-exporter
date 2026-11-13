import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from pptx import Presentation
from pptx.util import Inches

import app
import doc_composer
import profile_store
from ppt_adapter import detect_columns
from profile_store import ProfileError, validate_default_section

CONFIG = {
    "bagian_cocok": "Default Section", "judul_bab": "1. Kesimpulan Diskusi Hari 1 dan Arahan ADG",
    "judul_kolom": "1.{n} Kesimpulan dan Arahan {nama}",
    "kolom": [{"cocok": "MONETER", "nama": "Moneter"}, {"cocok": "MAKROPRUDENSIAL", "nama": "Makroprudensial"},
              {"cocok": "SP", "nama": "Sistem Pembayaran"}],
}


class DetectColumnsTests(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(10), Inches(7.5)
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def text(self, text, left, top, width=3.0, height=3.0):
        self.slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height)).text_frame.text = text

    def columns(self):
        return detect_columns(self.slide, self.prs.slide_width, self.prs.slide_height)

    def three_columns(self):
        for header, left in (("MONETER", 0.5), ("MAKROPRUDENSIAL", 3.7), ("SP PUR", 7.0)):
            self.text(header, left, 0.4, width=2.0, height=0.4)
        # urutan shape sengaja acak: kolom ditentukan posisi, bukan urutan shape
        self.text("KESIMPULAN\nK#1: kolom kanan", 6.8, 1.5)
        self.text("KESIMPULAN\nK1: kolom kiri\nPROSPEK", 0.2, 1.5)
        self.text("KESIMPULAN\nK1: kolom tengah", 3.5, 1.5)

    def test_columns_are_split_by_position_with_header_names(self):
        self.three_columns()
        columns = self.columns()
        self.assertEqual([c["header"] for c in columns], ["MONETER", "MAKROPRUDENSIAL", "SP PUR"])
        self.assertEqual(columns[0]["text"], ["KESIMPULAN", "K1: kolom kiri", "PROSPEK"])
        self.assertEqual(columns[1]["text"], ["KESIMPULAN", "K1: kolom tengah"])
        self.assertEqual(columns[2]["text"], ["KESIMPULAN", "K#1: kolom kanan"])

    def test_slide_title_is_not_part_of_any_column(self):
        self.three_columns()
        title = self.prs.slides.add_slide(self.prs.slide_layouts[5])
        title.shapes.title.text = "Kesimpulan dan Arahan"
        self.assertEqual(len(detect_columns(title, self.prs.slide_width, self.prs.slide_height)), 0)

    def test_one_header_or_no_headers_is_not_a_column_layout(self):
        self.text("MONETER", 0.5, 0.4, width=2.0, height=0.4)
        self.text("isi", 0.2, 1.5)
        self.assertEqual(self.columns(), [])

    def test_text_box_spanning_the_columns_keeps_the_slide_whole(self):
        self.three_columns()
        self.text("Catatan melintang di seluruh lebar slide yang tidak boleh hilang", 0.2, 5.5, width=9.5, height=1)
        self.assertEqual(self.columns(), [])

    def test_table_keeps_the_slide_whole(self):
        self.three_columns()
        self.slide.shapes.add_table(2, 2, Inches(1), Inches(5), Inches(3), Inches(1))
        self.assertEqual(self.columns(), [])

    def test_column_without_text_keeps_the_slide_whole(self):
        for header, left in (("MONETER", 0.5), ("SP PUR", 7.0)):
            self.text(header, left, 0.4, width=2.0, height=0.4)
        self.text("hanya kolom kiri", 0.2, 1.5)
        self.assertEqual(self.columns(), [])

    def test_overlapping_headers_are_not_columns(self):
        self.text("MONETER", 0.5, 0.4, width=2.0, height=0.4)
        self.text("SP PUR", 1.5, 0.4, width=2.0, height=0.4)
        self.text("isi", 0.2, 1.5)
        self.assertEqual(self.columns(), [])


class ProfileConfigTests(unittest.TestCase):
    def test_valid_config_is_normalized(self):
        cleaned = validate_default_section({**CONFIG, "judul_kolom": None, "judul_bab": None})
        self.assertEqual(cleaned["judul_kolom"], "{nama}")
        self.assertEqual(cleaned["judul_bab"], "")

    def test_invalid_configs_are_rejected(self):
        for bad in ([], {**CONFIG, "kolom": []}, {**CONFIG, "kolom": [{"cocok": "A"}]},
                    {**CONFIG, "bagian_cocok": " "}, {k: v for k, v in CONFIG.items() if k != "kolom"}):
            with self.assertRaises(ProfileError):
                validate_default_section(bad)

    def test_missing_or_broken_file_means_no_config(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "p").mkdir()
            (root / "p" / "meta.json").write_text("{}", encoding="utf-8")  # penanda profil ada
            with patch.object(profile_store, "PROFILES_DIR", root):
                self.assertIsNone(profile_store.load_default_section("p"))
                (root / "p" / "default_section.json").write_text("{rusak", encoding="utf-8")
                self.assertIsNone(profile_store.load_default_section("p"))
                (root / "p" / "default_section.json").write_text(json.dumps(CONFIG), encoding="utf-8")
                self.assertEqual(profile_store.load_default_section("p")["kolom"][2]["nama"], "Sistem Pembayaran")

    def test_shipped_profile_has_the_default_section_config(self):
        config = profile_store.load_default_section("dokumen-2d")
        self.assertEqual([k["nama"] for k in config["kolom"]], ["Moneter", "Makroprudensial", "Sistem Pembayaran"])


def column_slide(headers=("MONETER", "MAKROPRUDENSIAL", "SP PUR")):
    return {"ref": "slide-2", "title": "Kesimpulan dan Arahan ADG – RDGB Maret 2026 Hari I", "text": ["x"],
            "tables": [], "charts": [], "vision_reasons": [], "needs_vision": True,
            "kolom": [{"header": h, "text": [f"KESIMPULAN {h}", f"K1: isi {h}"]} for h in headers]}


def deck_of(slides, section="Default Section"):
    return {"items": slides, "sections": [{"nama": section, "refs": [s["ref"] for s in slides]}]}


class SplitTests(unittest.TestCase):
    def test_match_column_uses_whole_words_so_sp_pur_matches_sp_but_not_spesial(self):
        cols = CONFIG["kolom"]
        self.assertEqual(doc_composer.match_column("SP PUR", cols), 2)
        self.assertEqual(doc_composer.match_column("sp", cols), 2)
        self.assertEqual(doc_composer.match_column("moneter", cols), 0)
        self.assertIsNone(doc_composer.match_column("SPESIAL", cols))
        self.assertIsNone(doc_composer.match_column("LAINNYA", cols))

    def test_slide_is_split_into_one_titled_verbatim_slide_per_column(self):
        deck = deck_of([column_slide()])
        self.assertEqual(doc_composer.split_default_section(deck, CONFIG), ["slide-2"])
        self.assertEqual([i["ref"] for i in deck["items"]], ["slide-2-k1", "slide-2-k2", "slide-2-k3"])
        self.assertEqual([i["title"] for i in deck["items"]], [
            "1.1 Kesimpulan dan Arahan Moneter", "1.2 Kesimpulan dan Arahan Makroprudensial",
            "1.3 Kesimpulan dan Arahan Sistem Pembayaran"])
        self.assertEqual(deck["sections"][0]["refs"], ["slide-2-k1", "slide-2-k2", "slide-2-k3"])
        self.assertEqual(deck["items"][2]["verbatim"][0]["lines"], ["KESIMPULAN SP PUR", "K1: isi SP PUR"])
        self.assertTrue(deck["section_bawaan_dipakai"])
        self.assertFalse(any(i["needs_vision"] for i in deck["items"]))  # tidak dideskripsikan model

    def test_columns_follow_profile_order_not_slide_order(self):
        deck = deck_of([column_slide(("SP PUR", "MONETER", "MAKROPRUDENSIAL"))])
        doc_composer.split_default_section(deck, CONFIG)
        self.assertEqual([i["title"][:3] for i in deck["items"]], ["1.1", "1.2", "1.3"])

    def test_left_as_is_without_config_without_matching_section_or_with_an_unknown_header(self):
        for config, deck in (
            (None, deck_of([column_slide()])),
            (CONFIG, deck_of([column_slide()], section="Moneter")),
            (CONFIG, deck_of([column_slide(("MONETER", "MAKROPRUDENSIAL", "LAINNYA"))])),
            (CONFIG, deck_of([column_slide(("MONETER", "MONETER"))])),
        ):
            before = [i["ref"] for i in deck["items"]]
            self.assertEqual(doc_composer.split_default_section(deck, config), [])
            self.assertEqual([i["ref"] for i in deck["items"]], before)
            self.assertNotIn("section_bawaan_dipakai", deck)

    def test_slide_chapters_use_the_profile_chapter_name_only_when_columns_were_split(self):
        deck = deck_of([column_slide()])
        self.assertEqual(app._slide_chapters(deck, CONFIG)[0], ["Kesimpulan dan Arahan ADG – RDGB Maret 2026 Hari I"])
        doc_composer.split_default_section(deck, CONFIG)
        chapters, mapping = app._slide_chapters(deck, CONFIG)
        self.assertEqual(chapters, ["1. Kesimpulan Diskusi Hari 1 dan Arahan ADG"])
        self.assertEqual(mapping[chapters[0]], ["slide-2-k1", "slide-2-k2", "slide-2-k3"])

    def test_pin_places_split_slides_into_chapters_with_the_same_number(self):
        deck = deck_of([column_slide()])
        doc_composer.split_default_section(deck, CONFIG)
        chapters = ["1. Kesimpulan Diskusi", "1.1 Kesimpulan dan Arahan Moneter", "1.2 Kesimpulan dan Arahan Makroprudensial",
                    "1.3 Kesimpulan dan Arahan SP", "2. Lain"]
        mapping = {c: [] for c in chapters}
        mapping["2. Lain"] = ["slide-2-k1", "slide-9"]  # salah petak oleh model
        doc_composer.pin_split_slides(chapters, mapping, deck["items"])
        self.assertEqual(mapping["1.1 Kesimpulan dan Arahan Moneter"], ["slide-2-k1"])
        self.assertEqual(mapping["1.3 Kesimpulan dan Arahan SP"], ["slide-2-k3"])
        self.assertEqual(mapping["2. Lain"], ["slide-9"])
        mapping = {"A": []}
        doc_composer.pin_split_slides(["A"], mapping, deck["items"])  # tanpa bab bernomor sama: tidak diubah
        self.assertEqual(mapping, {"A": []})

    def test_compose_gives_h1_chapter_then_h2_per_column_without_repeating_a_title_equal_to_the_chapter(self):
        deck = deck_of([column_slide()])
        doc_composer.split_default_section(deck, CONFIG)
        chapters, mapping = app._slide_chapters(deck, CONFIG)
        spec, _ = doc_composer.compose(deck, chapters, mapping, {}, {"penomoran": [], "caption_default": "kosong"}, [])
        headings = [(b["level"], b["text"]) for b in spec["blocks"] if b["type"] == "heading"]
        self.assertEqual(headings, [(1, "1. Kesimpulan Diskusi Hari 1 dan Arahan ADG"),
                                    (2, "1.1 Kesimpulan dan Arahan Moneter"),
                                    (2, "1.2 Kesimpulan dan Arahan Makroprudensial"),
                                    (2, "1.3 Kesimpulan dan Arahan Sistem Pembayaran")])
        pinned = {"1.1 Kesimpulan dan Arahan Moneter": ["slide-2-k1"]}
        spec, _ = doc_composer.compose(deck, list(pinned), pinned, {}, {"penomoran": [], "caption_default": "kosong"}, [])
        self.assertEqual([b["text"] for b in spec["blocks"] if b["type"] == "heading"], ["1.1 Kesimpulan dan Arahan Moneter"])


if __name__ == "__main__":
    unittest.main()
