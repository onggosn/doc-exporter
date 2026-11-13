import io
import unittest

from docx import Document

import doc_composer
import list_review
from docx_builder import build_docx

PAGE = {"lebar_pt": 595.28, "tinggi_pt": 841.89, "margin": {"kiri": 34, "kanan": 34, "atas": 60, "bawah": 58},
        "nomor_halaman": False}


def review(lines):
    blocks = doc_composer.verbatim_blocks(lines)
    return blocks, list_review.review_blocks(lines, blocks)


def kinds(findings):
    return [f["jenis"] for f in findings]


class ParseMarkerTests(unittest.TestCase):
    def test_styles_and_values(self):
        cases = {"1.": ("angka", 1), "12)": ("angka", 12), "(3)": ("angka", 3), "a.": ("abjad-kecil", 1),
                 "B)": ("abjad-besar", 2), "ii.": ("romawi-kecil", 2), "IV.": ("romawi-besar", 4),
                 "C)": ("abjad-besar", 3), "d.": ("abjad-kecil", 4), "i.": ("romawi-kecil", 1)}
        for marker, (style, value) in cases.items():
            parsed = list_review.parse_marker(marker)
            self.assertEqual((parsed["gaya"], parsed["nilai"]), (style, value), marker)

    def test_c_d_l_m_after_c_are_read_as_letters_so_a_b_c_d_is_one_sequence(self):
        self.assertEqual(review(["c. A", "d. B", "e. C"])[1], [])

    def test_ambiguous_single_letter_follows_the_previous_style(self):
        self.assertEqual(list_review.parse_marker("i.", "abjad-kecil")["nilai"], 9)  # a..h, lalu i
        self.assertEqual(list_review.parse_marker("i.", "romawi-kecil")["nilai"], 1)
        self.assertEqual(list_review.parse_marker("C)", "abjad-besar")["nilai"], 3)

    def test_unknown_marker(self):
        self.assertIsNone(list_review.parse_marker("--"))
        self.assertIsNone(list_review.parse_marker("1"))


class ReviewBlocksTests(unittest.TestCase):
    def test_clean_sequential_list_has_no_findings_and_no_comments(self):
        blocks, findings = review(["1. Satu", "2. Dua", "  a. Anak", "  b. Anak lain", "3. Tiga", "Butir biasa"])
        self.assertEqual(findings, [])
        self.assertFalse(any("comments" in e for b in blocks if b["type"] == "numbered" for e in b["items"]))

    def test_slide_without_bullets_or_numbers_is_not_reviewed(self):
        self.assertEqual(review(["JUDUL LABEL"])[1], [])
        self.assertEqual(list_review.review_blocks([], []), [])

    def test_restart_after_an_unnumbered_bullet_is_flagged_on_the_restarting_item(self):
        lines = ["a. Memperkuat strategi", "Asumsi intervensi baseline", "Asumsi intervensi PR", "a. Menetapkan Term Structure"]
        blocks, findings = review(lines)
        self.assertEqual(kinds(findings), ["urutan"])
        self.assertIn("mulai lagi", findings[0]["pesan"])
        self.assertIn("disela butir tanpa nomor", findings[0]["pesan"])
        restarted = blocks[-1]["items"][0]
        self.assertEqual(restarted["comments"][0][:19], "[Cek format: urutan")

    def test_duplicate_skipped_and_backwards_numbers(self):
        _, findings = review(["1. A", "1. B"])
        self.assertIn("Nomor ganda", findings[0]["pesan"])
        _, findings = review(["1. A", "3. B"])
        self.assertIn("tidak berurutan", findings[0]["pesan"])
        _, findings = review(["3. A", "2. B"])
        self.assertIn("tidak berurutan", findings[-1]["pesan"])

    def test_continued_list_may_start_from_a_later_marker_or_a_deeper_level(self):
        self.assertEqual(review(["5. Lima", "6. Enam"])[1], [])
        self.assertEqual(review(["  b. Dua", "  c. Tiga"])[1], [])

    def test_marker_style_change_at_the_same_depth_is_a_format_finding(self):
        _, findings = review(["1. Satu", "a. Huruf"])
        self.assertEqual(kinds(findings), ["format"])
        _, findings = review(["1. Satu", "2) Dua"])
        self.assertEqual(kinds(findings), ["format"])

    def test_roman_after_a_non_consecutive_letter_is_read_as_roman_not_as_a_style_change(self):
        _, findings = review(["a. Satu", "i. Romawi", "ii. Romawi dua", "iii. Romawi tiga"])
        self.assertEqual(kinds(findings), ["format"])  # hanya perpindahan a. -> i., bukan i. -> ii.
        self.assertIn("\"a.\" menjadi \"i.\"", findings[0]["pesan"])
        _, findings = review(["a. A", "b. B", "c. C", "d. D", "e. E", "f. F", "g. G", "h. H", "i. I"])
        self.assertEqual(findings, [])  # i sesudah h tetap abjad

    def test_list_continuing_from_i_to_j_is_alphabetic(self):
        self.assertEqual(review(["i. Sembilan", "j. Sepuluh", "k. Sebelas"])[1], [])
        self.assertEqual(review(["i. Satu", "ii. Dua", "iii. Tiga"])[1], [])

    def test_restart_to_one_without_interruption_is_a_normal_new_list(self):
        _, findings = review(["1. A", "2. B", "3. C", "1. D"])
        self.assertEqual(findings, [])

    def test_deeper_level_counters_restart_under_a_new_parent(self):
        _, findings = review(["1. Satu", "  a. x", "  b. y", "2. Dua", "  a. z"])
        self.assertEqual(findings, [])

    def test_depth_jump_and_deep_start(self):
        _, findings = review(["1. Satu", "    a. terlalu dalam"])
        self.assertTrue(any("melompat" in f["pesan"] for f in findings))

    def test_depth_beyond_document_limit_is_reported_as_flattened(self):
        _, findings = review(["Satu", "  Dua", "    Tiga", "      Empat", "        Lima"])
        self.assertEqual(kinds(findings), ["format", "format"])
        self.assertIn("diratakan", findings[0]["pesan"])

    def test_content_mismatch_between_slide_and_blocks_is_reported(self):
        lines = ["1. Satu", "2. Dua"]
        blocks = doc_composer.verbatim_blocks(lines)
        blocks[0]["items"][1]["text"] = "Dua yang berubah"
        findings = list_review.review_blocks(lines, blocks)
        self.assertEqual(kinds(findings), ["isi"])
        self.assertIn("Baris ke-2", findings[0]["pesan"])
        blocks = doc_composer.verbatim_blocks(lines)
        blocks[0]["items"].pop()
        self.assertIn("Jumlah baris berbeda", list_review.review_blocks(lines, blocks)[0]["pesan"])

    def test_label_resets_numbering_between_sections(self):
        _, findings = review(["1. Satu", "2. Dua", "BAGIAN BARU", "1. Satu lagi"])
        self.assertEqual(findings, [])


class ComposeTests(unittest.TestCase):
    SETTINGS = {"penomoran": [], "caption_default": "kosong"}

    def compose(self, chapter, lines, review_lists, verbatim=None):
        slide = {"ref": "slide-1", "title": "Arahan", "text": lines, "text_original": lines, "tables": [], "charts": []}
        if verbatim:
            slide["verbatim"] = verbatim
        return doc_composer.compose({"items": [slide]}, [chapter], {chapter: ["slide-1"]}, {}, self.SETTINGS, [],
                                    review_lists=review_lists)

    def test_verbatim_chapter_lists_are_reviewed_and_reported_when_enabled(self):
        lines = ["a. Satu", "Butir", "a. Dua"]
        spec, report = self.compose("1. Kesimpulan dan Arahan", lines, True)
        self.assertEqual([(f["ref"], f["jenis"]) for f in report["tinjauan_format"]], [("slide-1", "urutan")])
        numbered = [b for b in spec["blocks"] if b["type"] == "numbered"]
        self.assertIn("comments", numbered[1]["items"][0])

    def test_review_is_off_by_default_and_leaves_blocks_untouched(self):
        spec, report = self.compose("1. Kesimpulan dan Arahan", ["a. Satu", "Butir", "a. Dua"], False)
        self.assertEqual(report["tinjauan_format"], [])
        self.assertFalse(any("comments" in e for b in spec["blocks"] if b["type"] == "numbered" for e in b["items"]))

    def test_narrated_slide_with_verbatim_numbered_text_is_reviewed_from_its_verbatim_blocks(self):
        verbatim = [{"label": "", "lines": ["1. Satu", "3. Tiga"]}]
        _, report = self.compose("Bab biasa", ["x"], True, verbatim=verbatim)
        self.assertEqual([f["jenis"] for f in report["tinjauan_format"]], ["urutan"])

    def test_slide_without_lists_is_not_reviewed(self):
        _, report = self.compose("1. Kesimpulan dan Arahan", ["JUDUL BESAR"], True)
        self.assertEqual(report["tinjauan_format"], [])


class BuilderTests(unittest.TestCase):
    def test_comments_on_numbered_bullets_and_labels_become_word_comments_by_the_format_reviewer(self):
        lines = ["LABEL", "a. Satu", "Butir", "a. Lagi"]
        blocks = doc_composer.verbatim_blocks(lines)
        blocks[0]["comments"] = ["[Cek format: format] pada label"]
        findings = list_review.review_blocks(lines, blocks)
        self.assertEqual(kinds(findings), ["urutan"])
        bullets = next(b for b in blocks if b["type"] == "bullets")
        bullets["item_comments"] = {0: ["[Cek format: format] pada butir"]}
        buffer = io.BytesIO()
        build_docx({"blocks": blocks}, PAGE, buffer)
        document = Document(io.BytesIO(buffer.getvalue()))
        comments = list(document.comments)
        self.assertEqual(sorted(c.text.split("] ")[1] for c in comments), ["Penomoran mulai lagi dari \"a.\" setelah \"a.\" karena disela butir tanpa nomor; cocokkan dengan slide apakah seharusnya berlanjut.", "pada butir", "pada label"])
        self.assertEqual({c.author for c in comments}, {"Peninjau Format"})
        self.assertEqual([p.text for p in document.paragraphs][:2], ["LABEL", "a.\tSatu"])  # teks tidak berubah


if __name__ == "__main__":
    unittest.main()
