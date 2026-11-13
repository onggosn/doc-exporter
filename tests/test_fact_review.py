import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document

import doc_composer
import fact_review
from bedrock_service import BedrockDocumentGenerator, BedrockError
from docx_builder import build_docx

PAGE = {"lebar_pt": 595.28, "tinggi_pt": 841.89, "margin": {"kiri": 34, "kanan": 34, "atas": 60, "bawah": 58},
        "nomor_halaman": False}
FACTS = "Inflasi turun dari 3,2% menjadi 2,8% pada Maret 2026. BI-Rate 4,75%. Proyeksi pertumbuhan 4,85 persen TW II."


def kinds(text, facts=FACTS):
    return [f["jenis"] for f in fact_review.check_text(text, facts)]


class CheckTextTests(unittest.TestCase):
    def test_text_that_matches_the_facts_has_no_findings(self):
        text = "Inflasi turun dari 3,2% menjadi 2,8% pada Maret 2026, sedangkan BI-Rate tetap 4,75%."
        self.assertEqual(fact_review.check_text(text, FACTS), [])

    def test_number_missing_from_facts_is_flagged_with_its_written_form(self):
        findings = fact_review.check_text("Inflasi mencapai 3,9%.", FACTS)
        self.assertEqual([f["jenis"] for f in findings], ["angka"])
        self.assertIn("3,9", findings[0]["pesan"])

    def test_rounded_fact_numbers_and_decimal_style_are_accepted(self):
        self.assertEqual(kinds("Pertumbuhan 4,9% dan inflasi 2.8%."), [])  # 4,85 dibulatkan; titik = koma

    def test_single_digit_counts_are_ignored_unless_percent(self):
        self.assertEqual(kinds("Ada 3 faktor utama."), [])
        self.assertEqual(kinds("Naik 7%."), ["angka"])

    def test_year_month_and_quarter_not_in_facts_are_flagged_as_periods(self):
        findings = fact_review.check_text("Pada Juli 2027, TW IV tumbuh.", FACTS)
        self.assertEqual([f["jenis"] for f in findings], ["periode", "periode", "periode"])
        messages = " ".join(f["pesan"] for f in findings)
        for expected in ("2027", "Juli", "TW IV"):
            self.assertIn(expected, messages)

    def test_period_present_in_facts_is_not_flagged(self):
        self.assertEqual(kinds("Pada Maret 2026 dan TW II."), [])
        self.assertEqual(kinds("Pada Mar 2026.", "Data Mar 2026"), [])  # singkatan bulan di slide

    def test_names_are_not_checked_because_image_summaries_may_name_what_only_the_image_shows(self):
        self.assertEqual(kinds("Kebijakan The Federal Reserve dan OJK dibahas."), [])

    def test_direction_contradicting_from_to_numbers_is_flagged(self):
        findings = fact_review.check_text("Inflasi naik dari 3,2% menjadi 2,8%.", FACTS)
        self.assertEqual([f["jenis"] for f in findings], ["arah"])
        self.assertEqual(kinds("Inflasi turun dari 3,2% menjadi 2,8%."), [])
        self.assertEqual(kinds("Inflasi menurun dari 2,8% menjadi 3,2%."), ["arah"])

    def test_tags_are_ignored_and_empty_text_is_safe(self):
        self.assertEqual(kinds("Inflasi <i>year-on-year</i> turun menjadi 2,8%."), [])
        self.assertEqual(fact_review.check_text("", FACTS), [])


class ParseReviewTests(unittest.TestCase):
    def test_keeps_only_valid_findings(self):
        data = {"temuan": [
            {"paragraf": 1, "kutipan": "tumbuh 9%", "alasan": "Slide menyebut 4,85%."},
            {"paragraf": "2", "alasan": "Periode salah."},
            {"paragraf": 9, "alasan": "di luar rentang"},
            {"paragraf": 1, "alasan": "  "},
            {"paragraf": True, "alasan": "bool bukan nomor"},
            "bukan objek",
        ]}
        result = doc_composer.parse_review(data, 2)
        self.assertEqual([(r["paragraf"], r["alasan"]) for r in result],
                         [(1, "Slide menyebut 4,85%."), (2, "Periode salah.")])
        self.assertEqual(doc_composer.parse_review({"temuan": "x"}, 2), [])
        self.assertEqual(doc_composer.parse_review(["x"], 2), [])


def deck_with(summary=None, panels=None):
    slide = {"ref": "slide-1", "title": "Inflasi", "text": [FACTS], "tables": [], "charts": [], "visual_desc": None,
             "image_path": "/tmp/slide.png"}
    if summary:
        slide["summary"] = summary
    if panels:
        slide["panel_images"] = panels
    return {"items": [slide]}


class ReviewDeckTests(unittest.TestCase):
    def test_only_image_summaries_are_reviewed_never_narration(self):
        deck = deck_with(summary="Inflasi 9,9%.")
        paragraphs = fact_review.collect_paragraphs(deck)
        self.assertEqual([p["teks"] for e in paragraphs for p in e["paragraf"]], ["Inflasi 9,9%."])
        self.assertEqual(paragraphs[0]["paragraf"][0]["gambar"], "/tmp/slide.png")
        self.assertEqual(fact_review.collect_paragraphs(deck_with()), [])  # slide tanpa ringkasan gambar

    def test_split_slide_reviews_each_panel_summary_with_its_own_image(self):
        panels = [{"summary": "Panel satu.", "image_path": "/tmp/a.png"},
                  {"summary": "Panel dua.", "image_path": "/tmp/b.png"}]
        entry = fact_review.collect_paragraphs(deck_with(summary="Slide utuh.", panels=panels))[0]
        self.assertEqual([(p["teks"], p["gambar"]) for p in entry["paragraf"]],
                         [("Panel satu.", "/tmp/a.png"), ("Panel dua.", "/tmp/b.png")])

    def test_rules_only_mode_marks_flagged_summaries_without_a_model(self):
        result = fact_review.review_deck(deck_with(summary="Inflasi 9,9% pada Juli 2027."), None)
        self.assertEqual(set(result["komentar"]), {("slide-1", "Inflasi 9,9% pada Juli 2027.")})
        self.assertEqual({f["sumber"] for f in result["temuan"]}, {"aturan"})
        self.assertEqual({f["jenis"] for f in result["temuan"]}, {"angka", "periode"})
        self.assertEqual(result["ditinjau_model"], [])
        self.assertTrue(result["komentar"][("slide-1", "Inflasi 9,9% pada Juli 2027.")][0].startswith("[Cek fakta: angka]"))

    def test_model_reviews_flagged_or_number_heavy_slides_and_never_changes_the_text(self):
        calls = []

        class Model:
            workers = 1

            def review_slide(self, slide, paragraphs):
                calls.append([p["teks"] for p in paragraphs])
                return [{"paragraf": 1, "kutipan": "melemah", "alasan": "Gambar menunjukkan menguat."}]

        deck = deck_with(summary="Rupiah melemah.")
        steps = []
        fact_review.review_deck(deck, Model(), lambda *args: steps.append(args))
        self.assertEqual(calls, [])  # tanpa temuan aturan dan angka sedikit: tidak ditinjau model
        with patch.dict("os.environ", {fact_review.MIN_NUMBERS_ENV: "0"}):
            result = fact_review.review_deck(deck, Model(), lambda *args: steps.append(args))
        self.assertEqual(calls, [["Rupiah melemah."]])
        self.assertEqual(result["ditinjau_model"], ["slide-1"])
        self.assertEqual(result["komentar"][("slide-1", "Rupiah melemah.")],
                         ["[Cek fakta: model] Gambar menunjukkan menguat. Kutipan: \"melemah\"."])
        self.assertEqual(deck["items"][0]["summary"], "Rupiah melemah.")  # teks tidak berubah
        self.assertEqual(steps[-1][:2], ("review", "running"))

    def test_model_failure_is_recorded_and_does_not_stop_the_review(self):
        class Broken:
            workers = 1

            def review_slide(self, slide, paragraphs):
                raise BedrockError("throttled")

        with patch.dict("os.environ", {fact_review.MIN_NUMBERS_ENV: "0"}):
            result = fact_review.review_deck(deck_with(summary="Inflasi 2,8%."), Broken())
        self.assertEqual(result["gagal_model"], [{"ref": "slide-1", "alasan": "throttled"}])
        self.assertEqual(result["ditinjau_model"], [])


class ComposeAndBuildTests(unittest.TestCase):
    def test_compose_attaches_fact_comments_to_image_summaries_only(self):
        deck = {"items": [{"ref": "slide-1", "title": "Inflasi", "text": ["x"], "tables": [], "charts": [],
                           "image_path": "/tmp/s.png", "summary": "Ringkasan gambar.", "visual_area": 0.2,
                           "vision_reasons": ["chart"]}]}
        review = {("slide-1", "Ringkasan gambar."): ["catatan"], ("slide-1", "Narasi."): ["tidak dipakai"]}
        with patch.object(doc_composer, "slide_needs_image", return_value=True):
            spec, _ = doc_composer.compose(
                deck, ["Bab"], {"Bab": ["slide-1"]}, {"slide-1": ["Narasi."]},
                {"penomoran": [], "caption_default": "kosong"}, [], review)
        texts = {}
        for block in spec["blocks"]:
            cells = [item for row in block.get("rows", []) for cell in row for item in cell]
            for item in [block, *cells]:
                if item.get("type") == "paragraph":
                    texts[item["text"]] = item.get("comments")
        self.assertEqual(texts, {"Ringkasan gambar.": ["catatan"], "Narasi.": None})

    def test_builder_writes_word_comments_without_touching_the_text(self):
        spec = {"blocks": [
            {"type": "paragraph", "text": "Inflasi <i>year-on-year</i> 9%.", "comments": ["[Cek fakta: angka] salah", "kedua"]},
            {"type": "paragraph", "text": "Tanpa catatan."},
            {"type": "table", "widths": [50, 50], "rows": [[
                [{"type": "paragraph", "text": "Di sel.", "comments": ["catatan sel"]}], []]]},
        ]}
        buffer = io.BytesIO()
        build_docx(spec, PAGE, buffer)
        document = Document(io.BytesIO(buffer.getvalue()))
        comments = list(document.comments)
        self.assertEqual([c.text for c in comments], ["[Cek fakta: angka] salah", "kedua", "catatan sel"])
        self.assertEqual({c.author for c in comments}, {"Peninjau Fakta"})
        self.assertEqual(document.paragraphs[0].text, "Inflasi year-on-year 9%.")
        self.assertEqual(document.paragraphs[1].text, "Tanpa catatan.")


class ReviewSlideTests(unittest.TestCase):
    def test_model_review_prompt_and_parsing(self):
        generator = object.__new__(BedrockDocumentGenerator)
        sent = []

        def converse(content, max_tokens=6000, step=None):
            sent.append((content, step))
            return '{"temuan": [{"paragraf": 1, "kutipan": "9%", "alasan": "Slide menyebut 4,85%."}]}'

        slide = {"ref": "slide-1", "title": "PDB", "text": ["PDB 4,85%"], "tables": [], "charts": []}
        with tempfile.TemporaryDirectory() as tmp:
            image = Path(tmp) / "panel.png"
            image.write_bytes(b"\x89PNG")
            with patch.object(generator, "_converse", side_effect=converse):
                result = generator.review_slide(slide, [{"teks": "PDB tumbuh 9%.", "gambar": str(image)}])
        self.assertEqual(result, [{"paragraf": 1, "kutipan": "9%", "alasan": "Slide menyebut 4,85%."}])
        content, step = sent[0]
        self.assertEqual(step, "review")
        prompt = content[0]["text"]
        self.assertIn("[1] PDB tumbuh 9%.", prompt)
        self.assertIn("PDB 4,85%", prompt)
        self.assertIn("jangan menulis ulang", prompt.lower())
        self.assertEqual(content[1], {"text": "Gambar untuk paragraf [1]:"})
        self.assertIn("image", content[2])  # gambar paragraf itu ikut dikirim
        self.assertIn("ringkasan sebuah gambar", prompt)


if __name__ == "__main__":
    unittest.main()
