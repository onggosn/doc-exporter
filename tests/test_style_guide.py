import json
import unittest
from unittest.mock import patch

from bs4 import BeautifulSoup

import style_guide
from bedrock_service import BedrockDocumentGenerator, BedrockError
from style_guide import StyleGuideError

HTML = """
<div class="title-level-2"><p>2.1 Judul bagian yang cukup panjang untuk lolos batas panjang</p></div>
<p>Inflasi diprakirakan tetap terkendali sehingga mendukung stabilitas harga ke depan.</p>
<ul>
  <li>Poin pertama yang cukup panjang untuk dihitung sebagai poin dokumen
    <ul><li>Sub poin yang juga cukup panjang untuk dihitung sebagai poin sendiri</li></ul>
  </li>
</ul>
<figure><figcaption><p>Caption panjang yang tidak boleh masuk ke blok isi dokumen</p></figcaption></figure>
<p>Source : Pendalaman Penurunan Kinerja dan Identifikasi Peluang Pembiayaan Inklusif</p>
<p>Terlalu pendek.</p>
<p>Inflasi diprakirakan tetap terkendali sehingga mendukung stabilitas harga ke depan.</p>
<p>Kata <i>demand</i> dan <em>funding</em> dicetak miring dalam kalimat yang cukup panjang ini.</p>
"""


def blocks_from(html=HTML):
    return style_guide.extract_blocks(BeautifulSoup(html, "html.parser"), "a.html")


class ExtractBlocksTests(unittest.TestCase):
    def test_keeps_body_text_and_drops_headings_captions_notes_duplicates_and_short_text(self):
        texts = [b["text"] for b in blocks_from()]
        self.assertEqual(len(texts), 4)
        self.assertEqual(sum("Inflasi diprakirakan" in t for t in texts), 1)
        for dropped in ("Judul bagian", "Caption panjang", "Source :", "Terlalu pendek"):
            self.assertFalse(any(dropped in t for t in texts), dropped)

    def test_nested_list_items_are_counted_once_with_own_text(self):
        points = [b["text"] for b in blocks_from() if b["kind"] == "poin"]
        self.assertEqual(len(points), 2)
        self.assertFalse(any("Sub poin" in p for p in points if p.startswith("Poin pertama")))

    def test_italic_terms(self):
        self.assertEqual(style_guide.italic_terms(BeautifulSoup(HTML, "html.parser")), ["demand", "funding"])


class StatisticsTests(unittest.TestCase):
    def stats(self, *texts):
        blocks = [{"source": "a", "kind": "paragraf", "text": t} for t in texts]
        return style_guide.compute_statistics(blocks, [], [])

    def test_number_formats_are_counted_separately(self):
        angka = self.stats("PDB tumbuh 4,85% dan inflasi 2.5% dengan nilai tukar Rp16.550 per dolar AS.",
                           "Selisih 3 % dan kurs Rp 16.000 serta pertumbuhan 5,0%.")["angka"]
        self.assertEqual(
            (angka["desimal_koma"], angka["desimal_titik"], angka["persen_rapat"], angka["persen_spasi"]),
            (2, 1, 3, 1),
        )
        self.assertEqual((angka["rupiah_rapat"], angka["rupiah_spasi"]), (1, 1))

    def test_abbreviations_are_compared_with_full_forms(self):
        result = {a["lengkap"]: a for a in self.stats("Hal yg penting dan yang lain, dg cara dgn teliti.")["singkatan"]}
        self.assertEqual(result["yang"]["jumlah_singkat"], 1)
        self.assertEqual(result["yang"]["jumlah_lengkap"], 1)
        self.assertEqual(result["dengan"]["singkatan"], {"dg": 1, "dgn": 1})

    def test_sentence_length_and_technical_terms(self):
        stats = self.stats("Satu dua tiga empat. Lima enam tujuh delapan sembilan sepuluh. BI Rate naik 25 bps yoy.")
        self.assertEqual(stats["kalimat"]["paragraf"]["jumlah"], 3)
        self.assertEqual(stats["istilah_teknis"], {"bps": 1, "yoy": 1})


class SelectSamplesTests(unittest.TestCase):
    def blocks(self, n=100):
        return [{"source": "a", "kind": "paragraf", "text": f"Paragraf nomor {i} " + "x" * 200} for i in range(n)]

    def test_selection_is_deterministic_spread_and_bounded(self):
        first = style_guide.select_samples(self.blocks(), [], max_chars=5000)
        again = style_guide.select_samples(self.blocks(), [], max_chars=5000)
        self.assertEqual(first, again)
        self.assertLessEqual(sum(len(s["text"]) for s in first), 5000)
        numbers = [int(s["text"].split()[2]) for s in first]
        self.assertLess(numbers[0], 10)
        self.assertGreater(numbers[-1], 60)

    def test_group_cap_and_caption_inclusion(self):
        captions = [{"source_file": "a", "type": "headline", "text": f"Caption {i}"} for i in range(80)]
        samples = style_guide.select_samples([], captions)
        self.assertEqual(len(samples), style_guide.MAX_BLOCKS_PER_GROUP)
        self.assertEqual({s["kind"] for s in samples}, {"caption"})

    def test_empty_inputs(self):
        self.assertEqual(style_guide.select_samples([], []), [])


def good_guide(**overrides):
    guide = {key: f"Isi {key}." for key in style_guide.GUIDE_TEXT_FIELDS}
    guide.update({
        "gaya_caption": {"headline": "Kalimat pesan utama.", "label": "Frasa pendek."},
        "larangan": ["Mengarang angka."],
        "contoh": [{"konteks": "paragraf", "teks": "Inflasi diprakirakan tetap terkendali"}],
    })
    guide.update(overrides)
    return guide


SAMPLES = [{"source": "a", "kind": "paragraf", "text": "Inflasi  diprakirakan tetap terkendali\nsehingga baik."}]


class ValidateGuideTests(unittest.TestCase):
    def test_accepts_guide_and_keeps_verbatim_examples_ignoring_whitespace(self):
        cleaned, dropped = style_guide.validate_guide(good_guide(), SAMPLES)
        self.assertEqual(dropped, 0)
        self.assertEqual(cleaned["contoh"][0]["konteks"], "paragraf")

    def test_drops_examples_that_are_not_exact_quotes_or_have_bad_context(self):
        examples = [
            {"konteks": "paragraf", "teks": "Inflasi diprakirakan melonjak tajam"},
            {"konteks": "lain", "teks": "Inflasi diprakirakan tetap terkendali"},
            "bukan objek",
            {"konteks": "poin", "teks": "tetap terkendali sehingga baik"},
        ]
        cleaned, dropped = style_guide.validate_guide(good_guide(contoh=examples), SAMPLES)
        self.assertEqual((len(cleaned["contoh"]), dropped), (1, 3))

    def test_rejects_missing_empty_or_overlong_fields(self):
        for bad in (
            good_guide(nada=""),
            good_guide(nada="x" * (style_guide.MAX_GUIDE_FIELD_CHARS + 1)),
            good_guide(gaya_caption={"headline": "ada"}),
            good_guide(gaya_caption="bukan objek"),
            good_guide(larangan="bukan daftar"),
            good_guide(contoh="bukan daftar"),
        ):
            with self.assertRaises(StyleGuideError):
                style_guide.validate_guide(bad, SAMPLES)
        incomplete = good_guide()
        del incomplete["format_angka"]
        with self.assertRaises(StyleGuideError):
            style_guide.validate_guide(incomplete, SAMPLES)

    def test_prompt_carries_statistics_samples_and_exact_quote_rule(self):
        prompt = style_guide.build_prompt(SAMPLES, {"angka": {"desimal_koma": 7}})
        self.assertIn('"desimal_koma": 7', prompt)
        self.assertIn("[paragraf] Inflasi", prompt)
        self.assertIn("SALIN PERSIS", prompt)
        self.assertIn("Jangan menulis fakta", prompt)


class GenerateStyleGuideTests(unittest.TestCase):
    def setUp(self):
        self.generator = object.__new__(BedrockDocumentGenerator)

    def test_returns_validated_guide_from_model_json(self):
        reply = "```json\n" + json.dumps(good_guide()) + "\n```"
        with patch.object(BedrockDocumentGenerator, "_converse", return_value=reply) as converse:
            guide, dropped = self.generator.generate_style_guide(SAMPLES, {})
        self.assertEqual((guide["nada"], dropped), ("Isi nada.", 0))
        self.assertIn("Inflasi", converse.call_args.args[0][0]["text"])

    def test_invalid_guide_becomes_bedrock_error(self):
        with patch.object(BedrockDocumentGenerator, "_converse", return_value=json.dumps({"nada": "x"})):
            with self.assertRaisesRegex(BedrockError, "tidak valid"):
                self.generator.generate_style_guide(SAMPLES, {})

    def test_non_json_reply_and_empty_samples(self):
        with patch.object(BedrockDocumentGenerator, "_converse", return_value="bukan json"):
            with self.assertRaisesRegex(BedrockError, "bukan JSON valid"):
                self.generator.generate_style_guide(SAMPLES, {})
        with self.assertRaisesRegex(BedrockError, "sampel"):
            self.generator.generate_style_guide([], {})


if __name__ == "__main__":
    unittest.main()
