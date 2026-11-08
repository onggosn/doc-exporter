import json
import unittest
from unittest.mock import patch

import doc_composer as dc
from bedrock_service import BedrockDocumentGenerator, BedrockError, NarrationPartial


def slide(ref, title="", text=(), tables=(), charts=(), reasons=(), image=None, notes="", desc=""):
    item = {
        "ref": ref, "title": title, "text": list(text), "tables": list(tables), "charts": list(charts),
        "vision_reasons": list(reasons), "notes": notes, "visual_desc": desc or None,
    }
    if image:
        item["image_path"] = image
    return item


def rule(keyword, prefix, start=1, mode="per_bagian", active=True):
    return {"bagian_cocok": keyword, "awalan": prefix, "mulai": start, "lanjut": mode, "aktif": active}


CHART = {"title": "Kredit", "categories": ["Q1", "Q2"], "series": [{"name": "UMKM", "values": [4.85, None, 5.0]}]}


class SlideRulesTests(unittest.TestCase):
    def test_which_slides_become_images(self):
        self.assertTrue(dc.slide_needs_image(slide("a", reasons=["picture"])))
        self.assertTrue(dc.slide_needs_image(slide("a", reasons=["chart"])))
        self.assertTrue(dc.slide_needs_image(slide("a", reasons=["diagram_or_object"])))
        self.assertTrue(dc.slide_needs_image(slide("a", tables=[[["a", "1"], ["b", "2"]]])))
        self.assertFalse(dc.slide_needs_image(slide("a", tables=[[["a", "b"], ["c", "d"]]], reasons=["picture"])))
        self.assertFalse(dc.slide_needs_image(slide("a", reasons=["low_text"])))
        self.assertFalse(dc.slide_needs_image(slide("a")))

    def test_long_text_slide_with_only_a_decorative_picture_is_text(self):
        long_text = ["Kalimat isi slide yang cukup panjang. " * 10] * 5
        self.assertFalse(dc.slide_needs_image(slide("a", "", long_text, reasons=["picture"])))
        self.assertTrue(dc.slide_needs_image(slide("a", "", ["Pendek."], reasons=["picture"])))
        self.assertTrue(dc.slide_needs_image(slide("a", "", long_text, reasons=["picture", "chart"], charts=[CHART])))
        self.assertTrue(dc.slide_needs_image(slide("a", "", long_text, reasons=["picture", "diagram_or_object"])))
        self.assertEqual(dc.narration_form(slide("a", "", long_text, reasons=["picture"])), "lengkap")
        self.assertEqual(dc.narration_form(slide("a", "", ["Pendek."])), "paragraf")

    def test_narration_form_depends_on_image(self):
        self.assertEqual(dc.narration_form(slide("a", image="x.png")), "gambar")
        self.assertEqual(dc.narration_form(slide("a")), "paragraf")

    def test_fact_text_covers_every_source_of_facts(self):
        facts = dc.fact_text(slide(
            "a", "Judul", ["baris"], tables=[[["h1", "h2"], ["1", "2"]]], charts=[CHART], notes="catatan", desc="visual"))
        for expected in ("Judul", "baris", "h1 h2", "Q1 Q2", "UMKM", "4.85 5.0", "catatan", "visual"):
            self.assertIn(expected, facts)

    def test_substance_ignores_title(self):
        self.assertFalse(dc.has_substance(slide("a", "Hanya judul")))
        self.assertTrue(dc.has_substance(slide("a", "Judul", ["isi"])))
        self.assertTrue(dc.has_substance(slide("a", desc="gambar grafik")))


class NumberTests(unittest.TestCase):
    def test_equivalent_spellings_normalize_the_same(self):
        for a, b in (("4,85", "4.85"), ("16.550", "16550"), ("5,0", "5"), ("10.000.000", "10000000"), ("2,50", "2.5")):
            self.assertEqual(dc.normalize_number(a), dc.normalize_number(b), (a, b))

    def test_different_numbers_stay_different(self):
        self.assertNotEqual(dc.normalize_number("4,85"), dc.normalize_number("4,58"))
        self.assertNotEqual(dc.normalize_number("16.550"), dc.normalize_number("16,55"))

    def test_unsupported_numbers_allow_decimal_style_change_but_flag_new_numbers(self):
        facts = "PDB 2025 tumbuh 4.85% dan kurs Rp16.550"
        self.assertEqual(dc.unsupported_numbers(["PDB 2025 tumbuh 4,85% dan kurs Rp16.550."], facts), [])
        self.assertEqual(dc.unsupported_numbers(["Tumbuh 5,1% pada 2026"], facts), ["2026", "5.1"])


class StyleContextTests(unittest.TestCase):
    STATS = {"angka": {"desimal_koma": 158, "desimal_titik": 17, "persen_rapat": 194, "persen_spasi": 8,
                       "rupiah_rapat": 59, "rupiah_spasi": 0}, "miring": {"jumlah": 475}}

    def test_format_rules_follow_the_dominant_usage(self):
        rules = " ".join(dc.format_rules(self.STATS))
        for expected in ("koma (4,85)", "rapat dengan angka (4,5%)", "rupiah rapat", "<i>"):
            self.assertIn(expected, rules)
        flipped = {**self.STATS, "angka": {**self.STATS["angka"], "desimal_koma": 1, "desimal_titik": 9, "rupiah_rapat": 0, "rupiah_spasi": 3}}
        self.assertIn("titik (4.85)", " ".join(dc.format_rules(flipped)))
        self.assertIn("dengan spasi (Rp 16.550)", " ".join(dc.format_rules(flipped)))

    def test_context_includes_guide_when_present_and_empty_without_style(self):
        self.assertEqual(dc.style_context(None), "")
        guide = {key: f"isi {key}" for key in (
            "nada", "sudut_pandang", "panjang_kalimat", "gaya_poin", "istilah_dan_singkatan", "format_angka", "istilah_asing")}
        guide.update({"larangan": ["mengarang"], "contoh": [{"konteks": "poin", "teks": "Kutipan gaya"}]})
        text = dc.style_context({"statistik": self.STATS, "panduan": {"panduan": guide}})
        for expected in ("Nada: isi nada", "Hindari: mengarang", '[poin] "Kutipan gaya"', "koma"):
            self.assertIn(expected, text)
        self.assertNotIn("Nada", dc.style_context({"statistik": self.STATS, "panduan": None}))


class PickWidthsTests(unittest.TestCase):
    def test_most_common_numeric_bucket_for_image_left_text_right(self):
        patterns = [{"tokens": "I,I", "width_buckets": {"50/20/30": 6}},
                    {"tokens": "I,T", "width_buckets": {"auto": 9, "40/60": 7, "50/50": 13}}]
        self.assertEqual(dc.pick_widths(patterns), [50, 50])

    def test_falls_back_to_default(self):
        self.assertEqual(dc.pick_widths([]), dc.DEFAULT_WIDTHS)
        self.assertEqual(dc.pick_widths([{"tokens": "I,T", "width_buckets": {"auto": 3}}]), dc.DEFAULT_WIDTHS)


class CaptionNumberingTests(unittest.TestCase):
    def numbering(self, rules, default="headline"):
        return dc.CaptionNumbering({"penomoran": rules, "caption_default": default})

    def test_label_kind_follows_content_and_counts_per_kind(self):
        n = self.numbering([rule("Lampiran", "Grafik"), rule("Lampiran", "Tabel")])
        self.assertEqual(n.caption_for(slide("a", "Kredit", charts=[CHART]), "Lampiran Data"),
                         {"text": "Grafik 1. Kredit", "tipe": "label"})
        self.assertEqual(n.caption_for(slide("b", "Sektor", tables=[[["x"]]]), "Lampiran Data")["text"], "Tabel 1. Sektor")
        self.assertEqual(n.caption_for(slide("c", "Lagi", charts=[CHART]), "Lampiran Data")["text"], "Grafik 2. Lagi")
        self.assertEqual(n.caption_for(slide("d", "Foto saja"), "Lampiran Data")["text"], "Grafik 3. Foto saja")

    def test_restart_per_section_versus_continue(self):
        s = slide("a", "T", charts=[CHART])
        restart = self.numbering([rule("Lampiran", "Grafik", mode="per_bagian")])
        self.assertEqual([restart.caption_for(s, c)["text"] for c in ("Lampiran A", "Lampiran B")], ["Grafik 1. T", "Grafik 1. T"])
        carry = self.numbering([rule("Lampiran", "Grafik", mode="berlanjut")])
        self.assertEqual([carry.caption_for(s, c)["text"] for c in ("Lampiran A", "Lampiran B")], ["Grafik 1. T", "Grafik 2. T"])

    def test_start_number_and_headline_prefix_without_space(self):
        n = self.numbering([rule("Kesimpulan", "K", start=3)])
        self.assertEqual(n.caption_for(slide("a", "Pesan"), "Kesimpulan Moneter"), {"text": "K3. Pesan", "tipe": "headline"})
        self.assertEqual(n.caption_for(slide("b", "Lain"), "Kesimpulan Moneter")["text"], "K4. Lain")

    def test_title_is_optional_and_matching_ignores_case(self):
        n = self.numbering([rule("lampiran", "Grafik")])
        self.assertEqual(n.caption_for(slide("a", charts=[CHART]), "LAMPIRAN")["text"], "Grafik 1")

    def test_default_caption_inactive_rules_and_no_title(self):
        n = self.numbering([rule("Lampiran", "Grafik", active=False)], default="headline")
        self.assertEqual(n.caption_for(slide("a", "Judul"), "Lampiran"), {"text": "Judul", "tipe": "headline"})
        self.assertIsNone(n.caption_for(slide("b"), "Lampiran"))
        self.assertIsNone(self.numbering([], default="kosong").caption_for(slide("c", "Judul"), "X"))

    def test_no_caption_when_setting_is_missing(self):
        self.assertIsNone(dc.CaptionNumbering({"penomoran": []}).caption_for(slide("a", "Judul"), "X"))

    def test_rules_for_reports_matching_prefixes(self):
        n = self.numbering([rule("Lampiran", "Grafik"), rule("Lampiran", "Tabel"), rule("Kesimpulan", "K")])
        self.assertEqual(n.rules_for("Lampiran Data"), ["Grafik", "Tabel"])
        self.assertEqual(n.rules_for("Pendahuluan"), [])


class TrimMarginsTests(unittest.TestCase):
    def test_plain_border_is_cropped_with_padding(self):
        from PIL import Image
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            path = f"{tmp}/s.png"
            image = Image.new("RGB", (400, 300), "white")
            image.paste((200, 0, 0), (100, 80, 200, 160))
            image.save(path)
            dc.trim_margins(path, padding=10)
            self.assertEqual(Image.open(path).size, (120, 100))

    def test_image_without_margin_or_blank_image_is_untouched(self):
        from PIL import Image
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            blank, full = f"{tmp}/b.png", f"{tmp}/f.png"
            Image.new("RGB", (50, 40), "white").save(blank)
            noise = Image.new("RGB", (50, 40), "white")
            noise.paste((0, 0, 0), (0, 0, 50, 5))
            noise.paste((0, 0, 0), (0, 35, 50, 40))
            noise.paste((0, 0, 0), (0, 0, 5, 40))
            noise.paste((0, 0, 0), (45, 0, 50, 40))
            noise.save(full)
            dc.trim_margins(blank)
            dc.trim_margins(full)
            self.assertEqual((Image.open(blank).size, Image.open(full).size), ((50, 40), (50, 40)))


class NarrationParsingTests(unittest.TestCase):
    def test_accepts_clean_lists_and_strips_bullet_markers(self):
        result = dc.parse_narration({"s1": ["- Poin satu", "• Poin <i>dua</i>"], "s2": ["Paragraf."]}, ["s1", "s2"])
        self.assertEqual(result, {"s1": ["Poin satu", "Poin <i>dua</i>"], "s2": ["Paragraf."]})

    def test_rejects_bad_shapes(self):
        for data in ([], {}, {"s1": "teks"}, {"s1": []}, {"s1": [" "]}, {"s1": [1]}, {"s1": ["a"] * 7},
                     {"s1": ["**tebal**"]}, {"s1": ["<p>html</p>"]}, {"s1": ["```kode```"]}):
            with self.assertRaises(dc.NarrationError, msg=repr(data)[:40]):
                dc.parse_narration(data, ["s1"])

    def test_dense_slide_limit_follows_fact_lines_and_is_capped(self):
        slide = lambda n: {"ref": "s1", "title": "", "text": [f"baris {i}" for i in range(n)]}
        self.assertEqual(dc.narration_limit(slide(3)), dc.MAX_NARRATION_ITEMS)
        self.assertEqual(dc.narration_limit(slide(18)), 18)
        self.assertEqual(dc.narration_limit(slide(90)), dc.MAX_NARRATION_ITEMS_DENSE)
        self.assertEqual(dc.parse_narration({"s1": ["a"] * 17}, ["s1"], limits={"s1": 18})["s1"], ["a"] * 17)
        with self.assertRaisesRegex(dc.NarrationError, "1-18 butir"):
            dc.parse_narration({"s1": ["a"] * 19}, ["s1"], limits={"s1": 18})

    def test_divider_slides_may_be_skipped_but_slides_with_numbers_or_images_may_not(self):
        divider = {"ref": "s1", "title": "", "text": ["ASESMEN EKONOMI", "DEPARTEMEN KEBIJAKAN"]}
        self.assertTrue(dc.may_skip(divider))
        self.assertFalse(dc.may_skip({**divider, "text": ["Inflasi 3,2%"]}))
        self.assertFalse(dc.may_skip({**divider, "text": ["x" * 201]}))
        self.assertFalse(dc.may_skip({**divider, "image_path": "a.png"}))
        self.assertTrue(dc.may_skip({"ref": "s2", "title": "J", "text": [], "image_path": "a.png"}))

    def test_overlong_item_is_split_at_sentence_boundaries_instead_of_rejected(self):
        sentence = "Kalimat ini cukup panjang untuk uji pemecahan butir. "
        result = dc.parse_narration({"s1": [sentence * 30]}, ["s1"], limits={"s1": 6})["s1"]
        self.assertGreater(len(result), 1)
        self.assertTrue(all(len(item) <= dc.MAX_NARRATION_CHARS and item.endswith(".") for item in result))
        self.assertEqual(" ".join(result), (sentence * 30).strip())
        words = dc.split_long("kata " * 400)  # tanpa tanda baca: dipotong di spasi
        self.assertTrue(all(len(item) <= dc.MAX_NARRATION_CHARS for item in words))

    def test_fallback_groups_short_fragments_and_keeps_long_lines_apart(self):
        slide_ = slide("s1", "J", ["616", "69 Bank", "10%", "Kalimat penuh yang cukup panjang untuk berdiri sendiri sebagai butir.", "10%", "19,1%"])
        self.assertEqual(dc.fallback_narration([slide_])["s1"],
                         ["616; 69 Bank; 10%", "Kalimat penuh yang cukup panjang untuk berdiri sendiri sebagai butir.", "19,1%"])

    def test_visual_only_slides_may_be_empty_or_missing_but_others_may_not(self):
        data = {"s1": ["Teks."], "s2": []}
        self.assertEqual(dc.parse_narration(data, ["s1", "s2", "s3"], {"s2", "s3"}), {"s1": ["Teks."], "s2": [], "s3": []})
        for bad in ({"s1": [], "s2": []}, {"s2": []}):
            with self.assertRaises(dc.NarrationError):
                dc.parse_narration(bad, ["s1", "s2"], {"s2"})

    def test_visual_only_ignores_title_and_model_description(self):
        self.assertTrue(dc.visual_only(slide("a", "Judul", desc="gambar bulat kuning")))
        self.assertFalse(dc.visual_only(slide("a", "Judul", ["fakta"])))
        self.assertFalse(dc.visual_only(slide("a", "Judul", charts=[CHART])))

    def test_prompt_forbids_interpretation_and_allows_empty_for_pure_images(self):
        prompt = dc.narration_prompt("Bab", [slide("s1", "J", ["f"])], "")
        self.assertIn("Jangan menambah tafsiran", prompt)
        self.assertIn("daftar kosong []", prompt)

    def test_prompt_lists_slides_forms_rules_and_style(self):
        prompt = dc.narration_prompt(
            "Bab Uji", [slide("s1", "Judul", ["fakta satu"], image="x.png"), slide("s2", "Lain", ["fakta dua"])],
            "Pisahkan desimal dengan koma.")
        for expected in ("Bab Uji", "[s1] bentuk: gambar", "[s2] bentuk: paragraf", "fakta satu", "Pisahkan desimal",
                         "Hanya pakai fakta", '{"s1": ["…"], "s2": ["…"]}'):
            self.assertIn(expected, prompt)


class ComposeTests(unittest.TestCase):
    def setUp(self):
        self.deck = {"items": [
            slide("s1", "Ringkasan Umum", ["Ringkasan ditulis sebagai kalimat utuh."]),
            slide("s2", "Kredit UMKM", charts=[CHART], reasons=["chart"], image="/tmp/s2.png"),
            slide("s3", "Sektor", tables=[[["Sektor", "12,5"]]], image="/tmp/s3.png"),
            slide("s4", "Hanya judul"),
        ]}
        self.chapters = ["Ringkasan", "Lampiran Data", "Penutup"]
        self.mapping = {"Ringkasan": ["s2", "s1"], "Lampiran Data": ["s3", "s4"], "Penutup": []}
        self.narration = {"s1": ["Paragraf satu.", "Paragraf dua."], "s2": ["Poin A", "Poin B"]}
        self.settings = {"penomoran": [rule("Lampiran", "Grafik"), rule("Lampiran", "Tabel")], "caption_default": "headline"}
        self.patterns = [{"tokens": "I,T", "width_buckets": {"40/60": 5}}]

    def compose(self):
        return dc.compose(self.deck, self.chapters, self.mapping, self.narration, self.settings, self.patterns)

    def test_structure_follows_slide_order_within_chapter(self):
        spec, _ = self.compose()
        kinds = [(b["type"], b.get("level") or "") for b in spec["blocks"]]
        self.assertEqual(kinds, [("heading", 1), ("heading", 2), ("paragraph", ""), ("paragraph", ""),
                                 ("table", ""), ("paragraph", ""), ("paragraph", ""), ("heading", 1), ("table", "")])
        self.assertEqual(spec["blocks"][1]["text"], "Ringkasan Umum")

    def test_text_slide_is_written_as_paragraphs_under_nested_heading_and_reported(self):
        long_text = ["Kalimat isi slide yang cukup panjang. " * 10] * 5
        self.deck["items"].append(slide("s9", "Kesimpulan", long_text, reasons=["picture"]))
        self.chapters = ["2.1 Rinci"]
        self.mapping = {"2.1 Rinci": ["s9"]}
        self.narration = {"s9": ["Butir satu.", "Butir dua."]}
        spec, report = self.compose()
        self.assertEqual([(b["type"], b.get("level")) for b in spec["blocks"]],
                         [("heading", 2), ("heading", 3), ("paragraph", None), ("paragraph", None)])
        self.assertEqual((report["slide_teks"], report["gambar"]), (["s9"], 0))

    def test_sub_chapters_get_nested_heading_levels(self):
        for chapter, level in (("1. Kesimpulan", 1), ("2.1 Ringkasan", 2), ("4.7.2 Pokok", 3), ("4.7.2.1 Dalam", 3),
                               ("Lampiran", 1), ("2024 Review", 1)):
            self.assertEqual(dc.heading_level(chapter), level, chapter)

    def test_cover_agenda_divider_and_closing_slides_are_filler_but_content_is_not(self):
        filler = [
            ["Agenda Pembahasan", "RDG Mei Hari I", "Bauran Kebijakan – Mei 2025"],
            ["ASESMEN EKONOMI DAN MONETER", "DEPARTEMEN KEBIJAKAN EKONOMI DAN MONETER"],
            ["TERIMA KASIH"], ["Jakarta, 21 Oktober 2025"],
        ]
        for text in filler:
            self.assertTrue(dc.is_filler(slide("x", "", text, reasons=["picture"])), text)
        for text in (["Inflasi turun menjadi 3,2%."], ["Rupiah menguat."], ["Rp16.500 per USD"], ["x" * 301], [f"b{i}" for i in range(11)]):
            self.assertFalse(dc.is_filler(slide("x", "", text)), text)
        self.assertFalse(dc.is_filler(slide("x", "", ["Judul"], charts=[CHART], reasons=["chart"])))
        self.assertFalse(dc.is_filler(slide("x", "", [], reasons=["picture"])))  # foto tanpa teks tetap dipakai
        self.assertFalse(dc.is_filler(slide("x", "Proyeksi PDB Nominal", ["Updated 20/05/2025"])))  # ber-judul = isi

    def test_image_with_summary_is_a_two_column_row_image_left_caption_and_summary_right(self):
        self.deck["items"][1]["summary"] = "Kredit UMKM tumbuh pada periode yang ditampilkan."
        spec, report = self.compose()
        table = spec["blocks"][4]
        self.assertEqual((table["type"], table["widths"]), ("table", [40, 60]))
        (image_cell, caption_cell), = table["rows"]
        self.assertEqual(image_cell, [{"type": "image", "source": "/tmp/s2.png"}])
        self.assertEqual(caption_cell, [
            {"type": "caption", "text": "Kredit UMKM", "tipe": "headline", "align": "left"},
            {"type": "paragraph", "text": "Kredit UMKM tumbuh pada periode yang ditampilkan."},
        ])
        self.assertEqual(report["gambar_diringkas"], 1)

    def test_narration_of_an_image_slide_follows_the_table_instead_of_sharing_the_row(self):
        spec, _ = self.compose()
        self.assertEqual(spec["blocks"][5:7], [{"type": "paragraph", "text": "Poin A"}, {"type": "paragraph", "text": "Poin B"}])
        self.assertNotIn("bullets", [b["type"] for b in spec["blocks"]])  # narasi model tidak pernah jadi bullet
        self.assertTrue(all(len(row) <= 2 for row in spec["blocks"][4]["rows"]))

    def test_numbered_label_caption_sits_beside_its_image_and_image_without_text_stays_alone(self):
        spec, _ = self.compose()
        image_cell, caption_cell = spec["blocks"][8]["rows"][0]
        self.assertEqual([i["type"] for i in image_cell], ["image"])
        self.assertEqual((caption_cell[0]["text"], caption_cell[0]["tipe"]), ("Tabel 1. Sektor", "label"))
        self.settings = {"penomoran": [], "caption_default": "kosong"}
        spec, _ = self.compose()
        self.assertEqual(len(spec["blocks"][8]["rows"][0]), 1)  # tanpa caption dan tanpa ringkasan: hanya gambar

    def test_image_summary_prompt_forbids_extra_facts_and_includes_the_context(self):
        prompt = dc.image_summary_prompt("Judul S", "PANEL A", "Kredit tumbuh 8,88%", "\n\nUlangi.")
        for expected in ("Jangan menambah fakta", "Judul panel gambar: PANEL A", "Kredit tumbuh 8,88%", "Ulangi."):
            self.assertIn(expected, prompt)

    def test_clean_summary_strips_markdown_and_tags_and_cuts_at_a_sentence(self):
        self.assertEqual(dc.clean_summary("- **Kredit** <b>naik</b>.\nLalu turun."), "Kredit naik. Lalu turun.")
        long = ("Kalimat pendek nomor satu. " * 30).strip()
        cleaned = dc.clean_summary(long)
        self.assertLessEqual(len(cleaned), dc.SUMMARY_MAX_CHARS)
        self.assertTrue(cleaned.endswith("."))

    def test_report_says_which_numbering_rules_matched_each_chapter(self):
        report = self.compose()[1]
        self.assertEqual(report["aturan_bab"], {"Ringkasan": [], "Lampiran Data": ["Grafik", "Tabel"]})

    def test_report_counts_and_skips(self):
        spec, report = self.compose()
        self.assertEqual((report["gambar"], report["narasi"]), (2, 2))
        self.assertEqual((report["slide_dilewati"], report["bab_kosong"]), (["s4"], ["Penutup"]))
        final = dc.build_report(spec, report, {"s2": ["5.1"], "s1": []})
        self.assertEqual(final["gambar_ganda"], [])
        self.assertEqual(final["angka_tidak_didukung"], [{"ref": "s2", "angka": ["5.1"]}])

    def test_duplicate_image_is_reported(self):
        self.mapping["Lampiran Data"] = ["s3"]
        self.deck["items"].append(slide("s5", "Sama", tables=[[["y", "7"]]], image="/tmp/s3.png"))
        self.mapping["Penutup"] = ["s5"]
        spec, report = self.compose()
        self.assertEqual(dc.build_report(spec, report, {})["gambar_ganda"], ["/tmp/s3.png"])


class ExclusionAndVerbatimTests(unittest.TestCase):
    def test_cover_and_everything_from_the_first_appendix_divider_are_excluded(self):
        items = [slide("s1", "", ["Materi Rapat", "Jakarta, 21 Mei 2025"]),
                 slide("s2", "Isi", ["Inflasi turun menjadi 3,2% pada 2025."]),
                 slide("s3", "Lampiran dalam judul bukan pembatas yang panjang", ["x" * 400]),
                 slide("s4", "", ["LAMPIRAN KEBIJAKAN MONETER"]),
                 slide("s5", "Data", ["Isi lampiran."]),
                 slide("s6", "", ["LAMPIRAN SISTEM PEMBAYARAN"]),
                 slide("s7", "Data lain", ["Isi lampiran lain."])]
        self.assertEqual(dc.excluded_slides(items), {"s1": "sampul", "s4": "lampiran", "s5": "lampiran", "s6": "lampiran", "s7": "lampiran"})

    def test_first_slide_with_real_content_is_not_a_cover(self):
        items = [slide("s1", "Ringkasan", ["x" * 500]), slide("s2", "Grafik", charts=[CHART], reasons=["chart"])]
        self.assertEqual(dc.excluded_slides(items), {})
        self.assertEqual(dc.excluded_slides([slide("s1", "Judul", tables=[[["a", "1"]]])]), {})

    def test_split_verbatim_moves_block_lines_out_of_the_text_and_keeps_the_original(self):
        item = slide("s1", "Judul", ["HASIL PEMBAHASAN DEWAN GUBERNUR", "Kalimat satu.", "  Sub butir.", "Teks lain di slide."])
        item["verbatim"] = [{"label": "HASIL PEMBAHASAN DEWAN GUBERNUR", "lines": ["Kalimat satu.", "  Sub butir."]}]
        dc.split_verbatim(item)
        dc.split_verbatim(item)  # aman diulang
        self.assertEqual(item["text"], ["Teks lain di slide."])
        self.assertEqual(len(item["text_original"]), 4)
        self.assertEqual(dc.fact_text(item), "Judul\nTeks lain di slide.")  # tidak dinarasikan ulang
        self.assertIn("Sub butir.", dc.all_facts(item))  # tetap dihitung untuk pemeriksaan angka
        self.assertFalse(dc.has_substance({**item, "text": []}))

    def test_verbatim_blocks_keep_text_unchanged_with_labels_bold_and_nesting_kept(self):
        blocks = dc.verbatim_blocks(["KESIMPULAN", "Dampak tarif <5% perlu ditakar.", "  Butir turunan.", "ARAHAN", "DKEM memperdalam proyeksi."])
        self.assertEqual(blocks, [
            {"type": "paragraph", "text": "<b>KESIMPULAN</b>"},
            {"type": "bullets", "items": ["Dampak tarif 5% perlu ditakar.", "  Butir turunan."]},
            {"type": "paragraph", "text": "<b>ARAHAN</b>"},
            {"type": "bullets", "items": ["DKEM memperdalam proyeksi."]},
        ])
        self.assertFalse(dc._is_label("Bank Indonesia menjaga stabilitas."))
        self.assertTrue(dc._is_label("KEY TAKEAWAYS"))

    def test_conclusion_and_direction_chapters_carry_slide_text_verbatim_without_narration_or_summary(self):
        item = slide("s1", "Kesimpulan Diskusi", ["MONETER", "KESIMPULAN", "Dampak tarif perlu ditakar.", "ARAHAN", "DKEM memperdalam proyeksi."],
                     reasons=["picture"], image="/tmp/w.png")
        item["text"] = item["text"] * 1
        item["summary"] = "Ringkasan yang tidak boleh dipakai."
        spec, report = dc.compose({"items": [item]}, ["1.1 Kesimpulan dan Arahan Moneter"], {"1.1 Kesimpulan dan Arahan Moneter": ["s1"]},
                                  {"s1": ["Narasi model yang tidak boleh dipakai."]}, {"penomoran": [], "caption_default": "kosong"}, [])
        texts = [b.get("text") or b.get("items") for b in spec["blocks"]]
        self.assertIn(["Dampak tarif perlu ditakar."], texts)
        self.assertIn("<b>ARAHAN</b>", texts)
        self.assertNotIn("Narasi model yang tidak boleh dipakai.", str(spec))
        self.assertEqual(report["slide_teks_asli"], ["s1"])

    def test_labelled_blocks_are_carried_verbatim_in_other_chapters_while_the_rest_is_narrated(self):
        item = slide("s1", "Prospek global", ["HASIL PEMBAHASAN DEWAN GUBERNUR", "Tarif turun menjadi 11,9%.", "Teks lain."],
                     reasons=["chart"], charts=[CHART], image="/tmp/w.png")
        item["verbatim"] = [{"label": "HASIL PEMBAHASAN DEWAN GUBERNUR", "lines": ["Tarif turun menjadi 11,9%."]}]
        dc.split_verbatim(item)
        spec, report = dc.compose({"items": [item]}, ["2.1 Ringkasan Moneter"], {"2.1 Ringkasan Moneter": ["s1"]},
                                  {"s1": ["Narasi dari teks lain."]}, {"penomoran": [], "caption_default": "kosong"}, [])
        kinds = [(b["type"], b.get("text") or "") for b in spec["blocks"]]
        self.assertEqual([k[0] for k in kinds], ["heading", "heading", "paragraph", "bullets", "table", "paragraph"])
        self.assertEqual(spec["blocks"][3]["items"], ["Tarif turun menjadi 11,9%."])
        self.assertEqual(spec["blocks"][5], {"type": "paragraph", "text": "Narasi dari teks lain."})
        self.assertEqual(report["slide_teks_asli"], ["s1"])


class WriteChapterNarrationTests(unittest.TestCase):
    SLIDES = [slide("s1", "Judul", ["PDB tumbuh 4.85% pada 2025"])]

    def run_with(self, replies):
        generator = object.__new__(BedrockDocumentGenerator)
        with patch.object(BedrockDocumentGenerator, "_converse", side_effect=replies) as converse:
            try:
                return generator.write_chapter_narration("Bab", self.SLIDES, "gaya"), converse
            except BedrockError as exc:
                return exc, converse

    def test_clean_answer_is_returned_after_one_call(self):
        (narration, flagged), converse = self.run_with([json.dumps({"s1": ["PDB tumbuh 4,85% pada 2025."]})])
        self.assertEqual((narration, flagged), ({"s1": ["PDB tumbuh 4,85% pada 2025."]}, {}))
        self.assertEqual(converse.call_count, 1)

    def test_invented_number_triggers_one_retry_with_feedback(self):
        replies = [json.dumps({"s1": ["PDB tumbuh 5,1%."]}), json.dumps({"s1": ["PDB tumbuh 4,85%."]})]
        (narration, flagged), converse = self.run_with(replies)
        self.assertEqual((narration["s1"], flagged), (["PDB tumbuh 4,85%."], {}))
        self.assertIn("5.1", converse.call_args.args[0][0]["text"])

    def test_number_still_wrong_after_retry_is_returned_as_warning(self):
        replies = [json.dumps({"s1": ["Tumbuh 5,1%."]})] * 2
        (narration, flagged), converse = self.run_with(replies)
        self.assertEqual((narration["s1"], flagged), (["Tumbuh 5,1%."], {"s1": ["5.1"]}))
        self.assertEqual(converse.call_count, 2)

    def test_pure_image_slide_may_come_back_empty(self):
        self.SLIDES = [slide("s1", "Judul", ["PDB 4.85%"]), slide("s2", "Ilustrasi", desc="lingkaran kuning")]
        replies = [json.dumps({"s1": ["PDB 4,85%."], "s2": []})]
        (narration, flagged), _ = self.run_with(replies)
        self.assertEqual((narration, flagged), ({"s1": ["PDB 4,85%."], "s2": []}, {}))

    def test_large_chapter_is_split_into_batches_and_merged(self):
        self.SLIDES = [slide(f"s{i}", "J", [f"PDB {i}"]) for i in range(1, 13)]
        replies = [
            json.dumps({f"s{i}": [f"PDB {i}."] for i in range(start, min(start + 5, 13))}) for start in (1, 6, 11)
        ]
        (narration, flagged), converse = self.run_with(replies)
        self.assertEqual(converse.call_count, 3)
        self.assertEqual((list(narration), flagged), ([f"s{i}" for i in range(1, 13)], {}))
        self.assertEqual(converse.call_args.kwargs["max_tokens"], BedrockDocumentGenerator.NARRATION_MAX_TOKENS)

    def test_failed_batch_does_not_discard_the_batches_that_worked(self):
        self.SLIDES = [slide(f"s{i}", "J", [f"PDB {i}"]) for i in range(1, 8)]  # kelompok 1: s1-s5, kelompok 2: s6-s7
        replies = [json.dumps({f"s{i}": [f"PDB {i}."] for i in range(1, 6)}), "bukan json", "tetap bukan json"]
        error, converse = self.run_with(replies)
        self.assertIsInstance(error, NarrationPartial)
        self.assertEqual(sorted(error.partial), [f"s{i}" for i in range(1, 6)])
        self.assertEqual([s["ref"] for s in error.failed_slides], ["s6", "s7"])
        self.assertEqual(converse.call_count, 3)

    def test_batches_also_split_by_requested_item_count(self):
        dense = lambda ref: slide(ref, "J", [f"Baris {n}" for n in range(30)])  # batas 30 butir per slide
        self.SLIDES = [dense("s1"), dense("s2")]
        replies = [json.dumps({"s1": ["a."]}), json.dumps({"s2": ["b."]})]
        (narration, _), converse = self.run_with(replies)
        self.assertEqual((sorted(narration), converse.call_count), (["s1", "s2"], 2))

    def test_invalid_shape_is_retried_then_becomes_error(self):
        (narration, _), _ = self.run_with(["bukan json", json.dumps({"s1": ["ok 4,85"]})])
        self.assertEqual(narration["s1"], ["ok 4,85"])
        error, converse = self.run_with(["bukan json", json.dumps({"s1": []})])
        self.assertIsInstance(error, BedrockError)
        self.assertIn("tidak valid", str(error))
        self.assertEqual(converse.call_count, 2)


if __name__ == "__main__":
    unittest.main()
