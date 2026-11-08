import unittest

from profile_builder import _print_scale, analyze_html, classify_caption, html_typography

SAMPLE = r"""
<div class=\"title-level-1\"><p><strong>2. Ringkasan</strong></p></div>
<div class=\"title-level-2\"><p><strong>2.1 Asesmen</strong></p></div>
<table>
  <colgroup><col style=\"width:41.47%;\"><col style=\"width:58.53%;\"></colgroup>
  <tbody>
    <tr><td colspan=\"2\" style=\"background-color:hsl(0, 0%, 90%);\"><strong>Banner</strong></td></tr>
    <tr>
      <td><p>PROYEKSI INFLASI 2025 BASELINE DAN RISIKO</p>
          <figure class=\"image\"><img src=\"a.png\"></figure></td>
      <td>{narasi}</td>
    </tr>
  </tbody>
</table>
<figure class=\"image\"><img src=\"b.png\"><figcaption>Grafik&nbsp;3. Kredit UMKM</figcaption></figure>
""".replace("{narasi}", "Narasi penjelasan. " * 20)


class ClassifyCaptionTests(unittest.TestCase):
    def test_empty_caption(self):
        self.assertEqual(classify_caption("  \xa0 ")["type"], "kosong")

    def test_numbered_caption_with_source_prefix(self):
        result = classify_caption("Sumber: SIPP, diolah Grafik 9 . Pertumbuhan Paylater")
        self.assertEqual(result["type"], "label_bernomor")
        self.assertEqual((result["kind"], result["number"]), ("Grafik", 9))
        self.assertEqual(result["sumber_data"], "SIPP, diolah")
        self.assertEqual(result["label"], "Pertumbuhan Paylater")

    def test_uppercase_sentence_is_headline(self):
        result = classify_caption("PROYEKSI INFLASI 2025 BASELINE DAN RISIKO")
        self.assertEqual((result["type"], result["style"]), ("headline", "kapital"))

    def test_numbered_caption_exposes_uniform_numbering(self):
        result = classify_caption("Grafik 3. Kredit UMKM")
        self.assertEqual(result["penomoran"], {"awalan": "Grafik", "nomor": 3, "pemisah": "."})

    def test_prefixed_headline_keeps_type_and_records_numbering(self):
        result = classify_caption("K2.Perang Timur Tengah memperlebar defisit APBN 2026")
        self.assertEqual(result["type"], "headline")
        self.assertEqual(result["penomoran"], {"awalan": "K", "nomor": 2, "pemisah": "."})
        self.assertTrue(result["teks_tanpa_nomor"].startswith("Perang Timur Tengah"))

    def test_any_prefix_letter_is_numbering_not_only_k(self):
        result = classify_caption("P3) Likuiditas perbankan tetap memadai untuk mendukung kredit")
        self.assertEqual((result["penomoran"]["awalan"], result["penomoran"]["nomor"]), ("P", 3))

    def test_year_after_acronym_is_not_numbering(self):
        for text in ("BI 2025: Proyeksi pertumbuhan ekonomi dan inflasi", "PDB 2025 diprakirakan lebih rendah"):
            self.assertNotIn("penomoran", classify_caption(text))

    def test_unnumbered_caption_has_no_numbering(self):
        self.assertNotIn("penomoran", classify_caption("Komunikasi proyeksi"))

    def test_short_phrase_is_topic_label(self):
        self.assertEqual(classify_caption("Komunikasi proyeksi")["type"], "label_topik")


class AnalyzeHtmlTests(unittest.TestCase):
    def setUp(self):
        self.items = analyze_html(SAMPLE, "contoh.html")

    def test_reads_escaped_export_and_section_path(self):
        self.assertEqual(len(self.items), 2)
        self.assertEqual(self.items[0]["section"], "2. Ringkasan > 2.1 Asesmen")

    def test_table_layout_with_banner_and_image_left(self):
        table = self.items[0]
        self.assertEqual(table["name"], "banner+gambar-kiri-narasi-kanan")
        self.assertEqual(table["widths"], [41, 59])
        self.assertEqual(table["width_bucket"], [40, 60])
        self.assertEqual(table["rows"], ["T*2s", "I,T"])

    def test_long_neighbour_text_is_narrative_not_caption(self):
        table = self.items[0]
        self.assertEqual(table["captions"][0]["type"], "headline")
        self.assertEqual(table["captions"][0]["source"], "teks_sel_gambar")
        self.assertEqual(len(table["narasi_lengths"]), 1)

    def test_standalone_figure_uses_figcaption(self):
        figure = self.items[1]
        self.assertEqual(figure["name"], "gambar-tunggal")
        self.assertEqual(figure["captions"][0]["number"], 3)

    def test_caption_location_is_not_overwritten_by_data_source(self):
        html = '<figure class="image"><img src="x.png"><figcaption>Sumber: BI, diolah Grafik 2. Kredit</figcaption></figure>'
        caption = analyze_html(html)[0]["captions"][0]
        self.assertEqual((caption["source"], caption["sumber_data"]), ("figcaption", "BI, diolah"))


def _table(widths: str, rows: str) -> str:
    cols = "".join(f'<col style="width:{w}%;">' for w in widths.split("/"))
    return f"<table><colgroup>{cols}</colgroup><tbody>{rows}</tbody></table>"


IMG = '<figure class="image"><img src="x.png"></figure>'


class NormalizationTests(unittest.TestCase):
    def test_table_wrapped_in_figure_table_keeps_text_and_is_not_double_counted(self):
        html = f'<figure class="table">{_table("50/50", f"<tr><td>{IMG}</td><td>Narasi di sel kanan.</td></tr>")}</figure>'
        items = analyze_html(html, "x.html")
        self.assertEqual([i["name"] for i in items], ["gambar-kiri-narasi-kanan"])
        self.assertEqual(items[0]["rows"], ["I,T"])
        self.assertEqual(items[0]["images"], 1)

    def test_similar_widths_share_signature(self):
        row = f"<tr><td>{IMG}</td><td>Teks</td></tr>"
        a = analyze_html(_table("41/59", row))[0]
        b = analyze_html(_table("43/57", row))[0]
        self.assertEqual(a["signature"], b["signature"])

    def test_repeated_rows_fold_and_blank_spacers_are_dropped(self):
        row = f"<tr><td>{IMG}</td><td>Teks</td></tr>"
        spacer = "<tr><td></td><td></td></tr>"
        item = analyze_html(_table("50/50", row + spacer + row + spacer + row))[0]
        self.assertEqual((item["rows"], item["repeats"], item["row_count"]), (["I,T"], [3], 5))

    def test_row_roles_distinguish_side_by_side_and_image_only(self):
        side = f"<tr><td>{IMG}</td><td>{IMG}</td></tr>"
        alone = f"<tr><td>{IMG}</td><td></td></tr>"
        self.assertEqual(analyze_html(_table("50/50", side))[0]["name"], "gambar-berdampingan")
        self.assertEqual(analyze_html(_table("50/50", alone))[0]["name"], "gambar-saja")

    def test_many_distinct_roles_become_komposit(self):
        rows = (
            f'<tr><td colspan="2">{IMG}</td></tr>'
            f"<tr><td>{IMG}</td><td>Teks</td></tr>"
            f"<tr><td>{IMG}</td><td>{IMG}</td></tr>"
            f"<tr><td>{IMG}</td><td></td></tr>"
        )
        self.assertEqual(analyze_html(_table("50/50", rows))[0]["name"], "komposit")

    def test_inline_image_in_paragraph_is_counted(self):
        items = analyze_html('<p>Teks <img src="x.png"> lanjut</p>')
        self.assertEqual([(i["name"], i["images"]) for i in items], [("gambar-sebaris", 1)])


class TypographyTests(unittest.TestCase):
    def test_declared_size_and_font_are_weighted_by_text_length(self):
        html = (
            '<p><span style="font-family:\'Frutiger 45 Light\';font-size:11pt;">Teks panjang sekali di sini</span>'
            '<span style="font-size:10pt;">pendek</span></p>'
        )
        result = html_typography(html, "x.html")
        self.assertEqual(result["ukuran"][0]["pt"], 11.0)
        self.assertGreater(result["ukuran"][0]["bagian"], result["ukuran"][1]["bagian"])
        self.assertEqual(result["font"], [{"nama": "Frutiger 45 Light", "bagian": 1.0}])

    def test_print_scale_pairs_pdf_and_html_by_period(self):
        pdfs = [{"file": "RDGB_Mei 2025_Buku.pdf", "tipografi_pdf": {"isi_dominan": {"ukuran_pt": 7.3}}}]
        typography = [
            {"file": "DKEM RDGB Maret 2026.html", "ukuran": [{"pt": 11.0, "bagian": 1.0}]},
            {"file": "DKEM RDGB Mei 2025_06.html", "ukuran": [{"pt": 11.0, "bagian": 1.0}]},
        ]
        scale = _print_scale(pdfs, typography)
        self.assertEqual([(s["dokumen"], s["html"], s["skala"]) for s in scale],
                         [("Mei 2025", "DKEM RDGB Mei 2025_06.html", 0.66)])

    def test_print_scale_skips_pdf_without_matching_html(self):
        pdfs = [{"file": "tanpa periode.pdf", "tipografi_pdf": {"isi_dominan": {"ukuran_pt": 7.0}}}]
        self.assertEqual(_print_scale(pdfs, [{"file": "a.html", "ukuran": [{"pt": 11.0, "bagian": 1.0}]}]), [])


if __name__ == "__main__":
    unittest.main()
