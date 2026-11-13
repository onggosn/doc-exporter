import io
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image

from docx_builder import (
    BODY_FONT, DEFAULT_PAGE_SETUP, DocxBuildError, build_docx, page_setup_from_profile,
)

SETUP = {
    "lebar_pt": 595.28, "tinggi_pt": 841.89,
    "margin": {"kiri": 34.0, "kanan": 34.0, "atas": 60.0, "bawah": 58.0}, "nomor_halaman": True,
}
CONTENT_WIDTH = 595.28 - 68.0


def png(color="red", size=(300, 200)) -> bytes:
    buffer = io.BytesIO()
    Image.new("RGB", size, color).save(buffer, "PNG")
    return buffer.getvalue()


def build(blocks, setup=SETUP) -> Document:
    buffer = io.BytesIO()
    build_docx({"blocks": blocks}, setup, buffer)
    buffer.seek(0)
    return Document(buffer)


def texts(document: Document) -> list[str]:
    return [p.text for p in document.paragraphs]


class PageSetupTests(unittest.TestCase):
    PAGE = {"pdf": [
        {"ukuran": {"nama": "A4", "lebar_pt": 595.9, "tinggi_pt": 842.9},
         "margin_pt": {"kiri": 33.7, "kanan": 33.4, "atas": 61.2, "bawah": 56.7}, "nomor_halaman": {"x": 1}},
        {"ukuran": {"nama": "A4", "lebar_pt": 595.9, "tinggi_pt": 842.9},
         "margin_pt": {"kiri": 35.2, "kanan": 34.2, "atas": 58.3, "bawah": 60.3}, "nomor_halaman": {"x": 1}},
    ]}

    def test_averages_margins_and_uses_exact_a4(self):
        setup = page_setup_from_profile(self.PAGE)
        self.assertEqual((setup["lebar_pt"], setup["tinggi_pt"]), (595.28, 841.89))
        self.assertEqual(setup["margin"], {"kiri": 34.1, "kanan": 34.1, "atas": 59.8, "bawah": 58.5})
        self.assertTrue(setup["nomor_halaman"])

    def test_defaults_when_profile_has_no_pdf(self):
        for page in (None, {}, {"pdf": []}):
            self.assertEqual(page_setup_from_profile(page), DEFAULT_PAGE_SETUP)

    def test_default_margins_are_not_shared_between_calls(self):
        page_setup_from_profile(None)["margin"]["kiri"] = 0
        self.assertEqual(DEFAULT_PAGE_SETUP["margin"]["kiri"], 34.0)


class NumberedListTests(unittest.TestCase):
    def test_numbered_items_have_marker_tab_text_with_hanging_indent_per_depth(self):
        document = build([{"type": "numbered", "items": [
            {"marker": "1.", "text": "Kebijakan Suku Bunga", "depth": 0},
            {"marker": "a.", "text": "Suku bunga tetap", "depth": 1},
        ]}])
        first, second = [p for p in document.paragraphs if p.text]
        self.assertEqual((first.text, second.text), ("1.\tKebijakan Suku Bunga", "a.\tSuku bunga tetap"))
        self.assertEqual((first.paragraph_format.left_indent.pt, first.paragraph_format.first_line_indent.pt), (22, -22))
        self.assertEqual((second.paragraph_format.left_indent.pt, second.paragraph_format.first_line_indent.pt), (46, -22))
        self.assertEqual(second.paragraph_format.tab_stops[0].position.pt, 46)


class DocumentFormatTests(unittest.TestCase):
    def setUp(self):
        self.document = build([
            {"type": "heading", "level": 1, "text": "1. Judul"},
            {"type": "paragraph", "text": "Teks dengan <i>istilah asing</i> dan <b>tebal</b>."},
            {"type": "bullets", "items": ["Poin satu", "Poin dua"]},
        ])

    def test_page_size_margins_and_body_font(self):
        section = self.document.sections[0]
        self.assertAlmostEqual(section.page_width.pt, 595.28, delta=0.1)
        self.assertAlmostEqual(section.left_margin.pt, 34.0, delta=0.1)
        self.assertAlmostEqual(section.top_margin.pt, 60.0, delta=0.1)
        normal = self.document.styles["Normal"].font
        self.assertEqual((normal.name, normal.size), (BODY_FONT, Pt(11)))

    def test_heading_font_has_no_theme_override_and_is_black_bold(self):
        style = self.document.styles["Heading 1"]
        fonts = style.element.rPr.rFonts
        self.assertEqual(fonts.get(qn("w:ascii")), BODY_FONT)
        self.assertFalse([a for a in fonts.attrib if a.endswith("Theme")])
        self.assertTrue(style.font.bold)
        self.assertEqual(str(style.font.color.rgb), "000000")

    def test_inline_markup_becomes_runs_without_tags(self):
        paragraph = next(p for p in self.document.paragraphs if p.text.startswith("Teks"))
        self.assertEqual(paragraph.text, "Teks dengan istilah asing dan tebal.")
        self.assertEqual([(r.text, r.italic, r.bold) for r in paragraph.runs if r.italic or r.bold],
                         [("istilah asing", True, None), ("tebal", None, True)])

    def test_nested_and_unbalanced_markup_is_tolerated(self):
        document = build([
            {"type": "paragraph", "text": "<i>a <b>b</b> c</i> d"},
            {"type": "paragraph", "text": "x</b> <b>tebal tanpa tutup"},
        ])
        first, second = document.paragraphs[0], document.paragraphs[1]
        self.assertEqual([(r.text, r.italic, r.bold) for r in first.runs],
                         [("a ", True, None), ("b", True, True), (" c", True, None), (" d", None, None)])
        self.assertEqual(second.text, "x tebal tanpa tutup")
        self.assertTrue(second.runs[-1].bold)

    def test_control_characters_are_cleaned_instead_of_crashing(self):
        document = build([
            {"type": "heading", "level": 2, "text": "Judul\x0bbaris dua"},
            {"type": "bullets", "items": ["a\x01b\x0cc"]},
        ])
        self.assertEqual(texts(document), ["Judul baris dua", "ab c"])

    def test_indented_bullets_become_nested_list_levels_without_the_indent(self):
        document = build([{"type": "bullets", "items": ["Utama", "  Anak", "    Cucu", "Utama lagi"]}])
        self.assertEqual([(p.style.name, p.text) for p in document.paragraphs],
                         [("List Bullet", "Utama"), ("List Bullet 2", "Anak"), ("List Bullet 3", "Cucu"), ("List Bullet", "Utama lagi")])

    def test_bullets_use_list_style(self):
        bullets = [p for p in self.document.paragraphs if p.style.name == "List Bullet"]
        self.assertEqual([p.text for p in bullets], ["Poin satu", "Poin dua"])

    def test_page_number_field_is_right_aligned_in_footer(self):
        footer = self.document.sections[0].footer.paragraphs[0]
        self.assertEqual(footer.alignment, 2)
        self.assertIn('w:instr="PAGE"', footer._p.xml)

    def test_page_number_can_be_disabled(self):
        document = build([{"type": "paragraph", "text": "x"}], {**SETUP, "nomor_halaman": False})
        self.assertNotIn("PAGE", document.sections[0].footer.paragraphs[0]._p.xml)

    def test_no_watermark_is_added(self):
        buffer = io.BytesIO()
        build_docx({"blocks": [{"type": "paragraph", "text": "x"}]}, SETUP, buffer)
        with zipfile.ZipFile(buffer) as archive:
            combined = b"".join(archive.read(name) for name in archive.namelist() if name.endswith(".xml"))
        self.assertNotIn(b"RAHASIA", combined)
        self.assertNotIn(b"PowerPlusWaterMarkObject", combined)


class TableTests(unittest.TestCase):
    def table_block(self, **overrides):
        block = {
            "type": "table", "widths": [40, 60], "banner": "Pita judul",
            "rows": [
                [[{"type": "caption", "text": "PESAN UTAMA", "tipe": "headline"}, {"type": "image", "source": png()}],
                 [{"type": "bullets", "items": ["a", "b"]}]],
                [[{"type": "image", "source": png("blue")}, {"type": "caption", "text": "Grafik 1. Kredit", "tipe": "label"}],
                 [{"type": "paragraph", "text": "Narasi"}]],
                [[{"type": "image", "source": png("green")}]],
            ],
        }
        return {**block, **overrides}

    def test_column_grid_cell_widths_and_table_properties(self):
        table = build([self.table_block()]).tables[0]
        self.assertEqual([int(column.width.pt) for column in table.columns], [210, 316])
        self.assertAlmostEqual(table.rows[1].cells[0].width.pt, CONTENT_WIDTH * 0.4, delta=0.5)
        order = [child.tag.split("}")[1] for child in table._tbl.tblPr]
        self.assertEqual(order.count("tblLayout"), 1)
        self.assertLess(order.index("tblInd"), order.index("tblLayout"))

    def test_every_table_has_solid_one_point_borders_on_all_sides_and_between_cells(self):
        document = build([self.table_block(), self.table_block(banner=None, rows=[[[{"type": "paragraph", "text": "a"}], [{"type": "paragraph", "text": "b"}]]])])
        self.assertEqual(len(document.tables), 2)
        for table in document.tables:
            borders = table._tbl.tblPr.find(qn("w:tblBorders"))
            sides = {child.tag.split("}")[1]: child for child in borders}
            self.assertEqual(set(sides), {"top", "left", "bottom", "right", "insideH", "insideV"})
            for side in sides.values():
                self.assertEqual((side.get(qn("w:val")), side.get(qn("w:sz")), side.get(qn("w:color"))),
                                 ("single", "8", "000000"))  # sz dalam seperdelapan poin: 8 = 1 pt
            order = [child.tag.split("}")[1] for child in table._tbl.tblPr]
            self.assertLess(order.index("tblInd"), order.index("tblBorders"))
            self.assertLess(order.index("tblBorders"), order.index("tblLayout"))

    def test_banner_row_is_merged_shaded_bold_and_centered(self):
        table = build([self.table_block()]).tables[0]
        banner = table.rows[0].cells[0]
        self.assertIs(banner._tc, table.rows[0].cells[-1]._tc)
        self.assertIn('w:fill="E6E6E6"', banner._tc.xml)
        self.assertTrue(banner.paragraphs[0].runs[0].bold)

    def test_single_cell_row_spans_all_columns(self):
        table = build([self.table_block()]).tables[0]
        last = table.rows[3]
        self.assertIs(last.cells[0]._tc, last.cells[1]._tc)
        self.assertEqual(len(last.cells[0].paragraphs[0].runs[0]._r.xpath(".//a:blip")), 1)

    def test_headline_caption_precedes_image_and_label_follows(self):
        table = build([self.table_block()]).tables[0]
        first, second = table.rows[1].cells[0], table.rows[2].cells[0]
        self.assertEqual(first.paragraphs[0].text, "PESAN UTAMA")
        self.assertTrue(first.paragraphs[0].runs[0].bold)
        self.assertEqual(len(first.paragraphs[1].runs[0]._r.xpath(".//a:blip")), 1)
        self.assertEqual(second.paragraphs[1].text, "Grafik 1. Kredit")
        self.assertFalse(second.paragraphs[1].runs[0].bold)

    def test_image_fits_inside_cell_with_padding(self):
        document = build([self.table_block()])
        cell_width = CONTENT_WIDTH * 0.4
        widest = max(shape.width.pt for shape in list(document.inline_shapes)[:2])
        self.assertLess(widest, cell_width - 2 * 5.4)
        self.assertEqual(len(document.inline_shapes), 3)

    def test_caption_stays_with_its_image_and_rows_do_not_split(self):
        table = build([self.table_block()]).tables[0]
        headline_cell, label_cell = table.rows[1].cells[0], table.rows[2].cells[0]
        self.assertTrue(headline_cell.paragraphs[0].paragraph_format.keep_with_next)  # caption di atas gambar
        self.assertTrue(label_cell.paragraphs[0].paragraph_format.keep_with_next)  # gambar di atas caption
        self.assertFalse(label_cell.paragraphs[1].paragraph_format.keep_with_next)
        self.assertTrue(all(row._tr.trPr is not None and row._tr.trPr.find(qn("w:cantSplit")) is not None
                            for row in table.rows))

    def test_image_without_caption_is_not_kept_with_next(self):
        table = build([self.table_block()]).tables[0]
        self.assertFalse(table.rows[3].cells[0].paragraphs[0].paragraph_format.keep_with_next)

    def test_row_with_wrong_cell_count_is_rejected(self):
        bad = self.table_block(widths=[30, 30, 40])
        with self.assertRaisesRegex(DocxBuildError, "Baris tabel #1 memiliki 2 sel"):
            build([bad])

    def test_invalid_widths_and_rows_are_rejected(self):
        for bad in (self.table_block(widths=[]), self.table_block(widths=[0, 10]), self.table_block(rows=[])):
            with self.assertRaises(DocxBuildError):
                build([bad])


class FigureAndErrorTests(unittest.TestCase):
    def test_figure_places_label_below_and_headline_above(self):
        below = build([{"type": "figure", "image": png(), "caption": {"text": "Tabel 1 Sektoral", "tipe": "label"}}])
        self.assertEqual(texts(below)[-1], "Tabel 1 Sektoral")
        above = build([{"type": "figure", "image": png(), "caption": {"text": "PESAN", "tipe": "headline"}}])
        self.assertEqual(texts(above)[0], "PESAN")

    def test_figure_width_percent(self):
        document = build([{"type": "figure", "image": png(), "width_percent": 50}])
        self.assertAlmostEqual(document.inline_shapes[0].width.pt, CONTENT_WIDTH * 0.5, delta=0.5)

    def test_image_can_come_from_a_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "g.png"
            path.write_bytes(png())
            self.assertEqual(len(build([{"type": "figure", "image": path}]).inline_shapes), 1)

    def test_errors_are_clear(self):
        cases = [
            ({"blocks": [{"type": "aneh"}]}, "tipe tidak dikenal"),
            ({"blocks": [{"type": "heading", "level": 9, "text": "x"}]}, "level judul"),
            ({"blocks": [{"type": "paragraph", "text": "<script>x</script>"}]}, "Tag tidak didukung"),
            ({"blocks": [{"type": "figure", "image": "/tidak/ada.png"}]}, "tidak ditemukan"),
            ({"blocks": [{"type": "figure", "image": b"bukan gambar"}]}, "tidak dapat dibaca"),
            ({"blocks": [{"type": "figure", "image": png(), "width_percent": 150}]}, "width_percent"),
            ({"blocks": "x"}, "daftar 'blocks'"),
            ([], "daftar 'blocks'"),
        ]
        for spec, message in cases:
            with self.subTest(message=message), self.assertRaisesRegex(DocxBuildError, message):
                build_docx(spec, SETUP, io.BytesIO())


if __name__ == "__main__":
    unittest.main()
