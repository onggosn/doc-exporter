import unittest

from pptx import Presentation
from pptx.util import Inches

from ppt_adapter import extract_slide


class CanvasExtractionTests(unittest.TestCase):
    def setUp(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(10), Inches(7.5)
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def add_text(self, text, left, top):
        self.slide.shapes.add_textbox(Inches(left), Inches(top), Inches(2), Inches(1)).text_frame.text = text

    def extract(self):
        return extract_slide(self.slide, 1, self.prs.slide_width, self.prs.slide_height)

    def test_text_outside_canvas_is_ignored(self):
        self.add_text("di dalam kanvas", 1, 1)
        self.add_text("di kanan kanvas", 11, 1)
        self.add_text("di bawah kanvas", 1, 8)
        self.add_text("di kiri kanvas", -3, 1)
        text = self.extract()["text"]
        self.assertEqual(text, ["di dalam kanvas"])

    def test_partially_visible_shape_is_kept(self):
        self.add_text("separuh terlihat", 9, 1)
        self.assertEqual(self.extract()["text"], ["separuh terlihat"])

    def test_speaker_notes_are_not_extracted(self):
        self.add_text("isi slide", 1, 1)
        self.slide.notes_slide.notes_text_frame.text = "catatan pembicara"
        item = self.extract()
        self.assertNotIn("notes", item)
        self.assertNotIn("catatan pembicara", str(item))

    def add_picture(self, left, top, width, height):
        import io
        from PIL import Image
        buffer = io.BytesIO()
        Image.new("RGB", (4, 4), "red").save(buffer, "PNG")
        buffer.seek(0)
        self.slide.shapes.add_picture(buffer, Inches(left), Inches(top), Inches(width), Inches(height))

    def test_visual_area_is_the_largest_object_as_fraction_of_the_canvas(self):
        self.add_picture(0, 0, 5, 3.75)  # seperempat kanvas
        self.add_picture(0, 3.75, 10, 3.75)  # separuh kanvas
        self.assertAlmostEqual(self.extract()["visual_area"], 0.5, places=2)

    def test_visual_area_ignores_text_and_off_canvas_objects(self):
        self.add_text("hanya teks", 1, 1)
        self.add_picture(11, 1, 5, 5)
        self.assertEqual(self.extract()["visual_area"], 0.0)

    def test_content_top_is_below_title_and_top_placeholders_only(self):
        slide = self.prs.slides.add_slide(self.prs.slide_layouts[5])  # judul saja
        slide.shapes.title.text_frame.text = "Judul slide"
        slide.shapes.title.left, slide.shapes.title.top = Inches(0.5), Inches(0.3)
        slide.shapes.title.width, slide.shapes.title.height = Inches(9), Inches(1)
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1.2), Inches(3), Inches(0.5))  # bukan placeholder
        textbox.text_frame.text = "Sub judul biasa"
        item = extract_slide(slide, 1, self.prs.slide_width, self.prs.slide_height)
        self.assertAlmostEqual(item["content_top"], 1.3 / 7.5, places=2)

    def test_slide_without_top_placeholder_has_zero_content_top(self):
        self.add_text("kotak teks biasa", 1, 0.2)
        self.assertEqual(self.extract()["content_top"], 0.0)

    def test_panel_facts_hold_the_table_and_text_inside_the_panel_but_not_its_header(self):
        from ppt_adapter import panel_facts
        self.add_text("Plan B : Full Response", 0.5, 0.2)
        self.slide.shapes.add_table(2, 2, Inches(0.5), Inches(1.5), Inches(4), Inches(2))
        self.add_text("teks di luar panel", 6, 6)
        panel = {"bbox": (0.0, 0.18, 0.6, 0.6), "header": "Plan B : Full Response"}
        panel_facts(self.slide, [panel], self.prs.slide_width, self.prs.slide_height)
        self.assertEqual(len(panel["facts"]["tables"]), 1)
        self.assertEqual(panel["facts"]["text"], [])  # judul panel dan teks di luar kotak tidak ikut


class ListNumberingTests(unittest.TestCase):
    NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

    def frame(self, paragraphs):
        """paragraphs: [(teks, level, marL, indent, tipe_nomor | None, startAt | None)]"""
        from lxml import etree
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        frame = slide.shapes.add_textbox(0, 0, Inches(5), Inches(3)).text_frame
        for index, (text, level, mar_l, indent, kind, start) in enumerate(paragraphs):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            paragraph.add_run().text = text
            paragraph.level = level
            p_pr = paragraph._p.get_or_add_pPr()
            if mar_l is not None:
                p_pr.set("marL", str(mar_l))
                p_pr.set("indent", str(indent))
            if kind:
                auto = etree.SubElement(p_pr, "{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum")
                auto.set("type", kind)
                if start:
                    auto.set("startAt", str(start))
        return frame

    def test_auto_numbers_follow_the_slide_with_indentation_from_margins(self):
        from ppt_adapter import _paragraph_lines
        frame = self.frame([
            ("Kebijakan Suku Bunga", 2, 349250, -349250, "arabicPeriod", None),
            ("Suku bunga tetap", 1, 719138, -363538, "alphaLcPeriod", None),
            ("Deposit Facility", 1, 719138, -363538, "alphaLcPeriod", None),
            ("Kebijakan Operasi Moneter", 2, 0, 0, None, None),
            ("Implementasi", 1, 623888, -274638, "alphaLcPeriod", None),
        ])
        self.assertEqual(_paragraph_lines(frame), [
            "1. Kebijakan Suku Bunga", "  a. Suku bunga tetap", "  b. Deposit Facility",
            "Kebijakan Operasi Moneter", "  a. Implementasi"])  # urutan a. mulai lagi di bawah induk berikutnya

    def test_number_formats_and_start_at(self):
        from ppt_adapter import _format_number
        self.assertEqual([_format_number(k, n) for k, n in (
            ("arabicPeriod", 3), ("arabicParenR", 2), ("arabicParenBoth", 4), ("alphaUcPeriod", 28),
            ("romanLcPeriod", 4), ("romanUcParenR", 9), ("arabicPlain", 5))],
            ["3.", "2)", "(4)", "AB.", "iv.", "IX)", "5"])
        from ppt_adapter import _paragraph_lines
        frame = self.frame([("Butir", 0, 342900, -342900, "arabicPeriod", 5), ("Butir lagi", 0, 342900, -342900, "arabicPeriod", 5)])
        self.assertEqual(_paragraph_lines(frame), ["5. Butir", "6. Butir lagi"])

    def test_numbered_shape_is_a_verbatim_block_and_plain_shape_is_not(self):
        from ppt_adapter import detect_verbatim
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        numbered = slide.shapes.add_textbox(0, 0, Inches(5), Inches(1))
        numbered.text_frame.text = "Butir bernomor"
        from lxml import etree
        auto = etree.SubElement(numbered.text_frame.paragraphs[0]._p.get_or_add_pPr(),
                                "{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum")
        auto.set("type", "arabicPeriod")
        slide.shapes.add_textbox(0, Inches(2), Inches(5), Inches(1)).text_frame.text = "Teks biasa"
        blocks = detect_verbatim(slide, prs.slide_width, prs.slide_height)
        self.assertEqual(blocks, [{"label": "", "lines": ["1. Butir bernomor"]}])


class SectionTests(unittest.TestCase):
    def test_sections_map_slide_ids_to_refs_in_deck_order(self):
        from lxml import etree
        from ppt_adapter import P14_NS, read_sections
        prs = Presentation()
        for _ in range(4):
            prs.slides.add_slide(prs.slide_layouts[6])
        ids = [node.get("id") for node in prs.part._element.xpath("p:sldIdLst/p:sldId")]
        xml = (f'<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:ext uri="{{521415D9}}">'
               f'<p14:sectionLst xmlns:p14="{P14_NS}">'
               f'<p14:section name="Moneter" id="{{A}}"><p14:sldIdLst><p14:sldId id="{ids[0]}"/><p14:sldId id="{ids[1]}"/></p14:sldIdLst></p14:section>'
               f'<p14:section name="Lampiran" id="{{B}}"><p14:sldIdLst><p14:sldId id="{ids[2]}"/><p14:sldId id="{ids[3]}"/></p14:sldIdLst></p14:section>'
               f'</p14:sectionLst></p:ext></p:extLst>')
        prs.part._element.append(etree.fromstring(xml))
        self.assertEqual(read_sections(prs), [{"nama": "Moneter", "refs": ["slide-1", "slide-2"]},
                                              {"nama": "Lampiran", "refs": ["slide-3", "slide-4"]}])

    def test_deck_without_sections_gives_empty_list(self):
        from ppt_adapter import read_sections
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[6])
        self.assertEqual(read_sections(prs), [])


if __name__ == "__main__":
    unittest.main()
