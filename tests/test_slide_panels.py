import io
import tempfile
import unittest
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Emu

import doc_composer as dc
import slide_panels as sp

W, H = 12192000, 6858000


def picture_stream(color="teal") -> io.BytesIO:
    buffer = io.BytesIO()
    Image.new("RGB", (200, 150), color).save(buffer, "PNG")
    buffer.seek(0)
    return buffer


def frac(x, y):
    return int(x * W), int(y * H)


def build_slide(layout):
    """layout: daftar ('judul', teks, (l, t, w, h)) atau ('gambar', None, (l, t, w, h)), dalam pecahan slide."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for kind, text, (left, top, width, height) in layout:
        (x, y), (w, h) = frac(left, top), frac(width, height)
        if kind == "judul":
            box = slide.shapes.add_textbox(x, y, w, h)
            box.text_frame.text = text
        else:
            slide.shapes.add_picture(picture_stream(), x, y, w, h)
    return slide


TWO_PANELS = [
    ("judul", "PERTUMBUHAN KREDIT MELAMBAT", (0.02, 0.15, 0.45, 0.05)),
    ("gambar", None, (0.02, 0.21, 0.45, 0.6)),
    ("judul", "2. SUKU BUNGA TURUN", (0.52, 0.15, 0.45, 0.05)),
    ("gambar", None, (0.52, 0.21, 0.45, 0.6)),
]


class DetectPanelsTests(unittest.TestCase):
    def test_two_headed_pictures_become_two_panels_in_reading_order(self):
        panels = sp.detect_panels(build_slide(TWO_PANELS), W, H)
        self.assertEqual([p["header"] for p in panels], ["PERTUMBUHAN KREDIT MELAMBAT", "2. SUKU BUNGA TURUN"])
        left, top, right, bottom = panels[0]["bbox"]
        self.assertAlmostEqual(left, 0.02, delta=0.01)
        self.assertGreaterEqual(top, 0.20)  # judul panel tidak ikut terpotong
        self.assertAlmostEqual(right, 0.47, delta=0.01)

    def test_single_panel_slide_is_not_split(self):
        self.assertEqual(sp.detect_panels(build_slide(TWO_PANELS[:2]), W, H), [])

    def test_mixed_case_labels_and_small_objects_are_not_panel_headers(self):
        layout = [("judul", "Label diagram biasa", (0.02, 0.15, 0.45, 0.05)), ("gambar", None, (0.02, 0.21, 0.45, 0.6)),
                  ("gambar", None, (0.52, 0.21, 0.45, 0.6)), ("gambar", None, (0.9, 0.02, 0.04, 0.04))]
        panels = sp.detect_panels(build_slide(layout), W, H)
        self.assertEqual([p["header"] for p in panels], [None, None])  # panel tanpa judul tetap dipakai bila besar

    def test_plain_caption_over_several_objects_groups_them_into_one_panel(self):
        layout = [("judul", "Plan B : Full Response", (0.4, 0.15, 0.58, 0.05)),
                  ("gambar", None, (0.4, 0.21, 0.28, 0.3)), ("gambar", None, (0.7, 0.21, 0.28, 0.3)),
                  ("judul", "Plan A : No Response", (0.02, 0.15, 0.35, 0.05)),
                  ("gambar", None, (0.02, 0.21, 0.35, 0.3)),  # satu objek, tetapi serumpun ("Plan")
                  ("judul", "Label lain", (0.02, 0.6, 0.35, 0.05)), ("gambar", None, (0.02, 0.66, 0.35, 0.3))]
        panels = sp.detect_panels(build_slide(layout), W, H)
        by_header = {p["header"]: p["bbox"] for p in panels}
        self.assertEqual(set(by_header) - {None}, {"Plan B : Full Response", "Plan A : No Response"})
        left, _, right, _ = by_header["Plan B : Full Response"]
        self.assertTrue(left < 0.41 and right > 0.97)  # kedua gambar dalam satu kotak, tidak dipotong terpisah

    def test_panel_box_starts_below_its_header_text(self):
        layout = [("judul", "Plan B : Full Response", (0.4, 0.15, 0.58, 0.07)),
                  ("gambar", None, (0.4, 0.19, 0.28, 0.3)), ("gambar", None, (0.7, 0.19, 0.28, 0.3)),
                  ("judul", "Plan A : No Response", (0.02, 0.15, 0.35, 0.07)), ("gambar", None, (0.02, 0.2, 0.35, 0.3))]
        panels = sp.detect_panels(build_slide(layout), W, H)
        for panel in panels:  # gambar menjorok ke kotak judul 0,03, tetapi panel tidak memotong teks judulnya
            self.assertGreaterEqual(panel["bbox"][1], 0.22 - 1e-9, panel["header"])

    def test_footnotes_are_not_headers(self):
        layout = TWO_PANELS + [("judul", "KET: DATA PER MEI 2025", (0.02, 0.83, 0.45, 0.04))]
        self.assertEqual([p["header"] for p in sp.detect_panels(build_slide(layout), W, H)],
                         ["PERTUMBUHAN KREDIT MELAMBAT", "2. SUKU BUNGA TURUN"])


class DetectVerbatimTests(unittest.TestCase):
    @staticmethod
    def slide_with(shapes):
        prs = Presentation()
        prs.slide_width, prs.slide_height = W, H
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for text, (left, top, width, height) in shapes:
            (x, y), (w, h) = frac(left, top), frac(width, height)
            box = slide.shapes.add_textbox(x, y, w, h)
            box.text_frame.text = text
        return slide

    def test_body_right_below_the_label_is_captured_line_by_line(self):
        from ppt_adapter import detect_verbatim
        slide = self.slide_with([
            ("HASIL PEMBAHASAN DEWAN GUBERNUR", (0.01, 0.08, 0.27, 0.04)),
            ("Kesepakatan tarif sementara menurunkan tarif efektif.\nPDB global lebih baik menjadi 3,0%.", (0.01, 0.13, 0.67, 0.2)),
            ("Teks di tempat lain yang tidak terkait dengan label.", (0.7, 0.6, 0.29, 0.1)),
            ("Teks jauh di bawah label tetapi sejajar secara horizontal.", (0.01, 0.7, 0.5, 0.1)),
        ])
        self.assertEqual(detect_verbatim(slide, W, H), [{
            "label": "HASIL PEMBAHASAN DEWAN GUBERNUR",
            "lines": ["Kesepakatan tarif sementara menurunkan tarif efektif.", "PDB global lebih baik menjadi 3,0%."],
        }])

    def test_key_takeaways_variants_match_and_slides_without_labels_have_no_blocks(self):
        from ppt_adapter import detect_verbatim
        slide = self.slide_with([("Key Takeaway", (0.0, 0.08, 0.2, 0.04)),
                                 ("Interkoneksi SP stabil dan adopsi SNAP meningkat.", (0.03, 0.13, 0.9, 0.1))])
        self.assertEqual([b["label"] for b in detect_verbatim(slide, W, H)], ["Key Takeaway"])
        plain = self.slide_with([("Judul biasa tanpa label", (0.0, 0.08, 0.2, 0.04)), ("Isi slide yang cukup panjang.", (0.03, 0.13, 0.9, 0.1))])
        self.assertEqual(detect_verbatim(plain, W, H), [])


class CropAndComposeTests(unittest.TestCase):
    def test_crop_top_removes_the_title_band_and_panels_start_below_it(self):
        with tempfile.TemporaryDirectory() as tmp:
            png = Path(tmp) / "slide.png"
            image = Image.new("RGB", (400, 200), "white")
            image.paste(Image.new("RGB", (380, 40), "red"), (10, 0))  # pita judul di atas
            image.paste(Image.new("RGB", (150, 100), "navy"), (20, 60))
            image.paste(Image.new("RGB", (150, 100), "maroon"), (220, 60))
            image.save(png)
            panels = [{"bbox": (0.0, 0.0, 0.5, 1.0), "header": "A"}, {"bbox": (0.5, 0.0, 1.0, 1.0), "header": "B"}]
            crops = dc.crop_panels(str(png), panels, Path(tmp), "s", content_top=0.25)
            for crop in crops:
                with Image.open(crop["image_path"]) as saved:
                    self.assertNotIn((255, 0, 0), [color for _, color in saved.convert("RGB").getcolors(100000)])
            dc.crop_top(str(png), 0.25)
            with Image.open(png) as saved:
                self.assertEqual(saved.size, (400, 150))
                self.assertNotIn((255, 0, 0), [color for _, color in saved.convert("RGB").getcolors(100000)])
            dc.crop_top(str(png), 0.0)  # tanpa pita atas: tidak berubah
            with Image.open(png) as saved:
                self.assertEqual(saved.size, (400, 150))

    def test_crop_panels_writes_one_file_per_panel_and_drops_tiny_ones(self):
        with tempfile.TemporaryDirectory() as tmp:
            png = Path(tmp) / "slide.png"
            image = Image.new("RGB", (400, 225), "white")
            image.paste(Image.new("RGB", (150, 120), "navy"), (20, 40))
            image.paste(Image.new("RGB", (150, 120), "maroon"), (220, 40))
            image.save(png)
            panels = [{"bbox": (0.03, 0.15, 0.45, 0.85), "header": "A"}, {"bbox": (0.52, 0.15, 0.97, 0.85), "header": "B"},
                      {"bbox": (0.0, 0.0, 0.01, 0.01), "header": "kecil"}]
            crops = dc.crop_panels(str(png), panels, Path(tmp), "slide-5")
            self.assertEqual([c["header"] for c in crops], ["A", "B"])
            self.assertTrue(all(Path(c["image_path"]).is_file() for c in crops))
            self.assertEqual(Path(crops[0]["image_path"]).name, "slide-5-panel1.png")

    def test_panel_over_half_the_canvas_is_full_width_and_smaller_panels_stay_two_column(self):
        item = {"ref": "s1", "title": "Judul", "text": ["Isi."], "tables": [], "charts": [], "visual_desc": None,
                "vision_reasons": ["picture"], "image_path": "/tmp/whole.png",
                "panel_images": [
                    {"header": "BESAR", "bbox": (0.0, 0.0, 1.0, 0.55), "image_path": "/tmp/big.png", "summary": "Ringkasan besar."},
                    {"header": "TEPAT", "bbox": (0.0, 0.6, 0.5, 1.0), "image_path": "/tmp/edge.png", "summary": "x"},
                    {"header": "KECIL", "bbox": (0.5, 0.6, 1.0, 1.0), "image_path": "/tmp/small.png", "summary": "y"}]}
        spec, _ = dc.compose({"items": [item]}, ["Bab"], {"Bab": ["s1"]}, {}, {"penomoran": [], "caption_default": "kosong"}, [])
        blocks = spec["blocks"]
        self.assertEqual([b["type"] for b in blocks[1:]], ["image", "caption", "paragraph", "table"])  # besar di luar tabel
        self.assertEqual(blocks[1]["source"], "/tmp/big.png")
        edge, small = blocks[4]["rows"]  # 0,2 dan 0,2 dari kanvas: tetap kolom kiri tabel dua kolom
        self.assertEqual((len(edge), len(small)), (2, 2))

    def test_each_panel_is_its_own_row_with_header_and_summary_beside_the_image(self):
        item = {"ref": "s1", "title": "Judul slide", "text": ["Isi."], "tables": [], "charts": [], "notes": "",
                "visual_desc": None, "vision_reasons": ["picture"], "image_path": "/tmp/whole.png",
                "panel_images": [
                    {"header": "PANEL B", "bbox": (0.52, 0.1, 0.97, 0.5), "image_path": "/tmp/b.png", "summary": "Ringkasan B."},
                    {"header": "PANEL A", "bbox": (0.02, 0.1, 0.47, 0.5), "image_path": "/tmp/a.png", "summary": "Ringkasan A."},
                    {"header": None, "bbox": (0.02, 0.6, 0.97, 0.95), "image_path": "/tmp/c.png", "summary": ""}]}
        spec, report = dc.compose({"items": [item]}, ["Bab"], {"Bab": ["s1"]}, {"s1": ["Poin satu.", "Poin dua."]},
                                  {"penomoran": [], "caption_default": "kosong"}, [])
        table = spec["blocks"][1]
        self.assertEqual(table["type"], "table")
        rows = table["rows"]
        self.assertEqual([r[0][0]["source"] for r in rows], ["/tmp/a.png", "/tmp/b.png", "/tmp/c.png"])  # urutan baca
        self.assertEqual(rows[0][1], [{"type": "caption", "text": "PANEL A", "tipe": "headline", "align": "left"},
                                      {"type": "paragraph", "text": "Ringkasan A."}])
        self.assertEqual((len(rows[2]), rows[2][1]), (2, []))  # tanpa judul dan ringkasan: kolom kanan kosong, gambar tetap di kiri
        self.assertEqual(spec["blocks"][2:4], [{"type": "paragraph", "text": "Poin satu."}, {"type": "paragraph", "text": "Poin dua."}])
        self.assertEqual((report["gambar"], report["gambar_diringkas"], report["slide_dipecah"]), (3, 2, ["s1 (3 panel)"]))

    def test_numbering_rule_turns_panel_header_into_numbered_label(self):
        item = {"ref": "s1", "title": "", "text": ["Isi."], "tables": [], "charts": [{"title": "x", "categories": [], "series": []}],
                "notes": "", "visual_desc": None, "vision_reasons": ["chart"], "image_path": "/tmp/w.png",
                "panel_images": [{"header": "KREDIT", "bbox": (0, 0.1, 0.4, 0.5), "image_path": "/tmp/a.png"},
                                 {"header": "DPK", "bbox": (0.5, 0.6, 0.9, 0.9), "image_path": "/tmp/b.png"}]}
        rule = {"bagian_cocok": "Lampiran", "awalan": "Grafik", "mulai": 1, "lanjut": "berlanjut", "aktif": True}
        spec, _ = dc.compose({"items": [item]}, ["Lampiran Data"], {"Lampiran Data": ["s1"]}, {"s1": ["Poin."]},
                             {"penomoran": [rule], "caption_default": "kosong"}, [])
        rows = spec["blocks"][1]["rows"]
        self.assertEqual([r[1][0]["text"] for r in rows], ["Grafik 1. KREDIT", "Grafik 2. DPK"])
        self.assertEqual(rows[0][1][0]["tipe"], "label")


if __name__ == "__main__":
    unittest.main()
