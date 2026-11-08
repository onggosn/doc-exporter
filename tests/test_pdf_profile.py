import json
import tempfile
import unittest
from pathlib import Path

from pdf_profile import analyze_pdf

A4 = (595.276, 841.89)


def make_pdf(page_streams: list[str], size: tuple[float, float] = A4) -> bytes:
    """PDF minimal berfont Helvetica; tiap string adalah isi content stream satu halaman."""
    objects: list[str] = []
    page_ids = [4 + 2 * i for i in range(len(page_streams))]
    objects.append("<< /Type /Catalog /Pages 2 0 R >>")
    kids = " ".join(f"{pid} 0 R" for pid in page_ids)
    objects.append(f"<< /Type /Pages /Kids [{kids}] /Count {len(page_ids)} >>")
    objects.append("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    for pid, stream in zip(page_ids, page_streams):
        objects.append(
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {size[0]} {size[1]}] "
            f"/Resources << /Font << /F1 3 0 R >> >> /Contents {pid + 1} 0 R >>"
        )
        objects.append(f"<< /Length {len(stream)} >>\nstream\n{stream}\nendstream")
    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for number, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += f"{number} 0 obj\n{body}\nendobj\n".encode()
    xref = len(out)
    out += f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode()
    for offset in offsets:
        out += f"{offset:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    return bytes(out)


def text(x: float, y: float, size: float, content: str, matrix: str = "1 0 0 1") -> str:
    return f"BT /F1 {size} Tf {matrix} {x} {y} Tm ({content}) Tj ET"


def page(number: int) -> str:
    return "\n".join([
        text(40, 780, 11, "Isi dokumen halaman"),
        text(40, 120, 11, "Baris terakhir isi"),
        "0.5 g",
        text(150, 250, 80, "RAHASIA", "0.707 0.707 -0.707 0.707"),
        "0 g",
        text(550, 47, 10, str(number)),
        text(40, 24, 10, "Approved by Budi Santoso at 05/06/2025 10:49"),
    ])


class AnalyzePdfTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        path = Path(tmp.name) / "contoh.pdf"
        path.write_bytes(make_pdf([page(1), page(2), page(3)]))
        self.result = analyze_pdf(path)

    def test_paper_size_is_recognised_as_a4(self):
        self.assertEqual((self.result["halaman"], self.result["ukuran"]["nama"]), (3, "A4"))

    def test_margins_are_measured_from_body_text_only(self):
        margin = self.result["margin_pt"]
        self.assertAlmostEqual(margin["kiri"], 40, delta=1)
        # Footer dan watermark tidak boleh menggeser batas bawah isi.
        self.assertGreater(margin["bawah"], 100)
        self.assertAlmostEqual(margin["atas"], 52, delta=8)

    def test_page_number_is_detected_bottom_right(self):
        number = self.result["nomor_halaman"]
        self.assertEqual((number["halaman_bernomor"], number["dari_halaman"], number["mulai"]), (3, 3, 1))
        self.assertEqual(number["perataan"], "kanan")

    def test_approval_stamp_is_excluded_and_name_is_not_stored(self):
        footers = self.result["footer_berulang"]
        self.assertEqual([(f["pola"], f["dikecualikan"]) for f in footers],
                         [("Approved by * at #/#/# #:#", True)])
        self.assertNotIn("Budi", json.dumps(self.result))

    def test_diagonal_watermark_is_found(self):
        watermark = self.result["watermark"]
        self.assertEqual((watermark["teks"], watermark["sudut_derajat"], watermark["halaman"]), ("RAHASIA", 45, 3))
        self.assertEqual(watermark["warna"], "#808080")

    def test_dominant_body_font_ignores_watermark_and_footer(self):
        body = self.result["tipografi_pdf"]["isi_dominan"]
        self.assertEqual((body["font"], body["ukuran_pt"]), ("Helvetica", 11.0))


class NoWatermarkTests(unittest.TestCase):
    def test_plain_pdf_has_no_watermark_or_page_number(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "polos.pdf"
            path.write_bytes(make_pdf([text(40, 780, 11, "Halaman polos")]))
            result = analyze_pdf(path)
        self.assertIsNone(result["watermark"])
        self.assertIsNone(result["nomor_halaman"])
        self.assertEqual(result["footer_berulang"], [])


if __name__ == "__main__":
    unittest.main()
