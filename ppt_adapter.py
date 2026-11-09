"""Adapter PPT -> JSON ternormalisasi (tahap ingest).

Deterministik: teks, tabel, notes, dan chart diambil lewat kode. Slide yang
bersifat visual ditandai needs_vision dan (opsional) dirender ke PNG lewat
LibreOffice supaya bisa dideskripsikan model multimodal.

Pemakaian:
    python ppt_adapter.py deck.pptx --source-id ppt-001 --render-dir out/png > deck.json

Dependensi: pip install python-pptx ; (opsional) LibreOffice + pdftoppm (poppler).
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Callable, Optional

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

from slide_panels import _box, detect_panels

# Slide dianggap "visual" bila ada gambar/chart/SmartArt, atau teksnya sangat sedikit.
MIN_TEXT_CHARS = 40

TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}


def _walk(shapes):
    """Iterasi semua shape, masuk ke dalam group secara rekursif."""
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk(sh.shapes)
        else:
            yield sh


def _is_title(sh) -> bool:
    return sh.is_placeholder and sh.placeholder_format.type in TITLE_TYPES


def _paragraph_lines(text_frame) -> list[str]:
    lines = []
    for p in text_frame.paragraphs:
        t = "".join(r.text for r in p.runs).strip()
        if t:
            # pertahankan hirarki poin dengan indentasi level
            lines.append(("  " * p.level) + t if p.level else t)
    return lines


def _table_rows(table) -> list[list[str]]:
    return [[c.text.strip() for c in row.cells] for row in table.rows]


def _chart_data(chart) -> dict:
    """Ambil data numerik chart langsung dari XML-nya (lebih akurat daripada vision)."""
    try:
        cats = [str(c) for c in chart.plots[0].categories]
        series = [
            {"name": s.name, "values": [None if v is None else float(v) for v in s.values]}
            for s in chart.plots[0].series
        ]
        title = chart.chart_title.text_frame.text if chart.has_title else ""
        return {"title": title, "categories": cats, "series": series}
    except Exception:
        return {}


# Label bagian yang isinya harus dibawa apa adanya ke dokumen (tidak diringkas model).
VERBATIM_LABELS = (
    re.compile(r"hasil\s+pembahasan\s+dewan\s+gubernur", re.I),
    re.compile(r"key\s*take\s*aways?", re.I),
)
VERBATIM_LABEL_MAX_CHARS = 60
VERBATIM_BODY_MIN_CHARS = 20
VERBATIM_ABOVE = 0.03  # isi boleh sedikit menumpuk di atas dasar label
VERBATIM_BELOW = 0.06  # dan paling jauh sekian di bawahnya


def detect_verbatim(slide, width: int, height: int) -> list[dict]:
    """Blok teks yang diberi label "Hasil Pembahasan Dewan Gubernur" atau "Key Takeaways": [{'label', 'lines'}].
    Isinya adalah bentuk teks tepat di bawah label dan sejajar dengannya; barisnya persis seperti di slide."""
    shapes = [(sh, _box(sh, width, height)) for sh in slide.shapes]
    shapes = [(sh, box) for sh, box in shapes if box is not None and sh.has_text_frame and sh.text_frame.text.strip()]
    blocks = []
    for label_shape, label_box in shapes:
        label = " ".join(label_shape.text_frame.text.split())
        if len(label) > VERBATIM_LABEL_MAX_CHARS or not any(p.search(label) for p in VERBATIM_LABELS):
            continue
        bodies = []
        for shape, box in shapes:
            text = " ".join(shape.text_frame.text.split())
            if (shape is label_shape or _is_title(shape) or len(text) < VERBATIM_BODY_MIN_CHARS
                    or any(p.search(text) for p in VERBATIM_LABELS)):
                continue
            if (label_box[3] - VERBATIM_ABOVE <= box[1] <= label_box[3] + VERBATIM_BELOW
                    and box[0] <= label_box[2] and box[2] >= label_box[0]):
                bodies.append((round(box[1], 2), box[0], _paragraph_lines(shape.text_frame)))
        lines = [line for _top, _left, group in sorted(bodies, key=lambda b: (b[0], b[1])) for line in group]
        if lines:
            blocks.append({"label": label, "lines": lines})
    return blocks


def extract_slide(slide, index: int) -> dict:
    title = ""
    text: list[str] = []
    tables: list[list[list[str]]] = []
    charts: list[dict] = []
    n_pictures = 0
    has_other_visual = False  # SmartArt, OLE, dsb.

    for sh in _walk(slide.shapes):
        if sh.has_text_frame and sh.text_frame.text.strip():
            if _is_title(sh) and not title:
                title = sh.text_frame.text.strip()
            else:
                text.extend(_paragraph_lines(sh.text_frame))
        elif getattr(sh, "has_table", False) and sh.has_table:
            tables.append(_table_rows(sh.table))
        elif getattr(sh, "has_chart", False) and sh.has_chart:
            charts.append(_chart_data(sh.chart))
        elif sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            n_pictures += 1
        elif sh.shape_type in (MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT, MSO_SHAPE_TYPE.LINKED_OLE_OBJECT):
            has_other_visual = True
        elif sh.element.tag.endswith("graphicFrame"):  # SmartArt/diagram
            has_other_visual = True

    # fallback judul: placeholder judul kosong -> pakai title bawaan python-pptx
    if not title and slide.shapes.title is not None:
        title = slide.shapes.title.text_frame.text.strip()

    notes = ""
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()

    total_chars = len(title) + sum(len(t) for t in text)
    reasons = []
    if n_pictures:
        reasons.append("picture")
    if has_other_visual:
        reasons.append("diagram_or_object")
    if total_chars < MIN_TEXT_CHARS and not tables and not charts:
        reasons.append("low_text")
    # chart sudah punya data numerik dari XML, tapi tetap butuh vision untuk tafsir visual

    return {
        "ref": f"slide-{index}",
        "title": title,
        "text": text,
        "tables": tables,
        "charts": charts,
        "visual_desc": None,
        "notes": notes,
        "needs_vision": bool(reasons or charts),
        "vision_reasons": reasons + (["chart"] if charts else []),
    }


def render_slides_png(pptx_path: Path, out_dir: Path, dpi: int = 110) -> dict[int, Path]:
    """PPTX -> PDF (LibreOffice) -> PNG per slide (pdftoppm). Return {nomor_slide: path}."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    pdftoppm = shutil.which("pdftoppm")
    if not soffice or not pdftoppm:
        raise RuntimeError("Butuh LibreOffice (soffice) dan poppler (pdftoppm) untuk render PNG")

    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        # Profil LibreOffice per-konversi di direktori temp: aman untuk container
        # non-root (OpenShift) dan mencegah konflik lock antar konversi paralel.
        profile_dir = Path(tmp) / "lo_profile"
        subprocess.run(
            [soffice, f"-env:UserInstallation={profile_dir.as_uri()}",
             "--headless", "--convert-to", "pdf", "--outdir", tmp, str(pptx_path)],
            check=True, capture_output=True, timeout=300,
        )
        pdf = Path(tmp) / (pptx_path.stem + ".pdf")
        prefix = out_dir / pptx_path.stem
        subprocess.run(
            [pdftoppm, "-png", "-r", str(dpi), str(pdf), str(prefix)],
            check=True, capture_output=True, timeout=300,
        )

    result = {}
    for png in sorted(out_dir.glob(f"{pptx_path.stem}-*.png")):
        result[int(png.stem.rsplit("-", 1)[1])] = png  # pdftoppm menambah nomor halaman
    return result


def ingest_pptx(
    path: str | Path,
    source_id: str,
    render_dir: Optional[str | Path] = None,
    describe: Optional[Callable[[Path, dict], str]] = None,
) -> dict:
    """Baca PPTX -> dict sesuai skema normalisasi.

    render_dir : bila diisi, slide needs_vision dirender ke PNG (field image_path).
    describe   : callback (png_path, item) -> deskripsi teks dari model multimodal.
                 Dipasang oleh orkestrator; adapter sendiri tidak memanggil AI.
    """
    path = Path(path)
    prs = Presentation(str(path))
    items = [extract_slide(s, i) for i, s in enumerate(prs.slides, start=1)]
    for item, slide in zip(items, prs.slides):
        try:  # kotak panel dan blok teks asli hanya pelengkap; kegagalan deteksi tidak boleh menggagalkan ingest
            item["panels"] = detect_panels(slide, prs.slide_width, prs.slide_height)
        except Exception:
            item["panels"] = []
        try:
            item["verbatim"] = detect_verbatim(slide, prs.slide_width, prs.slide_height)
        except Exception:
            item["verbatim"] = []

    if render_dir and any(it["needs_vision"] for it in items):
        pngs = render_slides_png(path, Path(render_dir))
        for i, it in enumerate(items, start=1):
            if it["needs_vision"] and i in pngs:
                it["image_path"] = str(pngs[i])
                if describe:
                    it["visual_desc"] = describe(pngs[i], it)

    return {
        "source_id": source_id,
        "type": "pptx",
        "filename": path.name,
        "items": items,
    }


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Ingest PPTX ke JSON ternormalisasi")
    ap.add_argument("pptx")
    ap.add_argument("--source-id", default="ppt-001")
    ap.add_argument("--render-dir", help="Folder output PNG untuk slide visual")
    args = ap.parse_args(argv)

    doc = ingest_pptx(args.pptx, args.source_id, args.render_dir)
    json.dump(doc, sys.stdout, ensure_ascii=False, indent=2)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
