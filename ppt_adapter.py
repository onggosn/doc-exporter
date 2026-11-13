"""Adapter PPT -> JSON ternormalisasi (tahap ingest).

Deterministik: teks, tabel, dan chart diambil lewat kode. Slide yang
bersifat visual ditandai needs_vision dan (opsional) dirender ke PNG lewat
LibreOffice supaya bisa dideskripsikan model multimodal.

Pemakaian:
    python ppt_adapter.py deck.pptx --source-id ppt-001 --render-dir out/png > deck.json

Dependensi: pip install python-pptx ; (opsional) LibreOffice + pymupdf.
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
from pptx.oxml.ns import qn

from slide_panels import _area, _box, _has_visual_content, detect_panels

# Slide dianggap "visual" bila ada gambar/chart/SmartArt, atau teksnya sangat sedikit.
MIN_TEXT_CHARS = 40

TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
NON_HEADING_TYPES = {PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.HEADER}


def _walk(shapes):
    """Iterasi semua shape, masuk ke dalam group secara rekursif."""
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk(sh.shapes)
        else:
            yield sh


def _on_canvas(shape, width: Optional[int], height: Optional[int]) -> bool:
    """Shape berada di dalam kanvas slide. Bentuk yang seluruhnya di luar kanvas (area kerja PowerPoint)
    tidak tampil di slide sehingga tidak diambil. Posisi tidak diketahui (diwarisi layout) dianggap di dalam."""
    if not width or not height:
        return True
    try:
        left, top, w, h = shape.left, shape.top, shape.width, shape.height
    except Exception:
        return True
    if None in (left, top, w, h):
        return True
    return left < width and top < height and left + w > 0 and top + h > 0


def _is_title(sh) -> bool:
    return sh.is_placeholder and sh.placeholder_format.type in TITLE_TYPES


LIST_DEPTH_STEP_PT = 24  # selisih posisi penanda antar tingkat bersarang
MAX_LIST_DEPTH = 2
EMU_PER_PT = 12700


def _alpha(number: int) -> str:
    letters = ""
    while number > 0:
        number, rest = divmod(number - 1, 26)
        letters = chr(ord("a") + rest) + letters
    return letters


def _roman(number: int) -> str:
    result = ""
    for value, symbol in ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"),
                          (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")):
        while number >= value:
            result, number = result + symbol, number - value
    return result


def _format_number(kind: str, number: int) -> str:
    """Penanda urutan otomatis PowerPoint (buAutoNum type=...) sebagai teks, mis. arabicPeriod 3 -> '3.'."""
    body = (
        _alpha(number) if kind.startswith("alpha") else _roman(number) if kind.startswith("roman") else str(number)
    )
    if kind.startswith(("alphaU", "romanU")):
        body = body.upper()
    if kind.endswith("ParenBoth"):
        return f"({body})"
    if kind.endswith("ParenR"):
        return f"{body})"
    if kind.endswith("Plain"):
        return body
    return f"{body}."


def _depth_from_margins(p_pr) -> Optional[int]:
    """Tingkat bersarang dari posisi penanda (marL + indent) bila paragraf punya margin eksplisit."""
    if p_pr is None or p_pr.get("marL") is None:
        return None
    try:
        marker_pt = (int(p_pr.get("marL")) + int(p_pr.get("indent", 0))) / EMU_PER_PT
    except ValueError:
        return None
    return max(0, min(int(marker_pt // LIST_DEPTH_STEP_PT), MAX_LIST_DEPTH))


def _paragraph_lines(text_frame) -> list[str]:
    """Baris teks per paragraf. Paragraf bernomor otomatis diberi penandanya ("1.", "a.", "ii)") seperti tampilan
    di PowerPoint; tingkat bersarang diberi indentasi dua spasi per tingkat, dari margin paragraf bila ada,
    selain itu dari level paragraf."""
    lines: list[str] = []
    counters: dict[int, tuple[str, int]] = {}
    for p in text_frame.paragraphs:
        text = "".join(r.text for r in p.runs).strip()
        if not text:
            continue
        p_pr = p._p.pPr
        auto = p_pr.find(qn("a:buAutoNum")) if p_pr is not None else None
        depth = _depth_from_margins(p_pr)
        depth = p.level if depth is None else depth
        # Urutan di tingkat lebih dalam mulai lagi di bawah butir berikutnya yang lebih dangkal. Yang diukur
        # adalah kedalaman visual (margin), karena level paragraf pada slide bisa terbalik terhadap induknya.
        for deeper in [k for k in counters if k > depth]:
            del counters[deeper]
        if auto is not None:
            kind = auto.get("type", "arabicPeriod")
            previous = counters.get(depth)
            number = previous[1] + 1 if previous and previous[0] == kind else int(auto.get("startAt", "1"))
            counters[depth] = (kind, number)
            text = f"{_format_number(kind, number)} {text}"
        else:
            counters.pop(depth, None)
        lines.append("  " * depth + text)
    return lines


def _has_auto_numbering(text_frame) -> bool:
    return any(
        p._p.pPr is not None and p._p.pPr.find(qn("a:buAutoNum")) is not None and "".join(r.text for r in p.runs).strip()
        for p in text_frame.paragraphs
    )


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
    # Daftar bernomor otomatis dibawa apa adanya (penanda dan indentasi seperti di slide), tidak diparafrasekan model.
    taken = {line for block in blocks for line in block["lines"]}
    numbered = []
    for shape in _walk([sh for sh in slide.shapes if _on_canvas(sh, width, height)]):
        if _is_title(shape) or not shape.has_text_frame or not _has_auto_numbering(shape.text_frame):
            continue
        lines = [line for line in _paragraph_lines(shape.text_frame) if line not in taken]
        top = shape.top if shape.top is not None else 0
        if lines:
            numbered.append((top, {"label": "", "lines": lines}))
    blocks += [block for _top, block in sorted(numbered, key=lambda item: item[0])]
    return blocks


# Judul dan placeholder teks di bagian atas slide tidak ikut gambar: pita atas ini dipotong dari render.
TOP_BAND = 0.25  # placeholder dianggap "di atas" bila tepi atasnya di dalam pita ini
HEADING_MAX_CHARS = 120  # placeholder atas yang lebih panjang adalah kalimat isi
TOP_BAND_MAX_BOTTOM = 0.35  # dan tepi bawahnya tidak lebih dari ini (placeholder yang lebih besar adalah isi slide)


def top_content_edge(shapes, width: int, height: int) -> float:
    """Batas bawah (pecahan tinggi slide) judul dan placeholder teks di bagian atas slide; 0 bila tidak ada."""
    edge = 0.0
    for sh in shapes:
        if not sh.is_placeholder or _has_visual_content(sh) or not (sh.has_text_frame and sh.text_frame.text.strip()):
            continue
        box = _box(sh, width, height)
        if box is not None and box[1] < TOP_BAND and box[3] <= TOP_BAND_MAX_BOTTOM:
            edge = max(edge, box[3])
    return edge


def panel_facts(slide, panels: list[dict], width: int, height: int) -> None:
    """Pasang `facts` ({'text', 'tables', 'charts'}) pada tiap panel: isi objek yang letaknya di dalam kotak panel
    (judul panel sendiri tidak dihitung). Dipakai agar ringkasan panel mengikuti fakta panelnya, bukan seluruh slide."""
    for panel in panels:
        left, top, right, bottom = panel["bbox"]
        facts: dict = {"text": [], "tables": [], "charts": []}
        for sh in slide.shapes:
            box = _box(sh, width, height)
            if box is None:
                continue
            cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
            if not (left <= cx <= right and top <= cy <= bottom):
                continue
            if getattr(sh, "has_table", False) and sh.has_table:
                facts["tables"].append(_table_rows(sh.table))
            elif getattr(sh, "has_chart", False) and sh.has_chart:
                facts["charts"].append(_chart_data(sh.chart))
            else:
                for inner in _walk([sh]):
                    if inner.has_text_frame and inner.text_frame.text.strip() and not _is_title(inner):
                        facts["text"].extend(
                            line for line in _paragraph_lines(inner.text_frame)
                            if " ".join(line.split()) != panel.get("header")
                        )
                    elif getattr(inner, "has_table", False) and inner.has_table:
                        facts["tables"].append(_table_rows(inner.table))
                    elif getattr(inner, "has_chart", False) and inner.has_chart:
                        facts["charts"].append(_chart_data(inner.chart))
        panel["facts"] = facts


def top_heading(shapes, width: int, height: int):
    """Placeholder teks paling atas di pita atas slide (tanpa tipe judul) yang berfungsi sebagai judul, atau None."""
    best = None
    for sh in shapes:
        titled_box = not sh.is_placeholder and sh.name.lower().startswith("title")  # judul yang placeholder-nya terlepas
        if not (sh.is_placeholder or titled_box) or _has_visual_content(sh) or _is_title(sh):
            continue
        if sh.is_placeholder and sh.placeholder_format.type in NON_HEADING_TYPES:
            continue
        if not (sh.has_text_frame and sh.text_frame.text.strip()) or _has_auto_numbering(sh.text_frame):
            continue
        heading = " ".join(sh.text_frame.text.split())
        if len(heading) > HEADING_MAX_CHARS or heading.endswith((".", ":", ";")):
            continue  # kalimat isi, bukan judul
        box = _box(sh, width, height)
        if box is not None and box[1] < TOP_BAND and box[3] <= TOP_BAND_MAX_BOTTOM and (best is None or box[1] < best[0]):
            best = (box[1], sh)
    return best[1] if best else None


COLUMN_HEADER_TOP_MAX = 0.2  # lencana judul kolom berada di pita atas slide
COLUMN_HEADER_MAX_WIDTH = 0.3
COLUMN_HEADER_MAX_CHARS = 30
COLUMN_HEADER_MIN_UPPER = 0.8
COLUMN_SPAN_MAX_WIDTH = 0.6  # kotak teks selebar ini atau lebih melintang antar kolom: slide tidak dipecah
COLUMN_ICON_MAX_WIDTH = 0.1  # gambar kecil (ikon lencana) boleh ada; gambar besar berarti isi slide bukan sekadar teks


def detect_columns(slide, width: int, height: int) -> list[dict]:
    """Slide berkolom dengan lencana judul di atas tiap kolom (mis. MONETER | MAKROPRUDENSIAL | SP PUR), tiap kolom
    berisi kotak teks di bawah lencananya. Mengembalikan [{"header", "text"}] urut dari kiri; kosong bila slide
    bukan susunan itu (kurang dari dua lencana, lencana tumpang tindih, ada tabel/chart/gambar besar, ada kotak
    teks yang melintang antar kolom, atau ada kolom tanpa teks), supaya slide dibiarkan apa adanya."""
    shapes = list(_walk([sh for sh in slide.shapes if _on_canvas(sh, width, height)]))
    headers = []
    for sh in shapes:
        if not sh.has_text_frame or _is_title(sh):
            continue
        text = " ".join(sh.text_frame.text.split())
        box = _box(sh, width, height)
        letters = [c for c in text if c.isalpha()]
        if (box and letters and len(text) <= COLUMN_HEADER_MAX_CHARS and len(sh.text_frame.paragraphs) == 1
                and sum(c.isupper() for c in letters) / len(letters) >= COLUMN_HEADER_MIN_UPPER
                and box[1] < COLUMN_HEADER_TOP_MAX and box[2] - box[0] <= COLUMN_HEADER_MAX_WIDTH):
            headers.append({"header": text, "box": box, "shape": sh})
    if len(headers) < 2:
        return []
    headers.sort(key=lambda h: h["box"][0])
    if any(a["box"][2] > b["box"][0] for a, b in zip(headers, headers[1:])):
        return []
    centers = [(h["box"][0] + h["box"][2]) / 2 for h in headers]
    cuts = [(a + b) / 2 for a, b in zip(centers, centers[1:])]
    top_edge = min(h["box"][1] for h in headers) - 0.01
    header_shapes = {id(h["shape"]) for h in headers}
    columns: list[list] = [[] for _ in headers]
    for sh in shapes:
        box = _box(sh, width, height)
        if id(sh) in header_shapes or box is None:
            continue
        if getattr(sh, "has_table", False) and sh.has_table or getattr(sh, "has_chart", False) and sh.has_chart:
            return []
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            if box[2] - box[0] > COLUMN_ICON_MAX_WIDTH:
                return []
            continue
        if not sh.has_text_frame or not sh.text_frame.text.strip() or _is_title(sh) or box[1] < top_edge:
            continue
        if box[2] - box[0] >= COLUMN_SPAN_MAX_WIDTH:
            return []
        center = (box[0] + box[2]) / 2
        columns[sum(center > cut for cut in cuts)].append((box[1], box[0], sh))
    result = []
    for header, members in zip(headers, columns):
        lines = [line for _, _, sh in sorted(members, key=lambda m: (m[0], m[1])) for line in _paragraph_lines(sh.text_frame)]
        if not lines:
            return []
        result.append({"header": header["header"], "text": lines})
    return result


def extract_slide(slide, index: int, width: Optional[int] = None, height: Optional[int] = None) -> dict:
    """Ambil isi slide yang berada di dalam kanvas (width/height = ukuran slide, EMU). Speaker notes tidak diambil."""
    title = ""
    text: list[str] = []
    tables: list[list[list[str]]] = []
    charts: list[dict] = []
    n_pictures = 0
    has_other_visual = False  # SmartArt, OLE, dsb.

    top_level = [sh for sh in slide.shapes if _on_canvas(sh, width, height)]
    for sh in _walk(top_level):
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
    if not title and slide.shapes.title is not None and _on_canvas(slide.shapes.title, width, height):
        title = slide.shapes.title.text_frame.text.strip()
    # slide tanpa placeholder judul: placeholder teks paling atas di pita atas berfungsi sebagai judul
    if not title and width and height:
        heading = top_heading(top_level, width, height)
        if heading is not None:
            title = heading.text_frame.text.strip()
            for line in _paragraph_lines(heading.text_frame):
                if line in text:
                    text.remove(line)

    # Luas objek visual terbesar (gambar, chart, tabel, diagram) sebagai pecahan kanvas; menentukan lebar gambar di DOCX.
    visual_area = 0.0
    content_top = 0.0
    if width and height:
        content_top = top_content_edge(top_level, width, height)
        for sh in top_level:
            box = _box(sh, width, height) if _has_visual_content(sh) else None
            if box is not None:
                visual_area = max(visual_area, _area(box))

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
        "visual_area": round(visual_area, 3),
        "content_top": round(content_top, 3),
        "needs_vision": bool(reasons or charts),
        "vision_reasons": reasons + (["chart"] if charts else []),
    }


def find_soffice() -> Optional[str]:
    return shutil.which("soffice") or shutil.which("libreoffice")


def can_render_slides() -> bool:
    """True bila render slide ke PNG memungkinkan (LibreOffice + PyMuPDF)."""
    if not find_soffice():
        return False
    try:
        import pymupdf  # noqa: F401
    except ImportError:
        return False
    return True


def render_slides_png(pptx_path: Path, out_dir: Path, dpi: int = 110) -> dict[int, Path]:
    """PPTX -> PDF (LibreOffice) -> PNG per slide (PyMuPDF). Return {nomor_slide: path}."""
    soffice = find_soffice()
    if not can_render_slides():
        raise RuntimeError("Butuh LibreOffice (soffice) dan PyMuPDF (pip install pymupdf) untuk render PNG")
    import pymupdf

    out_dir.mkdir(parents=True, exist_ok=True)
    result: dict[int, Path] = {}
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
        with pymupdf.open(pdf) as doc:
            for index, page in enumerate(doc, start=1):
                png = out_dir / f"{pptx_path.stem}-{index}.png"
                page.get_pixmap(dpi=dpi).save(png)
                result[index] = png
    return result


P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def read_sections(prs) -> list[dict]:
    """Section PowerPoint ("Moneter", "Makroprudensial", dst.) dalam urutan deck: [{'nama', 'refs'}]. Kosong bila
    presentasi tidak memakai section."""
    from lxml import etree
    root = etree.fromstring(prs.part.blob)
    namespaces = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main", "p14": P14_NS}
    order = {int(node.get("id")): number
             for number, node in enumerate(root.xpath("p:sldIdLst/p:sldId", namespaces=namespaces), start=1)}
    sections = []
    for section in root.xpath(".//p14:sectionLst/p14:section", namespaces=namespaces):
        refs = [f"slide-{order[int(node.get('id'))]}" for node in section.xpath("p14:sldIdLst/p14:sldId", namespaces=namespaces)
                if int(node.get("id")) in order]
        sections.append({"nama": (section.get("name") or "").strip(), "refs": refs})
    return sections


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
    items = [extract_slide(s, i, prs.slide_width, prs.slide_height) for i, s in enumerate(prs.slides, start=1)]
    for item, slide in zip(items, prs.slides):
        try:  # kotak panel dan blok teks asli hanya pelengkap; kegagalan deteksi tidak boleh menggagalkan ingest
            item["panels"] = detect_panels(slide, prs.slide_width, prs.slide_height)
            panel_facts(slide, item["panels"], prs.slide_width, prs.slide_height)
        except Exception:
            item["panels"] = []
        try:
            item["verbatim"] = detect_verbatim(slide, prs.slide_width, prs.slide_height)
        except Exception:
            item["verbatim"] = []
        try:  # kolom berlencana (mis. MONETER | MAKROPRUDENSIAL | SP PUR); dipakai bila profil mengaturnya
            item["kolom"] = detect_columns(slide, prs.slide_width, prs.slide_height)
        except Exception:
            item["kolom"] = []

    if render_dir and any(it["needs_vision"] for it in items):
        pngs = render_slides_png(path, Path(render_dir))
        for i, it in enumerate(items, start=1):
            if it["needs_vision"] and i in pngs:
                it["image_path"] = str(pngs[i])
                if describe:
                    it["visual_desc"] = describe(pngs[i], it)

    return {
        "sections": read_sections(prs),
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
