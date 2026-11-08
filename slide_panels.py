"""Memecah slide visual menjadi panel (chart, tabel, atau gambar beserta judulnya) dari geometri PPTX.

Deterministik, tanpa model: judul panel adalah teks persis dari slide, dan kotak potongnya dari posisi
objek visual. Hasilnya hanya kotak ternormalisasi (0-1 terhadap ukuran slide); pemotongan gambar dilakukan
terhadap PNG hasil render slide oleh pemanggil.
"""
from __future__ import annotations

import re
from typing import Any

from pptx.enum.shapes import MSO_SHAPE_TYPE

MIN_OBJECT_AREA = 0.015  # objek lebih kecil (logo, ikon) diabaikan
MAX_OBJECT_AREA = 0.9  # objek sebesar slide adalah latar
HEADER_MIN_CHARS = 6
HEADER_MAX_CHARS = 200
HEADER_MAX_LINES = 3
HEADER_MAX_HEIGHT = 0.10
HEADER_MIN_WIDTH = 0.25
HEADER_MIN_UPPER = 0.6  # judul panel berhuruf kapital atau bernomor; label diagram biasanya tidak
HEADER_MIN_TOP = 0.06  # di bawah pita judul slide
HEADER_MAX_GAP = 0.07  # jarak vertikal judul panel ke objeknya
HEADER_MIN_OVERLAP = 0.5  # porsi lebar yang harus sejajar
NOTE_PREFIXES = ("ket", "sumber", "source", "note", "catatan", "*", "#")  # '#' hanya bila bukan judul bernomor
MERGE_OVERLAP = 0.4
MIN_PANEL_AREA = 0.04
MIN_PANEL_SIDE = 0.12
MIN_TOTAL_AREA = 0.25
UNHEADED_MIN_AREA = 0.05
PAD = 0.004
REGION_MIN_HEIGHT = 0.12
REGION_MIN_SHAPES = 3
REGION_BOTTOM_LIMIT = 0.97
BODY_TEXT_CHARS = 120  # teks sepanjang ini di dalam wilayah berarti isinya paragraf, bukan diagram

Box = tuple[float, float, float, float]  # kiri, atas, kanan, bawah


def _box(shape, width: int, height: int) -> Box | None:
    try:
        left, top, w, h = shape.left, shape.top, shape.width, shape.height
    except Exception:
        return None
    if None in (left, top, w, h) or not w or not h:
        return None
    box = (left / width, top / height, (left + w) / width, (top + h) / height)
    clamped = (max(0.0, box[0]), max(0.0, box[1]), min(1.0, box[2]), min(1.0, box[3]))
    return clamped if clamped[2] > clamped[0] and clamped[3] > clamped[1] else None


def _area(box: Box) -> float:
    return (box[2] - box[0]) * (box[3] - box[1])


def _union(a: Box, b: Box) -> Box:
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def _intersection(a: Box, b: Box) -> float:
    return max(0.0, min(a[2], b[2]) - max(a[0], b[0])) * max(0.0, min(a[3], b[3]) - max(a[1], b[1]))


def _has_visual_content(shape) -> bool:
    """Gambar, chart, tabel, objek OLE, atau diagram; group hanya dihitung bila ada isi semacam itu di dalamnya."""
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return any(_has_visual_content(child) for child in shape.shapes)
    return bool(
        shape.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT, MSO_SHAPE_TYPE.LINKED_OLE_OBJECT)
        or getattr(shape, "has_chart", False) and shape.has_chart
        or getattr(shape, "has_table", False) and shape.has_table
        or shape.element.tag.endswith("graphicFrame")
    )


def _header_text(shape, box: Box) -> str | None:
    if not shape.has_text_frame or (shape.is_placeholder and "TITLE" in str(shape.placeholder_format.type)):
        return None
    paragraphs = [p for p in (" ".join("".join(r.text for r in para.runs).split()) for para in shape.text_frame.paragraphs) if p]
    text = " ".join(paragraphs)
    if not (HEADER_MIN_CHARS <= len(text) <= HEADER_MAX_CHARS) or len(paragraphs) > HEADER_MAX_LINES:
        return None
    if box[1] < HEADER_MIN_TOP or box[3] - box[1] > HEADER_MAX_HEIGHT or box[2] - box[0] < HEADER_MIN_WIDTH:
        return None
    lowered = text.lower()
    numbered = re.match(r"#?\d+[.)]\s", text)
    if lowered.startswith(NOTE_PREFIXES) and not numbered:
        return None
    letters = [c for c in text if c.isalpha()]
    upper = sum(c.isupper() for c in letters) / len(letters) if letters else 0
    return text if numbered or upper >= HEADER_MIN_UPPER else None


def _horizontal_overlap(a: Box, b: Box) -> float:
    overlap = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    return overlap / max(min(a[2] - a[0], b[2] - b[0]), 1e-9)


def _header_region(header: Box, headers: list[tuple[Box, str]], notes: list[Box], loose: list[tuple[Box, int]]) -> Box | None:
    """Wilayah di bawah judul panel yang belum punya objek visual: kumpulan bentuk lepas (diagram). Berhenti di
    judul panel lain atau catatan kaki di kolom yang sama; ditolak bila isinya paragraf teks."""
    limit = REGION_BOTTOM_LIMIT
    for other, _ in headers:
        if other != header and other[1] > header[3] and _horizontal_overlap(header, other) >= HEADER_MIN_OVERLAP:
            limit = min(limit, other[1])
    for note in notes:
        if note[1] > header[3] and _horizontal_overlap(header, note) >= HEADER_MIN_OVERLAP:
            limit = min(limit, note[1])
    in_span = [
        (box, chars) for box, chars in loose
        if box != header and header[0] - 0.01 <= (box[0] + box[2]) / 2 <= header[2] + 0.01
    ]
    for box, chars in in_span:  # paragraf teks: di tepat bawah judul berarti isinya teks; lebih jauh, jadi batas bawah
        if chars > BODY_TEXT_CHARS and box[1] > header[3] - 0.01 and box[1] < limit:
            if box[1] - header[3] <= 0.06:
                return None
            limit = box[1]
    if limit - header[3] < REGION_MIN_HEIGHT:
        return None
    inside = [(box, chars) for box, chars in in_span if header[3] <= (box[1] + box[3]) / 2 <= limit]
    if len(inside) < REGION_MIN_SHAPES or any(chars > BODY_TEXT_CHARS for _, chars in inside):
        return None
    tight = inside[0][0]
    for box, _ in inside[1:]:
        tight = _union(tight, box)
    region = (max(header[0], tight[0]), header[3], min(header[2], max(tight[2], header[0])), min(limit, tight[3]))
    return region if region[2] > region[0] and region[3] - region[1] >= REGION_MIN_HEIGHT else None


def detect_panels(slide, width: int, height: int) -> list[dict[str, Any]]:
    """Daftar panel [{'bbox': (kiri, atas, kanan, bawah), 'header': teks | None}] dalam urutan baca, atau []
    bila slide tidak layak dipecah (kurang dari dua panel, atau panel terlalu kecil/sedikit)."""
    objects: list[Box] = []
    headers: list[tuple[Box, str]] = []
    loose: list[tuple[Box, int]] = []  # bentuk lepas (kotak diagram dsb.) dan panjang teksnya
    notes: list[Box] = []  # catatan kaki ("Ket:", "Sumber:") yang membatasi wilayah di atasnya
    for shape in slide.shapes:
        box = _box(shape, width, height)
        if box is None:
            continue
        if _has_visual_content(shape):
            if MIN_OBJECT_AREA <= _area(box) <= MAX_OBJECT_AREA:
                objects.append(box)
            continue
        text = _header_text(shape, box)
        if text:
            headers.append((box, text))
        if shape.has_text_frame and shape.text_frame.text.strip().lower().startswith(NOTE_PREFIXES[:5]):
            notes.append(box)
        if not (shape.is_placeholder or shape.shape_type == MSO_SHAPE_TYPE.LINE):
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                loose.extend([(box, 0)] * max(1, len(shape.shapes)))  # isi group dihitung sebagai bentuk lepas
            else:
                loose.append((box, len(shape.text_frame.text.strip()) if shape.has_text_frame else 0))

    panels: list[dict[str, Any]] = []
    for box in objects:
        best = None
        for header_box, text in headers:
            gap = box[1] - header_box[3]
            if -0.03 <= gap <= HEADER_MAX_GAP and _horizontal_overlap(header_box, box) >= HEADER_MIN_OVERLAP:
                if best is None or gap < best[0]:
                    best = (gap, header_box, text)
        header = best[2] if best else None
        for panel in panels:
            if header and panel["header"] == header:
                panel["bbox"] = _union(panel["bbox"], box)
                break
        else:
            if header or _area(box) >= UNHEADED_MIN_AREA:
                panels.append({"bbox": box, "header": header})

    adjacent = True
    while adjacent:  # potongan tabel/gambar yang bersambungan tanpa judul adalah satu panel
        adjacent = False
        for i, first in enumerate(panels):
            for second in panels[i + 1:]:
                if first["header"] or second["header"]:
                    continue
                a, b = first["bbox"], second["bbox"]
                gap = max(b[1] - a[3], a[1] - b[3])
                if gap <= 0.03 and _horizontal_overlap(a, b) >= 0.6:
                    first["bbox"] = _union(a, b)
                    panels.remove(second)
                    adjacent = True
                    break
            if adjacent:
                break

    panels = [
        p for p in panels
        if _area(p["bbox"]) >= MIN_PANEL_AREA
        and p["bbox"][2] - p["bbox"][0] >= MIN_PANEL_SIDE and p["bbox"][3] - p["bbox"][1] >= MIN_PANEL_SIDE
    ]
    used = {p["header"] for p in panels if p["header"]}
    for header_box, text in headers:
        if text in used:
            continue
        region = _header_region(header_box, headers, notes, loose)
        if region is not None and not any(_intersection(region, p["bbox"]) / _area(region) >= 0.3 for p in panels):
            panels.append({"bbox": region, "header": text})

    merged = True
    while merged:
        merged = False
        for i, first in enumerate(panels):
            for second in panels[i + 1:]:
                smaller = min(_area(first["bbox"]), _area(second["bbox"]))
                if smaller and _intersection(first["bbox"], second["bbox"]) / smaller >= MERGE_OVERLAP:
                    first["bbox"] = _union(first["bbox"], second["bbox"])
                    first["header"] = first["header"] or second["header"]
                    panels.remove(second)
                    merged = True
                    break
            if merged:
                break

    panels = [
        p for p in panels
        if _area(p["bbox"]) >= MIN_PANEL_AREA
        and p["bbox"][2] - p["bbox"][0] >= MIN_PANEL_SIDE and p["bbox"][3] - p["bbox"][1] >= MIN_PANEL_SIDE
    ]
    if len(panels) < 2 or sum(_area(p["bbox"]) for p in panels) < MIN_TOTAL_AREA:
        return []
    panels.sort(key=lambda p: (round(p["bbox"][1], 1), p["bbox"][0]))
    return [
        {"bbox": (max(0.0, l - PAD), max(0.0, t - PAD), min(1.0, r + PAD), min(1.0, b + PAD)), "header": p["header"]}
        for p in panels for l, t, r, b in [p["bbox"]]
    ]
