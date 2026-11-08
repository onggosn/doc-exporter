"""Ukur properti halaman dari PDF referensi.

Yang diukur hanya hal yang memang ditentukan halaman: ukuran kertas, margin, nomor
halaman, watermark, dan teks footer yang berulang. Ukuran font di PDF hasil cetak
browser bisa terskala (tabel lebar disusutkan agar muat), jadi dicatat apa adanya
dan dibandingkan dengan ukuran yang dideklarasikan HTML di profile_builder.
"""
from __future__ import annotations

import math
import re
import statistics
from collections import Counter
from pathlib import Path
from typing import Any

import pdfplumber


PAPER_SIZES_PT = {"A4": (595.276, 841.89), "Letter": (612.0, 792.0), "Legal": (612.0, 1008.0)}
PAPER_TOLERANCE_PT = 2.0
WATERMARK_MIN_FONT_PT = 40.0
WATERMARK_MIN_CHARS = 3
FOOTER_ZONE_PT = 60.0
MIN_REPEAT_SHARE = 0.5
TOP_FONTS = 6
APPROVAL_STAMP = re.compile(r"^Approved by .+ at [\d/: ]+$")
_SUBSET_PREFIX = re.compile(r"^[A-Z]{6}\+")


def _font_name(raw: str) -> str:
    return _SUBSET_PREFIX.sub("", raw)


def _round(value: float) -> float:
    return round(float(value), 1)


def _paper_name(width: float, height: float) -> str | None:
    for name, (w, h) in PAPER_SIZES_PT.items():
        if abs(width - w) <= PAPER_TOLERANCE_PT and abs(height - h) <= PAPER_TOLERANCE_PT:
            return name
    return None


def _hex_color(color: Any) -> str | None:
    if not isinstance(color, (tuple, list)) or len(color) not in (1, 3):
        return None
    channels = [color[0]] * 3 if len(color) == 1 else color
    return "#" + "".join(f"{round(max(0.0, min(1.0, c)) * 255):02x}" for c in channels)


def _align(x0: float, x1: float, page_width: float) -> str:
    if page_width - x1 < page_width * 0.15:
        return "kanan"
    if abs((x0 + x1) / 2 - page_width / 2) < page_width * 0.1:
        return "tengah"
    return "kiri"


def _without_watermark(page):
    return page.filter(lambda obj: obj["object_type"] != "char" or obj["size"] < WATERMARK_MIN_FONT_PT)


def _watermark(pages) -> dict[str, Any] | None:
    texts: Counter[str] = Counter()
    sizes: list[float] = []
    fonts: Counter[str] = Counter()
    colors: Counter[str | None] = Counter()
    angles: list[float] = []
    with_watermark = 0
    for page in pages:
        big = [c for c in page.chars if c["size"] >= WATERMARK_MIN_FONT_PT]
        if len([c for c in big if c["text"].strip()]) < WATERMARK_MIN_CHARS:
            continue
        with_watermark += 1
        ordered = sorted(big, key=lambda c: c["x0"])
        texts["".join(c["text"] for c in ordered).replace(" ", "")] += 1
        sizes += [c["size"] for c in big]
        fonts.update(_font_name(c["fontname"]) for c in big)
        colors.update(_hex_color(c.get("non_stroking_color")) for c in big)
        a, b, *_ = big[0]["matrix"]
        angles.append(math.degrees(math.atan2(b, a)))
    if with_watermark / len(pages) < MIN_REPEAT_SHARE:
        return None
    return {
        "teks": texts.most_common(1)[0][0],
        "font": fonts.most_common(1)[0][0],
        # Ukuran huruf miring dihitung dari kotak pembatasnya, jadi hanya bisa diperkirakan sebagai rentang.
        "ukuran_pt_perkiraan": [_round(min(sizes)), _round(max(sizes))],
        "sudut_derajat": round(statistics.median(angles)),
        "warna": colors.most_common(1)[0][0],
        "halaman": with_watermark,
    }


def analyze_pdf(path: Path) -> dict[str, Any]:
    with pdfplumber.open(path) as pdf:
        pages = pdf.pages
        width, height = pages[0].width, pages[0].height
        footer_top = height - FOOTER_ZONE_PT

        lefts, right_gaps, tops, bottom_gaps = [], [], [], []
        typography: Counter[tuple[str, float]] = Counter()
        page_numbers: list[dict[str, Any]] = []
        footer_lines: dict[str, list[dict[str, Any]]] = {}
        for index, page in enumerate(pages, start=1):
            body = _without_watermark(page)
            words = body.extract_words(extra_attrs=["fontname", "size"])
            content = [w for w in words if w["top"] < footer_top]
            if content:
                lefts.append(min(w["x0"] for w in content))
                right_gaps.append(width - max(w["x1"] for w in content))
                tops.append(min(w["top"] for w in content))
                bottom_gaps.append(height - max(w["bottom"] for w in content))
            for word in content:
                typography[(_font_name(word["fontname"]), _round(word["size"]))] += len(word["text"])
            for word in (w for w in words if w["top"] >= footer_top):
                if re.fullmatch(r"\d+", word["text"]) and word["x0"] > width / 2:
                    page_numbers.append({**word, "page": index})
            for line in body.extract_text_lines():
                if line["top"] < footer_top or re.fullmatch(r"\d+", line["text"].strip()):
                    continue
                masked = re.sub(r"\d", "#", line["text"].strip())
                fonts = Counter(_font_name(c["fontname"]) for c in line["chars"])
                footer_lines.setdefault(masked, []).append({
                    "page": index, "x0": line["x0"], "top": line["top"],
                    "font": fonts.most_common(1)[0][0], "size": line["chars"][0]["size"],
                    "raw": line["text"].strip(),
                })

        total_chars = sum(typography.values()) or 1
        page_number = None
        if page_numbers:
            first = page_numbers[0]
            page_number = {
                "halaman_bernomor": len({p["page"] for p in page_numbers}),
                "dari_halaman": len(pages),
                "mulai": int(first["text"]),
                "posisi": {
                    "x0_pt": _round(statistics.median(p["x0"] for p in page_numbers)),
                    "top_pt": _round(statistics.median(p["top"] for p in page_numbers)),
                },
                "perataan": _align(first["x0"], first["x1"], width),
                "font": _font_name(first["fontname"]),
                "ukuran_pt": _round(first["size"]),
            }

        repeated = []
        for masked, occurrences in footer_lines.items():
            if len({o["page"] for o in occurrences}) / len(pages) < MIN_REPEAT_SHARE:
                continue
            is_stamp = bool(APPROVAL_STAMP.match(occurrences[0]["raw"]))
            repeated.append({
                # Nama penyetuju tidak disimpan di profil.
                "pola": "Approved by * at #/#/# #:#" if is_stamp else masked,
                "halaman": len({o["page"] for o in occurrences}),
                "posisi": {
                    "x0_pt": _round(statistics.median(o["x0"] for o in occurrences)),
                    "top_pt": _round(statistics.median(o["top"] for o in occurrences)),
                },
                "font": occurrences[0]["font"],
                "ukuran_pt": _round(occurrences[0]["size"]),
                "dikecualikan": is_stamp,
                "alasan": "dibubuhkan alur persetujuan, bukan bagian format dokumen" if is_stamp else None,
            })

        top_fonts = [
            {"font": font, "ukuran_pt": size, "bagian": round(count / total_chars, 3)}
            for (font, size), count in typography.most_common(TOP_FONTS)
        ]
        return {
            "file": path.name,
            "halaman": len(pages),
            "ukuran": {
                "lebar_pt": _round(width), "tinggi_pt": _round(height), "nama": _paper_name(width, height),
            },
            "margin_pt": {
                "kiri": _round(min(lefts)),
                "kanan": _round(min(right_gaps)),
                "atas": _round(min(tops)),
                "bawah": _round(min(bottom_gaps)),
                "indent_isi": _round(statistics.median(lefts)),
                "metode": "jarak terkecil konten ke tepi halaman; indent_isi = median tepi kiri teks per halaman",
            },
            "nomor_halaman": page_number,
            "footer_berulang": repeated,
            "watermark": _watermark(pages),
            "tipografi_pdf": {"isi_dominan": top_fonts[0] if top_fonts else None, "teratas": top_fonts},
        }
