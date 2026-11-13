"""Rakit DOCX dari spesifikasi dokumen terstruktur dan properti halaman profil.

Pembuat ini tidak memutuskan isi: teks, pilihan layout, dan caption sudah final di
spesifikasi. Yang dilakukan di sini hanya format, secara deterministik. Font mengikuti
HTML referensi; watermark belum dipasang.

Spesifikasi:
    {"blocks": [
        {"type": "heading", "level": 1, "text": "..."},
        {"type": "paragraph", "text": "teks dengan <i>istilah asing</i> dan <b>tebal</b>"},
        {"type": "bullets", "items": ["...", "..."]},
        {"type": "numbered", "items": [{"marker": "1.", "text": "...", "depth": 0}]},   # penanda dari slide
        {"type": "image", "source": <path|bytes>}   # selebar halaman, di luar tabel; diikuti blok caption/paragraph
        {"type": "caption", "text": "...", "tipe": "headline" | "label", "align": "left"}
        {"type": "figure", "image": <path|bytes>, "caption": {"text": "...", "tipe": "label"},
         "width_percent": 80},
        {"type": "table", "widths": [40, 60], "banner": "Judul pita" | None,
         "rows": [[cell, cell], [cell]]}   # baris dengan satu sel direntangkan penuh
    ]}
    cell = [item, ...]; item = {"type": "image", "source": ...}
         | {"type": "caption", "text": "...", "tipe": "headline" | "label"}
         | {"type": "paragraph", "text": "..."} | {"type": "bullets", "items": [...]}
"""
from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor


BODY_FONT = "Frutiger 45 Light"
COMMENT_AUTHOR = "Peninjau Fakta"  # penulis komentar Word dari blok yang punya "comments"
COMMENT_INITIALS = "PF"
COMMENT_AUTHORS = (("[Cek format", "Peninjau Format", "PFm"),)  # awalan komentar -> (penulis, inisial) khusus
BODY_SIZE_PT = 11
CAPTION_SIZE_PT = 10
PAGE_NUMBER_SIZE_PT = 10
HEADING_SIZES_PT = {1: 12, 2: 11, 3: 11}
PARAGRAPH_SPACE_AFTER_PT = 6
HEADING_SPACE_BEFORE_PT = 12
CELL_PADDING_PT = 5.4
IMAGE_SAFETY_PT = 2
LIST_DEPTH_STEP_PT = 24  # geser per tingkat daftar bernomor
LIST_HANG_PT = 22  # lebar penanda (nomor) sebelum teks
TWIPS_PER_PT = 20
BORDER_EIGHTHS_PT = 8  # garis solid 1 pt (satuan sz adalah seperdelapan poin)
BANNER_FILL = "E6E6E6"
TABLE_HEADER_FILL = "E6E6E6"
FOOTER_DISTANCE_PT = 30
A4_PT = (595.28, 841.89)
DEFAULT_PAGE_SETUP = {
    "lebar_pt": A4_PT[0], "tinggi_pt": A4_PT[1],
    "margin": {"kiri": 34.0, "kanan": 34.0, "atas": 60.0, "bawah": 58.0},
    "nomor_halaman": True,
}



class DocxBuildError(ValueError):
    """Spesifikasi dokumen tidak valid; pesannya aman ditampilkan ke UI."""


def page_setup_from_profile(page: dict[str, Any] | None) -> dict[str, Any]:
    """Ukuran kertas dan margin dari properti halaman profil (rata-rata semua PDF referensi)."""
    pdfs = (page or {}).get("pdf") or []
    if not pdfs:
        return {**DEFAULT_PAGE_SETUP, "margin": dict(DEFAULT_PAGE_SETUP["margin"])}

    def mean(values: list[float]) -> float:
        return round(sum(values) / len(values), 1)

    first = pdfs[0]["ukuran"]
    a4 = all(pdf["ukuran"]["nama"] == "A4" for pdf in pdfs)
    return {
        "lebar_pt": A4_PT[0] if a4 else first["lebar_pt"],
        "tinggi_pt": A4_PT[1] if a4 else first["tinggi_pt"],
        "margin": {
            "kiri": mean([m for pdf in pdfs for m in (pdf["margin_pt"]["kiri"], pdf["margin_pt"]["kanan"])]),
            "kanan": mean([m for pdf in pdfs for m in (pdf["margin_pt"]["kiri"], pdf["margin_pt"]["kanan"])]),
            "atas": mean([pdf["margin_pt"]["atas"] for pdf in pdfs]),
            "bawah": mean([pdf["margin_pt"]["bawah"] for pdf in pdfs]),
        },
        "nomor_halaman": any(pdf["nomor_halaman"] for pdf in pdfs),
    }


def _force_font(font, name: str) -> None:
    """Pasang font dan buang rujukan tema yang akan menimpanya di Word."""
    font.name = name
    fonts = font._element.rPr.rFonts
    for attribute in [a for a in fonts.attrib if a.endswith("Theme")]:
        del fonts.attrib[attribute]
    fonts.set(qn("w:cs"), name)
    fonts.set(qn("w:eastAsia"), name)


def _configure_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    _force_font(normal.font, BODY_FONT)
    normal.font.size = Pt(BODY_SIZE_PT)
    normal.paragraph_format.space_after = Pt(0)
    for level, size in HEADING_SIZES_PT.items():
        style = document.styles[f"Heading {level}"]
        _force_font(style.font, BODY_FONT)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.italic = level == 3
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(HEADING_SPACE_BEFORE_PT)
        style.paragraph_format.space_after = Pt(PARAGRAPH_SPACE_AFTER_PT)
        style.paragraph_format.keep_with_next = True
    _force_font(document.styles["List Bullet"].font, BODY_FONT)


def _configure_section(document: Document, setup: dict[str, Any]) -> None:
    section = document.sections[0]
    section.page_width, section.page_height = Pt(setup["lebar_pt"]), Pt(setup["tinggi_pt"])
    margin = setup["margin"]
    section.left_margin, section.right_margin = Pt(margin["kiri"]), Pt(margin["kanan"])
    section.top_margin, section.bottom_margin = Pt(margin["atas"]), Pt(margin["bawah"])
    section.footer_distance = Pt(FOOTER_DISTANCE_PT)
    if not setup["nomor_halaman"]:
        return
    paragraph = section.footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    field = parse_xml(
        f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"><w:r><w:rPr>'
        f'<w:sz w:val="{PAGE_NUMBER_SIZE_PT * 2}"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>'
    )
    paragraph._p.append(field)


_CONTROL = re.compile(r"[\x00-\x08\x0b-\x1f\x7f]")
_TAG = re.compile(r"</?(?:i|b)>|</?[a-zA-Z][^>]*>")


def _clean(text: str) -> str:
    """Buang karakter kontrol yang tidak boleh ada di XML Word. Tab vertikal (pindah baris lunak
    PowerPoint) dan form feed menjadi spasi."""
    return _CONTROL.sub(lambda match: " " if match.group() in "\x0b\x0c" else "", text)


def _add_runs(paragraph, text: str, *, size: float | None = None, bold: bool = False, italic: bool = False) -> None:
    """Tulis teks dengan penanda sebaris <i> dan <b>, boleh bersarang. Penutup tanpa pembuka diabaikan dan
    pembuka tanpa penutup berlaku sampai akhir teks; tag lain ditolak agar tidak tampil mentah."""

    def run(piece: str, piece_bold: bool, piece_italic: bool) -> None:
        if not piece:
            return
        item = paragraph.add_run(piece)
        item.bold = piece_bold or None
        item.italic = piece_italic or None
        if size:
            item.font.size = Pt(size)

    text = _clean(text)
    depth = {"b": 0, "i": 0}
    position = 0
    for match in _TAG.finditer(text):
        tag = match.group().strip("</>").lower()
        if tag not in depth:
            raise DocxBuildError(f"Tag tidak didukung di dalam teks (hanya <i> dan <b>): {match.group()[:60]!r}")
        run(text[position:match.start()], bold or depth["b"] > 0, italic or depth["i"] > 0)
        depth[tag] = max(0, depth[tag] + (-1 if match.group().startswith("</") else 1))
        position = match.end()
    run(text[position:], bold or depth["b"] > 0, italic or depth["i"] > 0)


def _image_stream(source: Any):
    if isinstance(source, (bytes, bytearray)):
        return io.BytesIO(source)
    path = Path(str(source))
    if not path.is_file():
        raise DocxBuildError(f"Berkas gambar tidak ditemukan: {path.name}")
    return str(path)


def _add_picture(paragraph, source: Any, width_pt: float) -> None:
    try:
        paragraph.add_run().add_picture(_image_stream(source), width=Pt(width_pt))
    except DocxBuildError:
        raise
    except Exception as exc:
        raise DocxBuildError(f"Gambar tidak dapat dibaca: {exc}") from exc


def _caption(paragraph, item: dict[str, Any]) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if item.get("align") == "left" else WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(PARAGRAPH_SPACE_AFTER_PT)
    headline = item.get("tipe") == "headline"
    _add_runs(paragraph, item["text"], size=BODY_SIZE_PT if headline else CAPTION_SIZE_PT, bold=headline)


def _shade(cell, fill: str) -> None:
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def _fill_container(
    next_paragraph, content: list[dict[str, Any]], width_pt: float,
    notes: list[tuple[Any, list[str]]] | None = None,
) -> None:
    """Isi sel atau badan dokumen; `next_paragraph()` memberi paragraf kosong berikutnya. Paragraf yang punya
    "comments" dicatat di `notes` untuk dijadikan komentar Word setelah dokumen selesai disusun."""
    for position, item in enumerate(content):
        kind = item.get("type")
        if kind == "image":
            paragraph = next_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_after = Pt(PARAGRAPH_SPACE_AFTER_PT)
            # Caption di bawah gambar harus ikut halaman yang sama.
            followed_by_caption = position + 1 < len(content) and content[position + 1].get("type") == "caption"
            paragraph.paragraph_format.keep_with_next = followed_by_caption
            _add_picture(paragraph, item["source"], width_pt)
        elif kind == "caption":
            paragraph = next_paragraph()
            _caption(paragraph, item)
            # Caption judul berada di atas gambar dan harus ikut halaman yang sama dengannya.
            paragraph.paragraph_format.keep_with_next = item.get("tipe") == "headline"
        elif kind == "paragraph":
            paragraph = next_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.space_after = Pt(PARAGRAPH_SPACE_AFTER_PT)
            _add_runs(paragraph, item["text"])
            if notes is not None and item.get("comments"):
                notes.append((paragraph, [str(text) for text in item["comments"]]))
        elif kind == "numbered":
            for entry in item["items"]:
                paragraph = next_paragraph()
                depth = min(max(int(entry.get("depth", 0)), 0), 2)
                left = LIST_DEPTH_STEP_PT * depth + LIST_HANG_PT
                paragraph.paragraph_format.left_indent = Pt(left)
                paragraph.paragraph_format.first_line_indent = Pt(-LIST_HANG_PT)
                paragraph.paragraph_format.tab_stops.add_tab_stop(Pt(left))
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                paragraph.paragraph_format.space_after = Pt(PARAGRAPH_SPACE_AFTER_PT)
                _add_runs(paragraph, f"{entry['marker']}\t{entry['text']}")
                if notes is not None and entry.get("comments"):
                    notes.append((paragraph, [str(text) for text in entry["comments"]]))
        elif kind == "bullets":
            item_notes = item.get("item_comments") or {}
            for number, text in enumerate(item["items"]):
                paragraph = next_paragraph()
                level = min((len(text) - len(text.lstrip(" "))) // 2, 2)  # butir menjorok = bersarang
                paragraph.style = "List Bullet" if level == 0 else f"List Bullet {level + 1}"
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                _add_runs(paragraph, text.strip())
                if notes is not None and item_notes.get(number):
                    notes.append((paragraph, [str(note) for note in item_notes[number]]))
        else:
            raise DocxBuildError(f"Tipe isi sel tidak dikenal: {kind!r}")


def _cell_paragraphs(cell):
    state = {"first": True}

    def next_paragraph():
        if state["first"]:
            state["first"] = False
            return cell.paragraphs[0]
        return cell.add_paragraph()

    return next_paragraph


def _add_table(document: Document, item: dict[str, Any], content_width_pt: float, notes: list | None = None) -> None:
    widths = item.get("widths")
    rows = item.get("rows")
    if not (isinstance(widths, list) and widths and all(isinstance(w, (int, float)) and w > 0 for w in widths)):
        raise DocxBuildError("Tabel memerlukan 'widths' berupa daftar persentase positif.")
    if not isinstance(rows, list) or not rows:
        raise DocxBuildError("Tabel memerlukan minimal satu baris.")
    total = sum(widths)
    column_pt = [content_width_pt * w / total for w in widths]
    banner = item.get("banner")

    table = document.add_table(rows=len(rows) + (1 if banner else 0), cols=len(widths))
    table.autofit = False  # menulis tblLayout fixed
    properties = table._tbl.tblPr
    # Lebar tabel, tblGrid, dan indentasi harus eksplisit; hanya mengatur lebar sel dibaca berbeda oleh
    # Word dan LibreOffice (tanpa tblInd isi sel bergeser dan gambar tampak terpotong).
    table_width = properties.find(qn("w:tblW"))
    table_width.set(qn("w:type"), "dxa")
    table_width.set(qn("w:w"), str(round(content_width_pt * TWIPS_PER_PT)))
    indent = OxmlElement("w:tblInd")
    indent.set(qn("w:w"), "0")
    indent.set(qn("w:type"), "dxa")
    properties.insert_element_before(indent, "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")
    sides = "".join(
        f'<w:{side} w:val="single" w:sz="{BORDER_EIGHTHS_PT}" w:space="0" w:color="000000"/>'
        for side in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    borders = parse_xml(f'<w:tblBorders {nsdecls("w")}>{sides}</w:tblBorders>')
    properties.insert_element_before(borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")
    for column, width in zip(table.columns, column_pt):
        column.width = Pt(width)
    for row in table.rows:
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for cell, width in zip(row.cells, column_pt):
            cell.width = Pt(width)

    offset = 0
    if banner:
        merged = table.rows[0].cells[0].merge(table.rows[0].cells[-1])
        _shade(merged, BANNER_FILL)
        _add_runs(merged.paragraphs[0], banner, bold=True)
        merged.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        offset = 1

    for index, row in enumerate(rows):
        cells = table.rows[index + offset].cells
        if len(row) == 1 and len(widths) > 1:
            target, width_pt = cells[0].merge(cells[-1]), content_width_pt
            targets = [(target, width_pt, row[0])]
        elif len(row) == len(widths):
            targets = [(cell, width, content) for cell, width, content in zip(cells, column_pt, row)]
        else:
            raise DocxBuildError(f"Baris tabel #{index + 1} memiliki {len(row)} sel; diharapkan 1 atau {len(widths)}.")
        for cell, width_pt, content in targets:
            _fill_container(_cell_paragraphs(cell), content, width_pt - 2 * CELL_PADDING_PT - IMAGE_SAFETY_PT, notes)
    document.add_paragraph().paragraph_format.space_after = Pt(PARAGRAPH_SPACE_AFTER_PT)


def _add_figure(document: Document, item: dict[str, Any], content_width_pt: float) -> None:
    percent = item.get("width_percent", 100)
    if not (isinstance(percent, (int, float)) and 0 < percent <= 100):
        raise DocxBuildError("'width_percent' harus di antara 0 dan 100.")
    content: list[dict[str, Any]] = [{"type": "image", "source": item["image"]}]
    caption = item.get("caption")
    if caption and caption.get("text"):
        entry = {"type": "caption", **caption}
        content = [entry, *content] if caption.get("tipe") == "headline" else [*content, entry]
    _fill_container(document.add_paragraph, content, content_width_pt * percent / 100)


def build_docx(spec: dict[str, Any], page_setup: dict[str, Any], target: Path | io.BytesIO) -> None:
    blocks = spec.get("blocks") if isinstance(spec, dict) else None
    if not isinstance(blocks, list):
        raise DocxBuildError("Spesifikasi dokumen harus berisi daftar 'blocks'.")
    document = Document()
    _configure_styles(document)
    _configure_section(document, page_setup)
    margin = page_setup["margin"]
    content_width = page_setup["lebar_pt"] - margin["kiri"] - margin["kanan"]
    notes: list[tuple[Any, list[str]]] = []

    for index, block in enumerate(blocks, start=1):
        kind = block.get("type") if isinstance(block, dict) else None
        if kind == "heading":
            level = block.get("level")
            if level not in HEADING_SIZES_PT:
                raise DocxBuildError(f"Blok #{index}: level judul harus 1-{len(HEADING_SIZES_PT)}.")
            _add_runs(document.add_paragraph(style=f"Heading {level}"), block["text"])
        elif kind == "paragraph":
            _fill_container(document.add_paragraph, [block], content_width, notes)
        elif kind in ("bullets", "numbered"):
            _fill_container(document.add_paragraph, [block], content_width, notes)
        elif kind in ("image", "caption"):  # di luar tabel; gambar selebar halaman, caption di bawahnya
            _fill_container(document.add_paragraph, [block], content_width)
            following = blocks[index] if index < len(blocks) else None
            if kind == "image" and isinstance(following, dict) and following.get("type") == "caption":
                document.paragraphs[-1].paragraph_format.keep_with_next = True  # caption tidak terpisah dari gambar
        elif kind == "figure":
            _add_figure(document, block, content_width)
        elif kind == "table":
            _add_table(document, block, content_width, notes)
        else:
            raise DocxBuildError(f"Blok #{index}: tipe tidak dikenal: {kind!r}.")
    for paragraph, comments in notes:  # komentar Word per catatan, menempel pada seluruh teks paragraf
        if not paragraph.runs:
            continue
        for text in comments:
            author, initials = next(((a, i) for prefix, a, i in COMMENT_AUTHORS if text.startswith(prefix)),
                                    (COMMENT_AUTHOR, COMMENT_INITIALS))
            document.add_comment(paragraph.runs, text=_clean(text), author=author, initials=initials)
    document.save(target)
