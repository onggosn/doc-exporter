"""Susun spesifikasi dokumen (untuk docx_builder) dari slide, pemetaan bab, dan narasi.

Semua keputusan format ada di sini dan deterministik: slide mana yang jadi gambar,
layout tabel, jenis dan nomor caption. Model bahasa hanya menulis narasi per slide,
dan angka di narasi diperiksa terhadap isi slide.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
import re
from collections import Counter
from typing import Any

from PIL import Image, ImageChops

import list_review


LABEL_KINDS = ("Grafik", "Tabel", "Gambar", "Foto", "Skema", "Diagram")
# Gambar yang luasnya lebih dari separuh kanvas slide dipasang selebar 100% halaman (satu sel penuh), bukan di
# kolom kiri tabel dua kolom. Slide utuh diukur dari objek visual terbesarnya (visual_area), panel dari bbox-nya.
FULL_WIDTH_MIN_AREA = 0.5
IMAGE_REASONS = {"picture", "diagram_or_object", "chart"}
DEFAULT_WIDTHS = [50, 50]
SIDE_BY_SIDE_TOKENS = "I,T"
TRIM_PADDING_PX = 12
MAX_NARRATION_ITEMS = 6
MAX_NARRATION_ITEMS_DENSE = 60
FRAGMENT_CHARS = 40
FRAGMENT_GROUP_CHARS = 200
MAX_NARRATION_CHARS = 700
SKIPPABLE_CHARS = 200
SUMMARY_MAX_CHARS = 400
DEFAULT_CAPTION_FACTS_MAX_CHARS = 3000  # batas teks slide di prompt ringkasan gambar; env CAPTION_FACTS_MAX_CHARS
VERBATIM_CHAPTER_KEYWORDS = ("kesimpulan", "arahan")  # bab yang teksnya dibawa apa adanya, tanpa narasi model
COVER_MAX_CHARS = 400
LAMPIRAN_DIVIDER_MAX_CHARS = 300
AGENDA_MAX_CHARS = 1500  # slide agenda/outline hanya daftar topik; yang lebih panjang dianggap isi
AGENDA_HEADING = re.compile(r"(agenda|outline|daftar isi|garis besar)\b", re.I)
LABEL_MAX_CHARS = 60
LABEL_MIN_UPPER = 0.8
TEXT_DOMINANT_CHARS = 1500
FILLER_CHARS = 300
FILLER_LINES = 10
FILLER_LINE_CHARS = 60

_NUMBER = re.compile(r"\d+(?:[.,]\d+)*")
_THOUSANDS = re.compile(r"\d{1,3}(?:[.,]\d{3})+")


def trim_margins(path: str, padding: int = TRIM_PADDING_PX) -> None:
    """Potong margin polos di sekeliling gambar slide yang dirender, sisakan sedikit ruang."""
    with Image.open(path) as image:
        rgb = image.convert("RGB")
        box = ImageChops.difference(rgb, Image.new("RGB", rgb.size, rgb.getpixel((0, 0)))).getbbox()
        if not box:
            return
        left, top, right, bottom = box
        crop = (max(0, left - padding), max(0, top - padding),
                min(rgb.width, right + padding), min(rgb.height, bottom + padding))
        if crop == (0, 0, rgb.width, rgb.height):
            return
        rgb.crop(crop).save(path)


def text_dominant(slide: dict[str, Any]) -> bool:
    """Isi slide pada dasarnya teks, jadi ditulis sebagai teks tanpa gambar: tanpa tabel, chart, atau diagram,
    hanya elemen gambar (biasanya dekorasi), dan teksnya panjang. Slide bertabel selalu dijadikan gambar."""
    if slide.get("tables") or slide.get("charts") or set(slide.get("vision_reasons") or []) & {"chart", "diagram_or_object"}:
        return False
    return len("\n".join(slide.get("text", []))) >= TEXT_DOMINANT_CHARS


MIN_CROP_PX = 40


def crop_top(png_path: str, edge: float) -> None:
    """Buang pita atas gambar slide (judul dan placeholder atas) sampai pecahan tinggi `edge`."""
    if not 0 < edge < 1:
        return
    with Image.open(png_path) as image:
        rgb = image.convert("RGB")
        cut = round(edge * rgb.height)
        if cut <= 0 or rgb.height - cut < MIN_CROP_PX:
            return
        cropped = rgb.crop((0, cut, rgb.width, rgb.height))
    cropped.save(png_path)


def crop_panels(
    png_path: str, panels: list[dict[str, Any]], out_dir: Path, ref: str, content_top: float = 0.0
) -> list[dict[str, Any]]:
    """Potong PNG slide (belum dipangkas marginnya) menjadi gambar per panel. Bagian di atas `content_top`
    (judul dan placeholder atas) tidak ikut. Panel yang hasil potongannya hampir kosong dibuang.
    Mengembalikan [{'header', 'bbox', 'image_path'}] sesuai urutan panel."""
    result = []
    with Image.open(png_path) as image:
        rgb = image.convert("RGB")
        for number, panel in enumerate(panels, start=1):
            left, top, right, bottom = panel["bbox"]
            top = max(top, content_top)
            box = (round(left * rgb.width), round(top * rgb.height), round(right * rgb.width), round(bottom * rgb.height))
            if box[2] - box[0] < MIN_CROP_PX or box[3] - box[1] < MIN_CROP_PX:
                continue
            path = Path(out_dir) / f"{ref}-panel{number}.png"
            rgb.crop(box).save(path)
            trim_margins(str(path))
            with Image.open(path) as saved:
                if saved.width < MIN_CROP_PX or saved.height < MIN_CROP_PX:
                    path.unlink(missing_ok=True)
                    continue
            result.append({"header": panel.get("header"), "bbox": panel["bbox"], "image_path": str(path),
                           "facts": panel.get("facts")})
    return result


def slide_needs_image(slide: dict[str, Any]) -> bool:
    """Slide dengan gambar, chart, diagram, atau tabel dijadikan gambar (render slide), kecuali yang isinya teks."""
    visual = bool(set(slide.get("vision_reasons") or []) & IMAGE_REASONS) or bool(slide.get("tables"))
    return visual and not text_dominant(slide)


def narration_form(slide: dict[str, Any]) -> str:
    """Bentuk narasi, semuanya paragraf: slide bergambar 'gambar' (ringkas), slide teks panjang 'lengkap' (satu
    paragraf per gagasan, isi tidak dihilangkan), slide teks lainnya 'paragraf' (padat)."""
    if slide.get("image_path"):
        return "gambar"
    return "lengkap" if text_dominant(slide) else "paragraf"


def is_verbatim_chapter(chapter: str) -> bool:
    """Bab kesimpulan/arahan: teks slide dibawa apa adanya ke dokumen."""
    lowered = chapter.lower()
    return any(keyword in lowered for keyword in VERBATIM_CHAPTER_KEYWORDS)


def excluded_slides(items: list[dict[str, Any]]) -> dict[str, str]:
    """Slide yang dikeluarkan dari seluruh proses: sampul (slide pertama bila hanya berisi teks pendek), lampiran
    (dari slide pembatas "LAMPIRAN ..." pertama sampai akhir dek), dan agenda/outline pembahasan (judul atau baris
    pertamanya diawali "Agenda", "Outline", "Daftar isi", atau "Garis besar", tanpa tabel dan chart, dan tidak
    panjang). Mengembalikan {ref: alasan}."""
    excluded: dict[str, str] = {}
    appendix = False
    for index, slide in enumerate(items):
        lines = [str(line).strip() for line in [slide.get("title") or "", *slide.get("text", [])] if str(line).strip()]
        short = (not slide.get("tables") and not slide.get("charts")
                 and sum(len(line) for line in lines) <= LAMPIRAN_DIVIDER_MAX_CHARS)
        if not appendix and short and any(re.match(r"lampiran\b", line, re.I) for line in lines):
            appendix = True
        if appendix:
            excluded[slide["ref"]] = "lampiran"
        elif index == 0 and not slide.get("tables") and not slide.get("charts") \
                and sum(len(line) for line in lines) <= COVER_MAX_CHARS:
            excluded[slide["ref"]] = "sampul"
        elif lines and AGENDA_HEADING.match(lines[0]) and len(lines[0]) <= LABEL_MAX_CHARS \
                and not slide.get("tables") and not slide.get("charts") \
                and sum(len(line) for line in lines) <= AGENDA_MAX_CHARS:
            excluded[slide["ref"]] = "agenda"
    return excluded


def _words(text: str) -> list[str]:
    return re.findall(r"\w+", text.lower())


def match_column(header: str, columns: list[dict[str, str]]) -> int | None:
    """Indeks kolom pengaturan profil yang cocok dengan teks lencana ("SP PUR" cocok dengan "SP"): sama persis
    atau seluruh kata `cocok` ada di lencana. None bila tidak ada yang cocok atau lebih dari satu."""
    header_words = _words(header)
    found = [
        index for index, column in enumerate(columns)
        if _words(column["cocok"]) == header_words
        or (_words(column["cocok"]) and all(word in header_words for word in _words(column["cocok"])))
    ]
    exact = [i for i in found if _words(columns[i]["cocok"]) == header_words]
    return exact[0] if len(exact) == 1 else found[0] if len(found) == 1 else None


def split_default_section(deck: dict[str, Any], config: dict[str, Any] | None) -> list[str]:
    """Pecah slide berkolom di section bawaan PowerPoint ("Default Section") menjadi satu slide semu per kolom,
    berjudul sesuai pengaturan profil (mis. "1.1 Kesimpulan dan Arahan Moneter"), berurutan seperti daftar kolom
    profil. Teks kolom dibawa apa adanya. Slide asal diganti di deck["items"] dan di daftar slide section.
    Tanpa pengaturan, tanpa section yang cocok, atau bila ada lencana yang tak cocok dengan pengaturan, slide
    dibiarkan apa adanya. Mengembalikan ref slide asal yang dipecah (dan memberi deck["section_bawaan_dipakai"])."""
    if not config:
        return []
    wanted = " ".join(config["bagian_cocok"].lower().split())
    sections = [s for s in deck.get("sections") or [] if " ".join(s["nama"].lower().split()) == wanted]
    if not sections:
        return []
    in_section = {ref for section in sections for ref in section["refs"]}
    split: list[str] = []
    for item in list(deck["items"]):
        columns = item.get("kolom") or []
        if item["ref"] not in in_section or len(columns) < 2:
            continue
        indexes = [match_column(column["header"], config["kolom"]) for column in columns]
        if None in indexes or len(set(indexes)) != len(indexes):
            continue
        virtual = []
        for index, column in sorted(zip(indexes, columns), key=lambda pair: pair[0]):
            title = config["judul_kolom"].replace("{n}", str(index + 1)).replace("{nama}", config["kolom"][index]["nama"])
            virtual.append({
                "ref": f"{item['ref']}-k{index + 1}", "title": title, "text": list(column["text"]),
                "verbatim": [{"label": "", "lines": list(column["text"])}], "tables": [], "charts": [],
                "visual_desc": None, "visual_area": 0.0, "content_top": 0.0, "needs_vision": False,
                "vision_reasons": [], "panels": [], "kolom_asal": item["ref"], "kolom_judul": title,
            })
        position = deck["items"].index(item)
        deck["items"][position:position + 1] = virtual
        for section in deck.get("sections") or []:
            if item["ref"] in section["refs"]:
                at = section["refs"].index(item["ref"])
                section["refs"][at:at + 1] = [v["ref"] for v in virtual]
        split.append(item["ref"])
    if split:
        deck["section_bawaan_dipakai"] = True
    return split


def pin_split_slides(chapters: list[str], mapping: dict[str, list[str]], items: list[dict[str, Any]]) -> None:
    """Mode daftar bab: slide semu hasil pemecahan kolom dipasang ke bab yang nomornya sama dengan judulnya
    ("1.1 ..." ke bab "1.1 ..."), tanpa menunggu pemetaan model. Tanpa bab bernomor sama, pemetaan model dibiarkan."""
    def number(text: str) -> str | None:
        match = re.match(r"\s*(\d+(?:\.\d+)*)\.?(?:\s|$)", text)
        return match.group(1) if match else None

    for item in items:
        title = item.get("kolom_judul")
        target = next((c for c in chapters if title and number(c) and number(c) == number(title)), None)
        if target is None:
            continue
        for refs in mapping.values():
            if item["ref"] in refs:
                refs.remove(item["ref"])
        mapping.setdefault(target, []).append(item["ref"])


def split_verbatim(slide: dict[str, Any]) -> None:
    """Pisahkan baris blok teks asli ("Hasil Pembahasan Dewan Gubernur", "Key Takeaways") dari teks slide agar
    tidak ikut diringkas model. Teks utuh asalnya disimpan di `text_original`. Aman dipanggil ulang."""
    blocks = slide.get("verbatim") or []
    if not blocks or "text_original" in slide:
        return
    slide["text_original"] = list(slide.get("text", []))
    remaining = list(slide["text_original"])
    for block in blocks:
        for line in [block["label"], *block["lines"]]:
            if line in remaining:
                remaining.remove(line)
    slide["text"] = remaining


def all_facts(slide: dict[str, Any]) -> str:
    """Fakta slide untuk pemeriksaan angka: teks yang dinarasikan ditambah blok teks asli."""
    extra = [line for block in slide.get("verbatim") or [] for line in block["lines"]]
    return "\n".join(part for part in (fact_text(slide), "\n".join(extra)) if part)


def _is_label(line: str) -> bool:
    text = line.strip()
    letters = [c for c in text if c.isalpha()]
    return bool(letters) and len(text) <= LABEL_MAX_CHARS and not text.endswith((".", ";", ":")) \
        and sum(c.isupper() for c in letters) / len(letters) >= LABEL_MIN_UPPER


LIST_MARKER = re.compile(r"^(\(?(?:\d{1,3}|[a-zA-Z]|[ivxlcdmIVXLCDM]{1,6})[.)])\s+(\S.*)$")


def verbatim_blocks(lines: list[str]) -> list[dict[str, Any]]:
    """Ubah baris teks asli menjadi blok dokumen tanpa mengubah isinya: baris bernomor ("1.", "a)") menjadi butir
    bernomor dengan penandanya sendiri, baris berhuruf kapital pendek menjadi label tebal, baris lain menjadi
    butir (yang menjorok tetap bersarang). Indentasi dua spasi per tingkat mengikuti slide."""
    blocks: list[dict[str, Any]] = []
    for line in lines:
        indent = len(line) - len(line.lstrip(" "))
        text = re.sub(r"[<>]", "", " ".join(line.split()))
        if not text:
            continue
        numbered = LIST_MARKER.match(text)
        if numbered:
            item = {"marker": numbered.group(1), "text": numbered.group(2), "depth": min(indent // 2, 2)}
            if blocks and blocks[-1]["type"] == "numbered":
                blocks[-1]["items"].append(item)
            else:
                blocks.append({"type": "numbered", "items": [item]})
        elif indent == 0 and _is_label(text):
            blocks.append({"type": "paragraph", "text": f"<b>{text}</b>"})
        else:
            item = " " * indent + text
            if blocks and blocks[-1]["type"] == "bullets":
                blocks[-1]["items"].append(item)
            else:
                blocks.append({"type": "bullets", "items": [item]})
    return blocks


def fact_text(slide: dict[str, Any]) -> str:
    """Seluruh teks faktual sebuah slide; satu-satunya sumber yang boleh dipakai narasi."""
    parts = [slide.get("title") or "", *slide.get("text", [])]
    for table in slide.get("tables") or []:
        parts += [" ".join(map(str, row)) for row in table]
    for chart in slide.get("charts") or []:
        parts += [chart.get("title", ""), " ".join(chart.get("categories", []))]
        for series in chart.get("series", []):
            parts += [str(series.get("name", "")), " ".join(str(v) for v in series.get("values", []) if v is not None)]
    parts += [slide.get("visual_desc") or ""]
    return "\n".join(p for p in parts if p)


class NarrationError(ValueError):
    """Narasi dari model tidak memenuhi bentuk yang diminta."""


def has_substance(slide: dict[str, Any]) -> bool:
    """Slide punya isi selain judulnya, jadi ada yang bisa dinarasikan."""
    return bool(fact_text({**slide, "title": ""}).strip())


def visual_only(slide: dict[str, Any]) -> bool:
    """Slide yang tidak punya fakta tertulis; satu-satunya isinya deskripsi visual dari model."""
    return not fact_text({**slide, "title": "", "visual_desc": None}).strip()


def may_skip(slide: dict[str, Any]) -> bool:
    """Slide yang boleh tanpa narasi: hanya gambar tanpa fakta tertulis, atau slide pembatas (teks pendek tanpa
    angka, tabel, chart, maupun gambar). Model boleh menjawab [] dan slide dicatat sebagai dilewati."""
    if visual_only(slide):
        return True
    facts = fact_text({**slide, "title": "", "visual_desc": None})
    return (not slide.get("image_path") and not slide.get("tables") and not slide.get("charts")
            and len(facts) <= SKIPPABLE_CHARS and not re.search(r"\d", facts))


def fallback_narration(slides: list[dict[str, Any]]) -> dict[str, list[str]]:
    """Narasi cadangan tanpa model: teks tertulis slide apa adanya (tanpa judul, baris kembar dibuang).
    Baris-baris pendek yang berurutan (angka, label) digabung satu butir agar tidak jadi bullet tunggal.
    Slide tanpa teks tertulis (hanya gambar) tidak mendapat narasi."""
    result: dict[str, list[str]] = {}
    for slide in slides:
        seen: set[str] = set()
        items: list[str] = []
        short: list[str] = []

        def flush_short() -> None:
            if short:
                items.append("; ".join(short))
                short.clear()

        for raw in slide.get("text", []):
            line = re.sub(r"[<>]", "", " ".join(str(raw).split()))
            if not line or line in seen:
                continue
            seen.add(line)
            if len(line) < FRAGMENT_CHARS:
                short.append(line)
                if len("; ".join(short)) >= FRAGMENT_GROUP_CHARS:
                    flush_short()
            else:
                flush_short()
                items.append(line)
        flush_short()
        result[slide["ref"]] = items
    return result


def split_long(text: str, limit: int = MAX_NARRATION_CHARS) -> list[str]:
    """Pecah butir yang terlalu panjang pada batas kalimat (atau spasi) menjadi butir-butir dalam batas."""
    if len(text) <= limit:
        return [text]
    pieces: list[str] = []
    current = ""
    for sentence in re.split(r"(?<=[.;!?])\s+", text):
        while len(sentence) > limit:  # satu kalimat lebih panjang dari batas: potong di spasi terakhir
            cut = sentence.rfind(" ", 0, limit) or limit
            if current:
                pieces.append(current)
                current = ""
            pieces.append(sentence[:cut].strip())
            sentence = sentence[cut:].strip()
        if current and len(current) + 1 + len(sentence) > limit:
            pieces.append(current)
            current = sentence
        else:
            current = f"{current} {sentence}".strip()
    if current:
        pieces.append(current)
    return [piece for piece in pieces if piece]


_MARK_NUMBER = re.compile(r"\s*\(\s*\d{1,3}\s*\)\s*$")
_MARK_CONTINUED = re.compile(r"\s*(?:…|\.{3})?\s*\(\s*lanjutan\s*(?:…|\.{3})?\s*\)", re.I)
_HEADING_LINE = re.compile(r"^\s*\d+(?:\.\d+)*\.?\s+(\S.*)$")


def strip_continuation_marker(text: str) -> str:
    """Hapus penanda sambungan "(1)", "(2)", ... di ujung dan "(lanjutan)" (beserta "…" sebelumnya)."""
    cleaned = _MARK_CONTINUED.sub("", text or "")
    cleaned = _MARK_NUMBER.sub("", cleaned)
    return re.sub(r"\s*…\s*$", "", cleaned).rstrip()


def _alnum(text: str) -> str:
    return re.sub(r"[^0-9a-z]", "", text.lower())


def merge_continuations(items: list[dict[str, Any]]) -> None:
    """Slide berurutan dengan judul sama yang bertanda sambungan ("(1)", "(2)", "lanjutan") dijadikan satu kesatuan:
    penanda dihapus dari judul, slide sambungan ditandai `continues` (judulnya tidak diulang di dokumen), dan baris
    sub-judul bernomor yang hanya mengulang ("2. Kebijakan Operasi Moneter (lanjutan)") dibuang dari teks slide."""
    previous_key, previous_marked, seen = "", False, set()
    for item in items:
        original = (item.get("title") or "").strip()
        cleaned = strip_continuation_marker(original)
        marked = cleaned != original
        key = _alnum(cleaned)
        continues = bool(key) and key == previous_key and (marked or previous_marked)
        if original:
            item["title"] = cleaned
        item["continues"] = continues
        if not continues:
            seen = set()
        previous_key, previous_marked = key, marked

        earlier = set(seen)

        def fix(line: str) -> str | None:
            base = strip_continuation_marker(line)
            heading = _HEADING_LINE.match(base.strip())
            if base != line.rstrip() and heading and _alnum(heading.group(1)) in earlier:
                return None  # sub-judul yang sama sudah muncul di slide sebelumnya
            return base

        def fix_all(lines: list[str]) -> list[str]:
            return [fixed for fixed in map(fix, lines) if fixed is not None and fixed.strip()]

        for name in ("text", "text_original"):
            if item.get(name):
                item[name] = fix_all(item[name])
        for block in item.get("verbatim") or []:
            block["lines"] = fix_all(block["lines"])
        for line in [*(item.get("text_original") or item.get("text") or []),
                     *(line for block in item.get("verbatim") or [] for line in block["lines"])]:
            heading = _HEADING_LINE.match(line.strip())
            if heading:
                seen.add(_alnum(heading.group(1)))


def heading_level(chapter: str) -> int:
    """Tingkat judul dari nomor bab: '1.' → 1, '2.1 ...' → 2, '4.7.2' → 3. Tanpa nomor → 1."""
    match = re.match(r"\s*(\d+(?:\.\d+)*)\.?(?:\s|$)", chapter)
    return min(match.group(1).count(".") + 1, 3) if match else 1


def is_filler(slide: dict[str, Any]) -> bool:
    """Sampul, agenda, pembatas bagian, atau penutup: tanpa judul, teks pendek berupa label (tanpa kalimat,
    persen, atau rupiah), tanpa tabel, chart, atau diagram. Dokumen referensi tidak memuat slide semacam ini."""
    if (slide.get("title") or "").strip() or slide.get("tables") or slide.get("charts") or slide.get("verbatim") \
            or set(slide.get("vision_reasons") or []) & {"chart", "diagram_or_object"}:
        return False
    facts = fact_text({**slide, "title": "", "visual_desc": None})
    lines = [line for line in facts.splitlines() if line.strip()]
    if not lines or len(facts) > FILLER_CHARS or len(lines) > FILLER_LINES:
        return False
    return all(
        len(line) <= FILLER_LINE_CHARS and not line.rstrip().endswith((".", ":", ";")) and not re.search(r"%|Rp", line)
        for line in lines
    )


def caption_facts_limit() -> int:
    try:
        return int(os.environ.get("CAPTION_FACTS_MAX_CHARS", DEFAULT_CAPTION_FACTS_MAX_CHARS))
    except ValueError:
        return DEFAULT_CAPTION_FACTS_MAX_CHARS


def trim_facts(facts: str, limit: int | None = None) -> str:
    """Potong teks fakta untuk prompt ringkasan gambar pada batas baris, supaya teks slide yang panjang
    tidak terkirim penuh di setiap panel. Pemeriksaan angka tetap memakai fakta lengkap. Batas 0 = tanpa potong."""
    limit = caption_facts_limit() if limit is None else limit
    if limit <= 0 or len(facts) <= limit:
        return facts
    kept, size = [], 0
    for line in facts.splitlines():
        if size + len(line) + 1 > limit:
            if not kept:  # satu baris melebihi batas: potong barisnya
                kept.append(line[:limit])
            break
        kept.append(line)
        size += len(line) + 1
    return "\n".join(kept) + "\n(teks slide dipotong; sisanya tidak ditampilkan)"


PANEL_FACTS_MAX_CHARS = 600
FALLBACK_LABELS = 3
FALLBACK_LABEL_CHARS = 45  # label baris yang lebih panjang (kalimat) tidak disebut
FALLBACK_COLUMNS = 6


def panel_fact_text(facts: dict[str, Any] | None, limit: int = PANEL_FACTS_MAX_CHARS) -> str:
    """Isi satu panel (teks, tabel, chart) sebagai teks untuk prompt ringkasan; dipotong pada batas baris."""
    if not facts:
        return ""
    lines = list(facts.get("text") or [])
    lines += [" | ".join(str(cell) for cell in row) for table in facts.get("tables") or [] for row in table]
    for chart in facts.get("charts") or []:
        lines.append("; ".join(part for part in (
            chart.get("title", ""), " ".join(chart.get("categories", [])),
            *(f"{item.get('name', '')}: {' '.join(str(v) for v in item.get('values', []) if v is not None)}"
              for item in chart.get("series", [])),
        ) if part))
    return trim_facts("\n".join(line for line in lines if line.strip()), limit)


def fallback_summary(facts: dict[str, Any] | None) -> str:
    """Ringkasan deterministik dari fakta tabel/chart panel (tanpa model) agar kolom caption tidak kosong.
    Hanya memuat apa yang tertulis di slide; kosong bila panel tidak punya tabel atau chart."""
    if not facts:
        return ""
    sentences = []
    tables, charts = facts.get("tables") or [], facts.get("charts") or []
    numeric = re.compile(r"[-+()]*\d[\d.,%\s()+-]*")
    cleaned = [[row for row in ([" ".join(str(cell).split()) for cell in raw] for raw in table) if any(row)] for table in tables]
    cleaned = [rows for rows in cleaned if rows]
    for number, rows in enumerate(cleaned, start=1):
        first_data = next((i for i, row in enumerate(rows) if any(numeric.fullmatch(cell) for cell in row[1:])), len(rows))
        header_rows, data_rows = rows[:first_data], rows[first_data:]
        title = ""
        if len(header_rows) > 1 and not any(header_rows[0][1:]):
            title, header_rows = header_rows[0][0], header_rows[1:]
        columns = list(dict.fromkeys(cell for row in header_rows for cell in row[1:] if cell))[:FALLBACK_COLUMNS]
        labels = [row[0] for row in data_rows if row and row[0] and len(row[0]) <= FALLBACK_LABEL_CHARS]
        name = "Tabel" if len(cleaned) == 1 else f"Tabel ke-{number}"
        sentence = f"{name} {title}".strip() + f" memuat {len(data_rows) or len(rows)} baris data"
        if labels:
            sentence += f", antara lain {', '.join(labels[:FALLBACK_LABELS])}"
        if columns:
            sentence += f" dengan kolom {', '.join(columns)}"
        sentences.append(sentence + ".")
    for chart in charts:
        names = [str(item.get("name", "")).strip() for item in chart.get("series", []) if str(item.get("name", "")).strip()]
        categories = [str(c) for c in chart.get("categories", [])]
        sentence = "Grafik" + (f" {chart['title']}" if chart.get("title") else "")
        if names:
            sentence += f" memuat seri {', '.join(names[:FALLBACK_COLUMNS])}"
        if len(categories) >= 2:
            sentence += f" untuk periode {categories[0]} hingga {categories[-1]}"
        if sentence != "Grafik":
            sentences.append(sentence + ".")
    return clean_summary(" ".join(sentences[:2]))


def image_summary_prompt(title: str | None, header: str | None, facts: str, feedback: str = "",
                         panel_facts: str = "") -> str:
    """Prompt ringkasan gambar untuk caption: hanya yang terlihat pada gambar, tanpa fakta tambahan."""
    return (
        "Tulis ringkasan 1-2 kalimat (maksimal 350 karakter) tentang apa yang ditampilkan gambar ini, untuk "
        "caption dokumen resmi berbahasa Indonesia.\n"
        "ATURAN KETAT:\n"
        "- Hanya apa yang terlihat jelas pada gambar. Jangan menambah fakta, tafsiran, sebab-akibat, "
        "rekomendasi, atau sumber. Yang tidak terbaca jangan disebut.\n"
        "- Sebut angka hanya bila tertulis jelas pada gambar DAN ada di teks slide di bawah; salin persis.\n"
        "- Jangan memulai dengan 'Gambar ini menunjukkan' atau sejenisnya; langsung isinya.\n"
        "- Teks polos tanpa markdown, tanpa tag, tanpa simbol poin.\n\n"
        f"Judul slide: {title or '(tanpa judul)'}\n"
        f"Judul panel gambar: {header or '(tanpa judul)'}\n"
        + (f"Isi panel ini menurut file slide (acuan fakta; pastikan sesuai gambar):\n{panel_facts}\n" if panel_facts else "")
        + f"Teks slide (hanya untuk mencocokkan angka dan istilah):\n{facts or '(tidak ada)'}"
        f"{feedback}"
    )


def panels_summary_prompt(title: str | None, headers: list[str | None], facts: str,
                          panel_facts: list[str] | None = None) -> str:
    """Prompt satu panggilan untuk beberapa panel dari slide yang sama: teks slide dikirim sekali saja."""
    keys = json.dumps({str(number): "…" for number in range(1, len(headers) + 1)}, ensure_ascii=False)
    panel_facts = panel_facts or [""] * len(headers)
    listing = "\n".join(
        f"Panel {number}: {header or '(tanpa judul)'}" + (f"\n  Isi panel menurut file slide: {detail}" if detail else "")
        for number, (header, detail) in enumerate(zip(headers, panel_facts), start=1)
    )
    return (
        f"Ada {len(headers)} gambar panel dari satu slide. Untuk SETIAP panel, tulis ringkasan 1-2 kalimat "
        "(maksimal 350 karakter) tentang apa yang ditampilkan gambarnya, untuk caption dokumen resmi berbahasa "
        "Indonesia. Setiap gambar didahului label 'Panel N: judul'.\n"
        "ATURAN KETAT:\n"
        "- Hanya apa yang terlihat jelas pada gambar panel itu sendiri. Jangan menambah fakta, tafsiran, "
        "sebab-akibat, rekomendasi, atau sumber. Yang tidak terbaca jangan disebut. Jangan mencampur isi antar panel.\n"
        "- Sebut angka hanya bila tertulis jelas pada gambar panel itu DAN ada di teks slide di bawah; salin persis.\n"
        "- Jangan memulai dengan 'Gambar ini menunjukkan' atau sejenisnya; langsung isinya.\n"
        "- Teks polos tanpa markdown, tanpa tag, tanpa simbol poin.\n"
        f"- Kembalikan HANYA satu objek JSON tanpa markdown dengan key nomor panel. Bentuk: {keys}\n\n"
        f"Judul slide: {title or '(tanpa judul)'}\n"
        f"Daftar panel:\n{listing}\n"
        f"Teks slide (hanya untuk mencocokkan angka dan istilah):\n{facts or '(tidak ada)'}"
    )


def clean_summary(text: str) -> str:
    """Bersihkan ringkasan dari markdown dan tag, lalu potong di batas kalimat bila terlalu panjang."""
    text = re.sub(r"```.*?```|[*_`#]+|</?[a-zA-Z][^>]*>", "", text or "", flags=re.S)
    text = re.sub(r"^\s*[-•]\s+", "", " ".join(text.split()))
    if len(text) <= SUMMARY_MAX_CHARS:
        return text
    cut = text[:SUMMARY_MAX_CHARS]
    end = max(cut.rfind(". "), cut.rfind("; "))
    return cut[:end + 1] if end > SUMMARY_MAX_CHARS // 2 else cut[:cut.rfind(" ")].rstrip(",;") + "."


def narration_limit(slide: dict[str, Any]) -> int:
    """Batas butir narasi sebuah slide: 6 untuk slide biasa, naik sesuai jumlah baris fakta pada slide padat."""
    lines = sum(1 for line in fact_text({**slide, "title": ""}).splitlines() if line.strip())
    return max(MAX_NARRATION_ITEMS, min(lines, MAX_NARRATION_ITEMS_DENSE))


def narration_prompt(chapter: str, slides: list[dict[str, Any]], style_text: str) -> str:
    listing = "\n\n".join(
        f"[{slide['ref']}] bentuk: {narration_form(slide)}, maksimal {narration_limit(slide)} butir\nJudul: {slide.get('title') or '(tanpa judul)'}\n"
        f"Fakta:\n{fact_text({**slide, 'title': ''})}"
        for slide in slides
    )
    example = json.dumps({slide["ref"]: ["…"] for slide in slides[:2]}, ensure_ascii=False)
    return (
        f"Tulis narasi untuk dokumen resmi berbahasa Indonesia, bab \"{chapter}\", dari fakta slide di bawah.\n\n"
        "ATURAN KETAT:\n"
        "- Hanya pakai fakta, termasuk angka, yang tertulis pada bagian Fakta slide yang bersangkutan. Jangan "
        "menambah fakta, angka, kutipan, atau sumber; jangan memakai fakta dari slide lain atau dari contoh gaya.\n"
        "- Salin angka persis seperti di slide, jangan membulatkan dan jangan menghitung angka baru. "
        "Pemisah desimal boleh disesuaikan dengan aturan gaya di bawah.\n"
        "- Setiap butir jawaban adalah SATU PARAGRAF berupa kalimat utuh; bukan poin, tanpa nomor, tanpa simbol "
        "poin. Bentuk 'gambar': 1-3 paragraf ringkas. Bentuk 'paragraf': 1-2 paragraf padat. Bentuk 'lengkap': satu "
        "paragraf untuk tiap gagasan penting slide, urut seperti di slide, tanpa menghilangkan isi; jumlah paragraf "
        "mengikuti batas tiap slide.\n"
        "- Slide bergambar dinarasikan maknanya dari data dan deskripsi visual; jangan menulis kata-kata seperti "
        "'pada gambar di atas'.\n"
        "- Jangan menambah tafsiran, sebab-akibat, atau implikasi yang tidak tertulis di slide (mis. 'menunjukkan "
        "penguatan akses kredit'). Boleh membandingkan atau menyebut urutan angka yang ada pada slide itu.\n"
        "- Slide yang hanya berupa gambar tanpa fakta tertulis yang jelas, atau slide pembatas yang hanya memuat judul "
        "bagian: isi dengan daftar kosong [].\n"
        f"- Batas butir per slide tertulis pada tiap slide; tiap butir maksimal {MAX_NARRATION_CHARS} karakter. "
        "Slide yang padat boleh dipecah menjadi lebih banyak butir, tetapi jangan melewati batasnya.\n"
        "- Teks polos. Satu-satunya tag yang boleh adalah <i> dan <b>; tanpa markdown, tanpa simbol poin.\n"
        "- Kembalikan HANYA satu objek JSON tanpa markdown; key persis seperti ref slide, nilai berupa daftar "
        f"teks. Bentuk: {example}\n\n"
        f"Aturan gaya penulisan:\n{style_text or '(tidak ada; gunakan bahasa Indonesia formal dan jelas)'}\n\n"
        f"Slide:\n{listing}"
    )


def parse_narration(
    data: Any, refs: list[str], optional: set[str] | frozenset[str] = frozenset(),
    limits: dict[str, int] | None = None,
) -> dict[str, list[str]]:
    """Validasi jawaban model. Slide di `optional` (hanya gambar, tanpa fakta tertulis) boleh kosong atau tidak ada.
    `limits` menaikkan batas butir per ref (bawaan MAX_NARRATION_ITEMS)."""
    if not isinstance(data, dict):
        raise NarrationError("jawaban bukan objek JSON")
    result: dict[str, list[str]] = {}
    for ref in refs:
        items = data.get(ref)
        if ref in optional and (items is None or items == []):
            result[ref] = []
            continue
        if not isinstance(items, list) or not items or not all(isinstance(t, str) for t in items):
            raise NarrationError(f"slide {ref} tidak berisi daftar teks")
        cleaned = [piece for t in items if t.strip() for piece in split_long(re.sub(r"^[-•*]\s+", "", t.strip()))]
        limit = (limits or {}).get(ref, MAX_NARRATION_ITEMS)
        if not cleaned or len(cleaned) > limit:
            raise NarrationError(f"slide {ref} harus berisi 1-{limit} butir")
        for text in cleaned:
            if len(text) > MAX_NARRATION_CHARS:
                raise NarrationError(f"slide {ref} memiliki butir lebih dari {MAX_NARRATION_CHARS} karakter")
            if "**" in text or "```" in text or re.search(r"</?[a-zA-Z]", re.sub(r"</?[ib]>", "", text)):
                raise NarrationError(f"slide {ref} memuat markdown atau tag selain <i> dan <b>")
        result[ref] = cleaned
    return result


REVIEW_MAX_FINDINGS = 6


def review_prompt(title: str | None, paragraphs: list[str], facts: str, with_image: bool) -> str:
    """Prompt tinjauan fakta ringkasan gambar satu slide: model hanya menandai, tidak menulis ulang."""
    listing = "\n".join(f"[{i}] {text}" for i, text in enumerate(paragraphs, start=1))
    source = "fakta tertulis slide" + (" dan gambar yang menyertai tiap paragraf" if with_image else "")
    return (
        "Anda peninjau fakta untuk dokumen yang disusun dari slide presentasi. Tiap paragraf di bawah ini adalah "
        f"ringkasan sebuah gambar (grafik, tabel, atau diagram) dari slide. Periksa apakah paragraf sesuai dengan {source}.\n\n"
        "Tandai HANYA klaim yang salah: angka, nama, periode, atau arah perubahan yang bertentangan dengan sumber, "
        "atau fakta penting yang tidak ada di sumber (misalnya sebab, prediksi, atau angka karangan). "
        "Angka dan teks yang jelas terbaca pada gambar adalah sumber yang sah. "
        "Jangan menilai gaya bahasa, jangan menandai ringkasan yang wajar, dan jangan menulis ulang paragraf.\n"
        "Jawab hanya dengan JSON: "
        '{"temuan": [{"paragraf": <nomor>, "kutipan": "<potongan kalimat bermasalah, maks 120 karakter>", '
        '"alasan": "<satu kalimat: apa yang salah dan apa kata sumber>"}]}. '
        f'Bila semua benar, jawab {{"temuan": []}}. Maksimal {REVIEW_MAX_FINDINGS} temuan.\n\n'
        f"Judul slide: {title or '(tanpa judul)'}\n\nFakta tertulis slide:\n{facts}\n\nParagraf yang ditinjau:\n{listing}"
    )


def parse_review(data: Any, count: int) -> list[dict[str, Any]]:
    """Temuan valid dari jawaban model: nomor paragraf di 1..count dan alasan tidak kosong; sisanya dibuang."""
    items = data.get("temuan") if isinstance(data, dict) else None
    result = []
    for item in items if isinstance(items, list) else []:
        if not isinstance(item, dict):
            continue
        number, reason = item.get("paragraf"), item.get("alasan")
        if isinstance(number, str) and number.strip().isdigit():
            number = int(number)
        if not (isinstance(number, int) and not isinstance(number, bool) and 1 <= number <= count):
            continue
        if not (isinstance(reason, str) and reason.strip()):
            continue
        quote = item.get("kutipan")
        result.append({
            "paragraf": number,
            "kutipan": " ".join(quote.split())[:120] if isinstance(quote, str) else "",
            "alasan": " ".join(reason.split())[:400],
        })
    return result[:REVIEW_MAX_FINDINGS]


def normalize_number(token: str) -> str:
    """Samakan penulisan angka: pemisah ribuan dibuang, koma/titik desimal disatukan, nol di belakang dibuang."""
    if _THOUSANDS.fullmatch(token):
        return re.sub(r"[.,]", "", token)
    token = token.replace(",", ".")
    if token.count(".") > 1:
        return token.replace(".", "")
    return token.rstrip("0").rstrip(".") if "." in token else token


def extract_numbers(text: str) -> set[str]:
    return {normalize_number(match) for match in _NUMBER.findall(text)}


def unsupported_numbers(narration: list[str], facts: str) -> list[str]:
    """Angka di narasi yang tidak ada di teks fakta slide, terurut."""
    return sorted(extract_numbers("\n".join(narration)) - extract_numbers(facts))


def format_rules(statistics: dict[str, Any] | None) -> list[str]:
    """Aturan format yang diturunkan langsung dari statistik terukur referensi."""
    if not statistics:
        return []
    numbers = statistics["angka"]
    rules = []
    if numbers["desimal_koma"] or numbers["desimal_titik"]:
        comma = numbers["desimal_koma"] >= numbers["desimal_titik"]
        rules.append("Pisahkan desimal dengan koma (4,85)." if comma else "Pisahkan desimal dengan titik (4.85).")
    if numbers["persen_rapat"] or numbers["persen_spasi"]:
        attached = numbers["persen_rapat"] >= numbers["persen_spasi"]
        rules.append("Tulis persen rapat dengan angka (4,5%)." if attached else "Tulis persen dengan spasi (4,5 %).")
    if numbers["rupiah_rapat"] or numbers["rupiah_spasi"]:
        attached = numbers["rupiah_rapat"] >= numbers["rupiah_spasi"]
        rules.append("Tulis rupiah rapat dengan angka (Rp16.550)." if attached else "Tulis rupiah dengan spasi (Rp 16.550).")
    italics = statistics.get("miring", {})
    if italics.get("jumlah"):
        rules.append("Istilah asing atau teknis berbahasa Inggris ditulis dengan <i>…</i> (miring).")
    return rules


def style_context(style: dict[str, Any] | None) -> str:
    """Teks gaya untuk prompt narasi, dari statistik dan (bila ada) panduan model."""
    if not style:
        return ""
    lines = format_rules(style.get("statistik"))
    saved = style.get("panduan")
    guide = saved["panduan"] if saved else None
    if guide:
        lines += [
            f"Nada: {guide['nada']}", f"Sudut pandang: {guide['sudut_pandang']}",
            f"Panjang kalimat: {guide['panjang_kalimat']}", f"Gaya poin: {guide['gaya_poin']}",
            f"Istilah dan singkatan: {guide['istilah_dan_singkatan']}", f"Format angka: {guide['format_angka']}",
            f"Istilah asing: {guide['istilah_asing']}",
        ]
        lines += [f"Hindari: {item}" for item in guide["larangan"]]
        if guide["contoh"]:
            lines.append("Contoh kutipan gaya (tiru gayanya saja, jangan pakai isinya):")
            lines += [f'- [{item["konteks"]}] "{item["teks"]}"' for item in guide["contoh"]]
    return "\n".join(lines)


def pick_widths(row_patterns: list[dict[str, Any]]) -> list[int]:
    """Lebar kolom paling umum untuk baris gambar-kiri-narasi-kanan di referensi."""
    for pattern in row_patterns:
        if pattern["tokens"] == SIDE_BY_SIDE_TOKENS:
            usable = {k: v for k, v in pattern["width_buckets"].items() if re.fullmatch(r"\d+(/\d+)+", k)}
            if usable:
                best = max(usable, key=usable.get)
                return [int(part) for part in best.split("/")]
    return list(DEFAULT_WIDTHS)


class CaptionNumbering:
    """Menentukan caption tiap gambar dan memberi nomor sesuai aturan penomoran profil."""

    def __init__(self, settings: dict[str, Any]):
        self.rules = [r for r in settings.get("penomoran", []) if r["aktif"]]
        self.default = settings.get("caption_default", "kosong")
        self.last: dict[Any, int] = {}

    @staticmethod
    def _title(slide: dict[str, Any]) -> str:
        return (slide.get("title") or "").strip()

    @staticmethod
    def _content_kind(slide: dict[str, Any]) -> str:
        return "Grafik" if slide.get("charts") else "Tabel" if slide.get("tables") else "Gambar"

    def rules_for(self, chapter: str) -> list[str]:
        """Awalan aturan penomoran yang cocok dengan judul bab (kosong bila tidak ada)."""
        return [r["awalan"] for r in self.rules if r["bagian_cocok"].lower() in chapter.lower()]

    def _number(self, index: int, rule: dict[str, Any], chapter: str) -> int:
        key = index if rule["lanjut"] == "berlanjut" else (index, chapter)
        number = self.last.get(key, rule["mulai"] - 1) + 1
        self.last[key] = number
        return number

    def caption_for(self, slide: dict[str, Any], chapter: str) -> dict[str, str] | None:
        title = self._title(slide)
        matched = [
            (i, r) for i, r in enumerate(self.rules) if r["bagian_cocok"].lower() in chapter.lower()
        ]
        labels = [(i, r) for i, r in matched if r["awalan"].capitalize() in LABEL_KINDS]
        if labels:
            wanted = [self._content_kind(slide), "Grafik", "Tabel", "Gambar"]
            index, rule = next(
                ((i, r) for kind in wanted for i, r in labels if r["awalan"].capitalize() == kind), labels[0]
            )
            prefix = f"{rule['awalan'].capitalize()} {self._number(index, rule, chapter)}"
            return {"text": f"{prefix}. {title}" if title else prefix, "tipe": "label"}
        if matched:
            index, rule = matched[0]
            prefix = f"{rule['awalan']}{self._number(index, rule, chapter)}"
            return {"text": f"{prefix}. {title}" if title else prefix, "tipe": "headline"}
        if self.default == "headline" and title:
            return {"text": title, "tipe": "headline"}
        return None


def compose(
    deck: dict[str, Any],
    chapters: list[str],
    mapping: dict[str, list[str]],
    narration: dict[str, list[str]],
    settings: dict[str, Any],
    row_patterns: list[dict[str, Any]],
    review: dict[tuple[str, str], list[str]] | None = None,
    review_lists: bool = False,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Kembalikan (spesifikasi untuk docx_builder, laporan susunan). `review` (ref, teks ringkasan gambar) -> catatan
    tinjauan fakta, dipasang sebagai komentar Word pada ringkasan yang sama. `review_lists` memeriksa format, urutan
    butir, dan penomoran pada blok daftar yang dibawa apa adanya (list_review) dan mencatat temuannya di laporan."""

    def paragraph(ref: str, text: str) -> dict[str, Any]:
        notes = (review or {}).get((ref, text))
        return {"type": "paragraph", "text": text, **({"comments": list(notes)} if notes else {})}

    order = {slide["ref"]: i for i, slide in enumerate(deck["items"])}
    by_ref = {slide["ref"]: slide for slide in deck["items"]}
    widths = pick_widths(row_patterns)
    numbering = CaptionNumbering(settings)
    blocks: list[dict[str, Any]] = []
    report: dict[str, Any] = {
        "slide": len(deck["items"]), "gambar": 0, "narasi": 0, "slide_dilewati": [], "slide_teks": [], "slide_dipecah": [], "slide_teks_asli": [], "gambar_diringkas": 0, "bab_kosong": [],
        "aturan_bab": {}, "tinjauan_format": [],
    }

    for chapter in chapters:
        refs = sorted(mapping.get(chapter, []), key=order.__getitem__)
        if not refs:
            report["bab_kosong"].append(chapter)
            continue
        blocks.append({"type": "heading", "level": heading_level(chapter), "text": chapter})
        report["aturan_bab"][chapter] = numbering.rules_for(chapter)

        verbatim_chapter = is_verbatim_chapter(chapter)
        last_title = ""  # judul slide yang terakhir tampil di bab ini; slide sambungan tidak mengulangnya
        for ref in refs:
            slide = by_ref[ref]
            if is_filler(slide):
                report["slide_dilewati"].append(ref)
                continue
            if verbatim_chapter:  # kesimpulan/arahan: seluruh teks slide apa adanya, tanpa narasi model
                points = []
                source_lines = slide.get("text_original") or slide.get("text", [])
            else:
                points = narration.get(ref) or []
                source_lines = [line for block in slide.get("verbatim") or [] for line in (block["label"], *block["lines"])]
            verbatim = verbatim_blocks(source_lines)
            if review_lists:
                report["tinjauan_format"] += [
                    {"ref": ref, **finding} for finding in list_review.review_blocks(source_lines, verbatim)]
            needs_image = bool(slide.get("image_path") and slide_needs_image(slide))
            if not (needs_image or points or verbatim):
                report["slide_dilewati"].append(ref)
                continue
            title = (slide.get("title") or "").strip()
            if (verbatim or (points and not needs_image)) and title and not (slide.get("continues") and title == last_title) \
                    and " ".join(title.lower().split()) != " ".join(chapter.lower().split()):
                blocks.append({"type": "heading", "level": min(heading_level(chapter) + 1, 3), "text": title})
                last_title = title
            if verbatim:
                blocks += verbatim
                report["slide_teks_asli"].append(ref)
            if needs_image:
                # Tiap gambar satu baris tabel dua kolom: gambar | caption (judul panel dan ringkasan gambar).
                units: list[tuple[str, dict[str, Any] | None, str, float]] = []
                panels = slide.get("panel_images") or []
                if len(panels) >= 2:
                    for panel in sorted(panels, key=lambda p: (round(p["bbox"][1], 1), p["bbox"][0])):
                        header = (panel.get("header") or "").strip()
                        caption = numbering.caption_for({**slide, "title": header}, chapter)
                        if caption is None and header:
                            caption = {"text": header, "tipe": "headline"}
                        left, top, right, bottom = panel["bbox"]
                        units.append((panel["image_path"], caption, panel.get("summary") or "", (right - left) * (bottom - top)))
                    report["slide_dipecah"].append(f"{ref} ({len(panels)} panel)")
                else:
                    units.append((slide["image_path"], numbering.caption_for(slide, chapter), slide.get("summary") or "", slide.get("visual_area") or 0.0))
                table_rows: list[list[list[dict[str, Any]]]] = []

                def flush_table() -> None:
                    if table_rows:
                        blocks.append({"type": "table", "widths": widths, "rows": list(table_rows)})
                        table_rows.clear()

                for path, caption, summary, area in units:
                    caption_cell: list[dict[str, Any]] = []
                    if caption:
                        caption_cell.append({"type": "caption", **caption, "align": "left"})
                    if summary:
                        caption_cell.append(paragraph(ref, summary))
                        report["gambar_diringkas"] += 1
                    image_cell = [{"type": "image", "source": path}]
                    if area > FULL_WIDTH_MIN_AREA:  # gambar besar: di luar tabel, selebar halaman, caption dan ringkasan di bawahnya
                        flush_table()
                        blocks += image_cell + caption_cell
                    else:
                        table_rows.append([image_cell, caption_cell])  # kolom kanan tetap ada walau kosong: lebar gambar seragam
                flush_table()
                blocks += [{"type": "paragraph", "text": text} for text in points]
                report["gambar"] += len(units)
            elif points:
                blocks += [{"type": "paragraph", "text": text} for text in points]
            report["narasi"] += bool(points)
            if not needs_image and text_dominant(slide) and (
                    set(slide.get("vision_reasons") or []) & IMAGE_REASONS or slide.get("tables")):
                report["slide_teks"].append(ref)
    return {"blocks": blocks}, report


def build_report(
    spec: dict[str, Any], report: dict[str, Any], warnings: dict[str, list[str]],
    fallback: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Tambahkan verifikasi akhir: setiap gambar slide terpakai tepat satu kali, angka yang tidak didukung,
    dan bab yang narasinya diganti teks slide apa adanya (`fallback`: bab -> alasan)."""
    images = Counter()
    for block in spec["blocks"]:
        for row in block.get("rows", []):
            for cell in row:
                images.update(item["source"] for item in cell if item.get("type") == "image")
    return {
        **report,
        "gambar_ganda": sorted(source for source, n in images.items() if n > 1),
        "angka_tidak_didukung": [{"ref": ref, "angka": numbers} for ref, numbers in warnings.items() if numbers],
        "bab_cadangan": [{"bab": chapter, "alasan": reason} for chapter, reason in (fallback or {}).items()],
    }
