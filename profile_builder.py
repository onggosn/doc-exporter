"""Bangun profil format dari dokumen HTML referensi (hasil CKEditor).

Profil memisahkan *format* dari *fakta*: yang diambil di sini hanya pola layout
tabel/gambar, jenis caption, dan penomorannya. Isi faktual tidak disimpan sebagai
sumber, hanya dipakai sebagai contoh pola.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import statistics
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from bs4 import BeautifulSoup, Tag

import profile_store
import style_guide
from pdf_profile import analyze_pdf
from profile_store import PROFILE_NAME


PROFILE_VERSION = 2
NARASI_MIN_CHARS = 200
HEADLINE_MIN_WORDS = 6
HEADLINE_MIN_LETTERS = 15
UPPERCASE_RATIO = 0.8
WIDTH_BUCKET_PERCENT = 10
MAX_ROLES_IN_NAME = 3
CAPTION_KINDS = ("Grafik", "Tabel", "Gambar", "Foto", "Skema", "Diagram")
WHITE_BACKGROUNDS = {"hsl(0,0%,100%)", "#fff", "#ffffff", "white", "rgb(255,255,255)", "transparent"}
PROFILE_FILES = (
    "meta", "layouts", "row_patterns", "sections", "captions", "numbering", "style_stats", "style_samples",
)
OPTIONAL_PROFILE_FILES = ("page",)
MONTHS = "Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember"

_NUMBERED_CAPTION = re.compile(
    r"^(?:sumber\s*:\s*(?P<source>.*?)\s+)?"
    r"(?P<kind>" + "|".join(CAPTION_KINDS) + r")\s*(?P<number>\d+)\s*(?P<sep>[.:]?)\s*(?P<label>.*)$",
    re.IGNORECASE | re.DOTALL,
)
# Awalan huruf bebas + nomor tanpa spasi, mis. "K1. ..." (poin kesimpulan). Dibatasi
# satu-dua huruf dan satu-dua digit agar "BI 2025: ..." tidak terbaca sebagai nomor.
_PREFIXED_CAPTION = re.compile(r"^(?P<prefix>[A-Za-z]{1,2})(?P<number>\d{1,2})(?P<sep>[.:)])\s*(?P<rest>\S.*)$", re.DOTALL)
_TITLE_LEVEL = re.compile(r"title-level-(\d+)")
_COL_WIDTH = re.compile(r"width\s*:\s*([\d.]+)%")
_BACKGROUND = re.compile(r"background(?:-color)?\s*:\s*([^;]+)")
_FONT_SIZE_PT = re.compile(r"font-size:\s*([\d.]+)pt")
_FONT_FAMILY = re.compile(r"font-family:\s*'?([^;',]+)")
_PERIOD = re.compile(rf"({MONTHS})\s+(\d{{4}})", re.IGNORECASE)


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def _load_html(raw: str) -> BeautifulSoup:
    # Berkas ekspor CKEditor menyimpan atribut dengan tanda kutip yang di-escape (\").
    return BeautifulSoup(raw.replace('\\"', '"'), "html.parser")


def _text_without_figures(tag: Tag) -> str:
    """Teks di dalam tag, tidak termasuk teks figure (gambar + figcaption) di dalamnya.

    Hanya figure yang berada di dalam `tag` yang dikecualikan; figure di luarnya
    (mis. pembungkus `figure.table` milik CKEditor) tidak memengaruhi hasil.
    """
    parts = []
    for string in tag.find_all(string=True):
        for parent in string.parents:
            if parent is tag:
                parts.append(string)
                break
            if parent.name == "figure":
                break
    return _clean(" ".join(parts))


def _title_level(tag: Tag) -> int | None:
    for css_class in tag.get("class", []):
        match = _TITLE_LEVEL.fullmatch(css_class)
        if match:
            return int(match.group(1))
    return None


def _is_shaded(tag: Tag) -> bool:
    match = _BACKGROUND.search(re.sub(r"\s+", "", tag.get("style", "")).lower())
    return bool(match) and match.group(1) not in WHITE_BACKGROUNDS


def classify_caption(text: str) -> dict[str, Any]:
    """Tentukan jenis caption: kosong, label_bernomor, headline, atau label_topik."""
    text = _clean(text)
    if not text:
        return {"type": "kosong"}
    numbered = _NUMBERED_CAPTION.match(text)
    if numbered:
        kind = numbered["kind"].capitalize()
        return {
            "type": "label_bernomor",
            "kind": kind,
            "number": int(numbered["number"]),
            "label": _clean(numbered["label"]),
            "sumber_data": _clean(numbered["source"] or "") or None,
            "penomoran": {"awalan": kind, "nomor": int(numbered["number"]), "pemisah": numbered["sep"]},
        }
    # Penomoran adalah atribut tambahan: caption "K1. <pesan>" tetap sebuah headline.
    prefixed = _PREFIXED_CAPTION.match(text)
    penomoran = None
    if prefixed:
        penomoran = {
            "awalan": prefixed["prefix"].upper(),
            "nomor": int(prefixed["number"]),
            "pemisah": prefixed["sep"],
        }
        text = _clean(prefixed["rest"])
    letters = [c for c in text if c.isalpha()]
    uppercase = sum(c.isupper() for c in letters) / len(letters) if letters else 0.0
    is_headline = (
        uppercase >= UPPERCASE_RATIO and len(letters) >= HEADLINE_MIN_LETTERS
    ) or len(text.split()) >= HEADLINE_MIN_WORDS
    result = {
        "type": "headline" if is_headline else "label_topik",
        "style": "kapital" if uppercase >= UPPERCASE_RATIO else "kalimat",
    }
    if penomoran:
        result.update({"penomoran": penomoran, "teks_tanpa_nomor": text})
    return result


def _describe_cell(cell: Tag) -> dict[str, Any]:
    captions = [_clean(fc.get_text(" ")) for fc in cell.find_all("figcaption")]
    return {
        "img": cell.find("img") is not None,
        "images": len(cell.find_all("img")),
        "text": _text_without_figures(cell),
        "figcaption": " ".join(c for c in captions if c),
        "colspan": int(cell.get("colspan", 1) or 1),
        "shaded": _is_shaded(cell),
    }


def _cell_token(cell: dict[str, Any]) -> str:
    kind = "I" if cell["img"] else ("T" if cell["text"] else "E")
    return kind + (f"*{cell['colspan']}" if cell["colspan"] > 1 else "") + ("s" if cell["shaded"] else "")


def _column_widths(table: Tag) -> list[int]:
    widths = []
    for col in table.find_all("col"):
        match = _COL_WIDTH.search(col.get("style", ""))
        if match:
            widths.append(round(float(match.group(1))))
    return widths


def _bucket_widths(widths: list[int]) -> list[int]:
    """Bulatkan proporsi kolom ke kelipatan 10% agar 41/59 dan 43/57 dianggap sama."""
    total = sum(widths)
    if not total:
        return []
    bucket = [round(w / total * 100 / WIDTH_BUCKET_PERCENT) * WIDTH_BUCKET_PERCENT for w in widths]
    bucket[-1] += 100 - sum(bucket)
    return bucket


def _row_role(row: list[dict[str, Any]]) -> str | None:
    """Peran sebuah baris yang memuat gambar; None bila baris tanpa gambar."""
    if not any(cell["img"] for cell in row):
        return None
    if len(row) == 1:
        return "gambar-penuh"
    if all(cell["img"] for cell in row):
        return "gambar-berdampingan"
    if not any(not cell["img"] and cell["text"] for cell in row):
        return "gambar-saja"
    return "gambar-kiri-narasi-kanan" if row[0]["img"] else "narasi-kiri-gambar-kanan"


def _fold_rows(rows: list[list[dict[str, Any]]]) -> list[dict[str, Any]]:
    """Buang baris kosong penyekat dan lipat baris identik yang berurutan."""
    folded: list[dict[str, Any]] = []
    for row in rows:
        if all(not c["img"] and not c["text"] and not c["shaded"] for c in row):
            continue
        tokens = ",".join(_cell_token(c) for c in row)
        if folded and folded[-1]["tokens"] == tokens:
            folded[-1]["repeat"] += 1
            continue
        folded.append({"tokens": tokens, "repeat": 1, "role": _row_role(row), "row": row})
    return folded


def _layout_name(folded: list[dict[str, Any]]) -> tuple[str, list[str]]:
    roles: list[str] = []
    for entry in folded:
        if entry["role"] and entry["role"] not in roles:
            roles.append(entry["role"])
    if not roles:
        return "tabel-data", []
    first = folded[0]["row"]
    banner = len(first) == 1 and not first[0]["img"] and first[0]["text"] and (
        first[0]["colspan"] > 1 or first[0]["shaded"]
    )
    if len(roles) > MAX_ROLES_IN_NAME:
        return "komposit", roles
    return "+".join((["banner"] if banner else []) + roles), roles


def _table_captions(rows: list[list[dict[str, Any]]]) -> tuple[list[dict[str, Any]], list[int]]:
    captions: list[dict[str, Any]] = []
    narasi_lengths: list[int] = []
    for row in rows:
        for index, cell in enumerate(row):
            if not cell["img"]:
                continue
            parts = [p for p in (cell["text"], cell["figcaption"]) if p]
            if len(parts) == 2:
                source = "figcaption+teks_sel_gambar"
            else:
                source = "figcaption" if cell["figcaption"] else "teks_sel_gambar"
            neighbours = [c for i, c in enumerate(row) if i != index and c["text"] and not c["img"]]
            neighbour_text = neighbours[0]["text"] if neighbours else ""
            if len(neighbour_text) >= NARASI_MIN_CHARS:
                narasi_lengths.append(len(neighbour_text))
            elif not parts and neighbour_text:
                parts, source = [neighbour_text], "sel_sebelah"
            text = " ".join(parts)
            captions.append({"text": text, "source": source if parts else "tidak_ada", **classify_caption(text)})
    return captions, narasi_lengths


def _single_image_item(source: str, section: str, name: str, tag: Tag, caption_text: str) -> dict[str, Any]:
    return {
        "source": source, "section": section, "name": name, "roles": [name], "signature": name,
        "widths": [], "width_bucket": [], "rows": [], "repeats": [], "row_roles": [], "row_count": 0,
        "images": len(tag.find_all("img")),
        "captions": [{"text": caption_text, "source": "figcaption" if caption_text else "tidak_ada",
                      **classify_caption(caption_text)}],
        "narasi_lengths": [],
    }


def count_images(raw: str) -> int:
    return len(_load_html(raw).find_all("img"))


def analyze_html(raw: str, source: str = "") -> list[dict[str, Any]]:
    """Kembalikan daftar item layout (tabel atau gambar tunggal) berurutan sesuai dokumen."""
    soup = _load_html(raw)
    path: dict[int, str] = {}
    items: list[dict[str, Any]] = []
    for tag in soup.find_all(["div", "table", "figure", "p"]):
        if tag.find_parent("table") is not None:
            continue
        if tag.name == "div":
            level = _title_level(tag)
            if level is not None:
                path = {lvl: title for lvl, title in path.items() if lvl < level}
                path[level] = _text_without_figures(tag)
            continue
        section = " > ".join(path[lvl] for lvl in sorted(path)) or "(tanpa bagian)"
        if tag.name == "p":
            # Gambar sebaris dalam paragraf, tanpa figure.
            if tag.find_parent(["table", "figure"]) is None and tag.find("img") is not None:
                items.append(_single_image_item(source, section, "gambar-sebaris", tag, ""))
            continue
        if tag.name == "figure":
            # figure.table hanyalah pembungkus tabel dari CKEditor; tabelnya dihitung sendiri.
            if "table" in tag.get("class", []) or tag.find("table") is not None or tag.find("img") is None:
                continue
            fig_caption = " ".join(_clean(fc.get_text(" ")) for fc in tag.find_all("figcaption"))
            items.append(_single_image_item(source, section, "gambar-tunggal", tag, fig_caption))
            continue
        rows = [
            [_describe_cell(c) for c in tr.find_all(["td", "th"], recursive=False)]
            for tr in tag.find_all("tr")
        ]
        rows = [row for row in rows if row]
        folded = _fold_rows(rows)
        if not folded:
            continue
        widths = _column_widths(tag)
        bucket = _bucket_widths(widths)
        name, roles = _layout_name(folded)
        captions, narasi_lengths = _table_captions(rows)
        items.append({
            "source": source, "section": section, "name": name, "roles": roles,
            "signature": f"w={'/'.join(map(str, bucket)) or 'auto'}|" + ";".join(f["tokens"] for f in folded),
            "widths": widths, "width_bucket": bucket,
            "rows": [f["tokens"] for f in folded], "repeats": [f["repeat"] for f in folded],
            "row_roles": [f["role"] for f in folded], "row_count": len(rows),
            "images": len(tag.find_all("img")),
            "captions": captions, "narasi_lengths": narasi_lengths,
        })
    return items


def _distribution(captions: list[dict[str, Any]]) -> dict[str, Any]:
    counts = Counter(c["type"] for c in captions)
    non_empty = Counter({t: n for t, n in counts.items() if t != "kosong"})
    return {
        "counts": dict(counts),
        "dominan": non_empty.most_common(1)[0][0] if non_empty else "kosong",
        "bernomor": sum(1 for c in captions if c.get("penomoran")),
    }


def _numbering(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Urutan penomoran per awalan (Grafik, Tabel, K, ...). Hanya informasi, bukan aturan keras."""
    sequences: dict[tuple[str, str], list[int]] = defaultdict(list)
    for item in items:
        for caption in item["captions"]:
            numbering = caption.get("penomoran")
            if numbering:
                sequences[(item["source"], numbering["awalan"])].append(numbering["nomor"])
    report = []
    for (source, awalan), numbers in sorted(sequences.items()):
        report.append({
            "source": source, "awalan": awalan, "count": len(numbers),
            "first": numbers[0], "last": numbers[-1],
            "sequential": numbers == list(range(numbers[0], numbers[0] + len(numbers))),
            "duplicates": sorted(n for n, c in Counter(numbers).items() if c > 1),
            "restarts": sum(1 for a, b in zip(numbers, numbers[1:]) if b < a),
        })
    return report


def _row_patterns(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Pola baris bergambar di seluruh tabel; satuan terkecil yang bisa dipilih per slide."""
    patterns: dict[str, dict[str, Any]] = {}
    for item in items:
        for tokens, repeat, role in zip(item["rows"], item["repeats"], item["row_roles"]):
            if role is None:
                continue
            entry = patterns.setdefault(tokens, {
                "tokens": tokens, "role": role, "rows": 0, "tables": 0, "width_buckets": Counter(),
            })
            entry["rows"] += repeat
            entry["tables"] += 1
            entry["width_buckets"]["/".join(map(str, item["width_bucket"])) or "auto"] += 1
    result = sorted(patterns.values(), key=lambda p: -p["rows"])
    for entry in result:
        entry["width_buckets"] = dict(entry["width_buckets"].most_common())
    return result


def html_typography(raw: str, source: str) -> dict[str, Any]:
    """Font dan ukuran yang dideklarasikan di HTML, dibobot jumlah karakter teksnya."""
    sizes: Counter[float] = Counter()
    families: Counter[str] = Counter()
    for tag in _load_html(raw).find_all(style=True):
        style = re.sub(r"\s+", " ", tag["style"])
        chars = len(_clean(" ".join(tag.find_all(string=True, recursive=False))))
        if not chars:
            continue
        size = _FONT_SIZE_PT.search(style)
        family = _FONT_FAMILY.search(style)
        if size:
            sizes[float(size.group(1))] += chars
        if family:
            families[family.group(1).strip()] += chars

    def shares(counter: Counter, key: str) -> list[dict[str, Any]]:
        total = sum(counter.values()) or 1
        return [{key: name, "bagian": round(n / total, 3)} for name, n in counter.most_common()]

    return {"file": source, "ukuran": shares(sizes, "pt"), "font": shares(families, "nama")}


def _period_key(filename: str) -> str | None:
    match = _PERIOD.search(filename)
    return f"{match.group(1).title()} {match.group(2)}" if match else None


def _print_scale(pdfs: list[dict[str, Any]], typography: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Bandingkan ukuran isi PDF dengan ukuran HTML dokumen yang sama (dipasangkan per periode)."""
    result = []
    for pdf in pdfs:
        period = _period_key(pdf["file"])
        html = next((t for t in typography if period and _period_key(t["file"]) == period), None)
        body = pdf["tipografi_pdf"]["isi_dominan"]
        if not (html and html["ukuran"] and body):
            continue
        html_size = html["ukuran"][0]["pt"]
        result.append({
            "dokumen": period, "pdf": pdf["file"], "html": html["file"],
            "ukuran_html_pt": html_size, "ukuran_pdf_pt": body["ukuran_pt"],
            "skala": round(body["ukuran_pt"] / html_size, 2),
        })
    return result


def assemble_page_profile(pdfs: list[dict[str, Any]], typography: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "pdf": pdfs,
        "html": typography,
        "skala_cetak": _print_scale(pdfs, typography),
        "catatan": "Skala cetak di bawah 1 berarti browser menyusutkan halaman saat mencetak; "
                   "ukuran font dokumen mengikuti HTML, bukan PDF.",
    }


def build_page_profile(pdf_paths: list[Path], typography: list[dict[str, Any]]) -> dict[str, Any]:
    return assemble_page_profile([analyze_pdf(path) for path in pdf_paths], typography)


class ProfileBuildError(ValueError):
    """Permintaan membangun atau memperbarui profil tidak bisa dijalankan."""


def read_html(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def html_sha(raw: str) -> str:
    return hashlib.sha256(raw.encode()).hexdigest()


def pdf_sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def html_outline(soup: BeautifulSoup) -> list[dict[str, Any]]:
    """Judul bab (tingkat 1) dan sub-bab (tingkat 2) dokumen referensi, berurutan."""
    outline = []
    for tag in soup.find_all("div"):
        level = _title_level(tag)
        title = _text_without_figures(tag) if level in (1, 2) else ""
        if title:
            outline.append({"level": level, "title": title})
    return outline


def analyze_html_source(path: Path, raw: str | None = None) -> dict[str, Any]:
    """Hasil analisis satu HTML referensi. Disimpan di sumber/ supaya profil bisa diperbarui tanpa membaca ulang
    dokumen lama, dan semua berkas profil dirakit ulang dari kumpulan hasil ini (assemble_profile)."""
    raw = read_html(path) if raw is None else raw
    soup = _load_html(raw)
    items = analyze_html(raw, path.name)
    return {
        "jenis": "html", "file": path.name, "sha256": html_sha(raw), "images_in_html": count_images(raw),
        "items": items, "blocks": style_guide.extract_blocks(soup, path.name),
        "italics": style_guide.italic_terms(soup), "typography": html_typography(raw, path.name),
        "outline": html_outline(soup),
    }


def analyze_pdf_source(path: Path) -> dict[str, Any]:
    return {"jenis": "pdf", "file": path.name, "sha256": pdf_sha(path), "analisis": analyze_pdf(path)}


def assemble_profile(
    html_sources: list[dict[str, Any]], pdf_sources: list[dict[str, Any]], title: str = ""
) -> dict[str, Any]:
    """Rakit semua data profil dari hasil analisis sumber (urutan sumber menentukan urutan hasil)."""
    items: list[dict[str, Any]] = []
    sources = []
    typography = []
    blocks: list[dict[str, str]] = []
    italics: list[str] = []
    for source in html_sources:
        blocks += source["blocks"]
        italics += source["italics"]
        typography.append(source["typography"])
        items.extend(source["items"])
        sources.append({
            "file": source["file"], "sha256": source["sha256"], "images_in_html": source["images_in_html"],
            "images_in_layouts": sum(item["images"] for item in source["items"]),
        })

    layouts: dict[str, dict[str, Any]] = {}
    for item in items:
        layout_id = f"{item['name']}:{hashlib.sha1(item['signature'].encode()).hexdigest()[:6]}"
        layout = layouts.setdefault(layout_id, {
            "id": layout_id, "name": item["name"], "roles": item["roles"],
            "width_bucket": item["width_bucket"], "rows": item["rows"], "example_repeats": item["repeats"],
            "count": 0, "sections": Counter(), "_widths": Counter(), "_row_counts": [],
            "_captions": [], "_narasi": [],
        })
        layout["count"] += 1
        layout["sections"][item["section"]] += 1
        layout["_widths"]["/".join(map(str, item["widths"])) or "auto"] += 1
        layout["_row_counts"].append(item["row_count"])
        layout["_captions"] += item["captions"]
        layout["_narasi"] += item["narasi_lengths"]
        item["layout_id"] = layout_id

    layout_list = []
    for layout in sorted(layouts.values(), key=lambda l: -l["count"]):
        narasi = layout.pop("_narasi")
        row_counts = layout.pop("_row_counts")
        layout["row_count_range"] = [min(row_counts), max(row_counts)]
        layout["widths_seen"] = dict(layout.pop("_widths").most_common(5))
        layout["caption"] = _distribution(layout.pop("_captions"))
        layout["narasi_chars"] = (
            {"min": min(narasi), "median": int(statistics.median(narasi)), "max": max(narasi)}
            if narasi else None
        )
        layout["sections"] = dict(layout["sections"].most_common(5))
        layout_list.append(layout)

    sections: dict[str, dict[str, Any]] = {}
    for item in items:
        entry = sections.setdefault(item["section"], {"layouts": Counter(), "_captions": []})
        entry["layouts"][item["layout_id"]] += 1
        entry["_captions"] += item["captions"]
    section_list = [
        {"title": title_, "layouts": dict(entry["layouts"]), "caption": _distribution(entry["_captions"])}
        for title_, entry in sections.items()
    ]

    captions = [
        {"source_file": item["source"], "section": item["section"], "layout_id": item["layout_id"], **caption}
        for item in items for caption in item["captions"]
    ]
    profile = {
        "meta": {
            "version": PROFILE_VERSION,
            "title": title,
            "created_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "sources": sources,
        },
        "layouts": layout_list,
        "row_patterns": _row_patterns(items),
        "sections": section_list,
        "captions": captions,
        "numbering": _numbering(items),
        "style_stats": style_guide.compute_statistics(blocks, captions, italics),
        "style_samples": style_guide.select_samples(blocks, captions),
    }
    if pdf_sources:
        profile["meta"]["pdf_sources"] = [{"file": src["file"], "sha256": src["sha256"]} for src in pdf_sources]
        profile["page"] = assemble_page_profile([src["analisis"] for src in pdf_sources], typography)
    return profile


def build_profile(paths: list[Path], title: str = "", pdf_paths: list[Path] | None = None) -> dict[str, Any]:
    """Bangun profil dari nol dari HTML (dan PDF) yang disebut, tanpa menyimpan apa pun."""
    return assemble_profile(
        [analyze_html_source(path) for path in paths], [analyze_pdf_source(path) for path in pdf_paths or []], title)


def save_profile(profile: dict[str, Any], out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for key in PROFILE_FILES + tuple(k for k in OPTIONAL_PROFILE_FILES if k in profile):
        _write_json(out_dir / f"{key}.json", profile[key])
    for key in OPTIONAL_PROFILE_FILES:  # berkas opsional yang datanya sudah tidak ada (mis. semua PDF dihapus)
        if key not in profile:
            (out_dir / f"{key}.json").unlink(missing_ok=True)


def _write_json(path: Path, data: Any) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8")) if path.is_file() else None


# --- Pembuatan dan pembaruan profil: satu pintu untuk semua berkas profil ------------------------------------

SOURCE_DIR = "sumber"  # hasil analisis per dokumen referensi; dasar pembaruan tanpa membaca ulang dokumen lama
MAX_HISTORY = 20
GENERATED_BY = "profile_builder.py"
DEFAULT_SECTION_NAME = "Default Section"
MONTH_INDEX = {m.lower(): i for i, m in enumerate(MONTHS.split("|"), start=1)}


def existing_document_types(root: Path) -> list[str]:
    """Tipe dokumen yang sudah punya profil di `root` (folder berisi meta.json)."""
    if not root.is_dir():
        return []
    return sorted(d.name for d in root.iterdir() if d.is_dir() and (d / "meta.json").is_file())


def document_type_exists(name: str, root: Path) -> bool:
    return (root / name / "meta.json").is_file()


def _validate_name(name: str, root: Path) -> None:
    if not PROFILE_NAME.fullmatch(name or ""):
        known = ", ".join(existing_document_types(root)) or "belum ada"
        raise ProfileBuildError(
            f"Nama tipe dokumen '{name}' tidak valid: pakai huruf kecil, angka, dan tanda hubung (mis. dokumen-2d). "
            f"Tipe yang sudah ada: {known}.")


def _period_sort_key(filename: str) -> tuple[int, int]:
    match = _PERIOD.search(filename)
    return (int(match.group(2)), MONTH_INDEX.get(match.group(1).lower(), 0)) if match else (0, 0)


def _source_cache_path(profile_dir: Path, source: dict[str, Any]) -> Path:
    slug = re.sub(r"[^A-Za-z0-9]+", "-", Path(source["file"]).stem).strip("-")[:60] or "sumber"
    return profile_dir / SOURCE_DIR / f"{source['jenis']}-{slug}-{source['sha256'][:8]}.json"


def load_cached_sources(profile_dir: Path) -> dict[tuple[str, str], dict[str, Any]]:
    cached: dict[tuple[str, str], dict[str, Any]] = {}
    for path in sorted((profile_dir / SOURCE_DIR).glob("*.json")) if (profile_dir / SOURCE_DIR).is_dir() else []:
        try:
            source = json.loads(path.read_text(encoding="utf-8"))
            cached[(source["jenis"], source["file"])] = source
        except (json.JSONDecodeError, KeyError, OSError):
            continue  # cache rusak dianggap tidak ada: dokumennya harus disertakan lagi
    return cached


def _write_source_cache(profile_dir: Path, sources: list[dict[str, Any]]) -> None:
    directory = profile_dir / SOURCE_DIR
    directory.mkdir(parents=True, exist_ok=True)
    wanted = {_source_cache_path(profile_dir, src) for src in sources}
    for old in directory.glob("*.json"):
        if old not in wanted:
            old.unlink()
    for src in sources:
        _write_json(_source_cache_path(profile_dir, src), src)


def derive_chapter_structures(html_sources: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Struktur bab tiap dokumen referensi (judul tingkat 1 dan 2), dinamai menurut periodenya."""
    structures = []
    for source in html_sources:
        titles = [" ".join(entry["title"].split()) for entry in source["outline"]]
        name = _period_key(source["file"]) or Path(source["file"]).stem
        if titles and len(titles) <= 20:
            structures.append({"nama": name[:60], "bab": titles, "_urut": _period_sort_key(source["file"])})
    return structures


def merge_chapters(existing: dict[str, Any] | None, derived: list[dict[str, Any]]) -> dict[str, Any]:
    """Gabungkan struktur dari referensi ke chapters.json: struktur bernama sama diganti, yang baru ditambah,
    struktur lain (mis. buatan menu) dipertahankan. Bawaan tetap bila masih ada, kalau tidak periode terbaru."""
    daftar = [dict(item) for item in (existing or {}).get("daftar", [])]
    by_name = {item["nama"].lower(): index for index, item in enumerate(daftar)}
    for item in sorted(derived, key=lambda d: d["_urut"]):
        entry = {"nama": item["nama"], "bab": item["bab"]}
        if item["nama"].lower() in by_name:
            daftar[by_name[item["nama"].lower()]] = entry
        else:
            by_name[item["nama"].lower()] = len(daftar)
            daftar.append(entry)
    default = (existing or {}).get("bawaan", "")
    latest = max(derived, key=lambda d: d["_urut"])["nama"] if derived else ""
    keep = default if default.lower() in by_name else latest
    while len(daftar) > 10:  # batas jumlah struktur per profil: buang yang tertua dari yang bukan bawaan
        drop = next((i for i, item in enumerate(daftar) if item["nama"].lower() != keep.lower()), None)
        if drop is None:
            break
        del daftar[drop]
    return {"bawaan": keep, "daftar": daftar}


def _first_number(title: str) -> str | None:
    match = re.match(r"\s*(\d+(?:\.\d+)*)\.?\s", title)
    return match.group(1) if match else None


def _strip_number(title: str) -> str:
    return re.sub(r"^\s*\d+(?:\.\d+)*\.?\s+", "", title).strip()


def derive_default_section(html_sources: list[dict[str, Any]]) -> dict[str, Any] | None:
    """Pengaturan pemecahan slide berkolom di section bawaan, dari bab pertama dokumen referensi terbaru yang
    punya dua sub-bab atau lebih (mis. "1.1 Kesimpulan dan Arahan Moneter", "1.2 ... Makroprudensial"). `cocok`
    = nama kolom huruf kapital, atau singkatannya bila lebih dari satu kata ("Sistem Pembayaran" -> "SP").
    None bila tidak ada bab seperti itu atau singkatannya bentrok."""
    for source in sorted(html_sources, key=lambda s: _period_sort_key(s["file"]), reverse=True):
        outline = source["outline"]
        first = next((i for i, e in enumerate(outline) if e["level"] == 1), None)
        if first is None:
            continue
        children = []
        for entry in outline[first + 1:]:
            if entry["level"] == 1:
                break
            children.append(_strip_number(entry["title"]))
        if len(children) < 2:
            continue
        split = [child.split() for child in children]
        common = 0
        while all(len(words) > common + 1 for words in split) and len({words[common].lower() for words in split}) == 1:
            common += 1
        names = [" ".join(words[common:]) for words in split]
        keys = [name.upper() if len(name.split()) == 1 else "".join(w[0] for w in name.split()).upper() for name in names]
        if len(set(keys)) != len(keys):
            return None
        chapter = " ".join(outline[first]["title"].split())
        number = _first_number(chapter)
        prefix = " ".join(split[0][:common])
        pattern = (f"{number}.{{n}} " if number else "") + (f"{prefix} " if prefix else "") + "{nama}"
        return {
            "dibuat_oleh": GENERATED_BY, "bagian_cocok": DEFAULT_SECTION_NAME, "judul_bab": chapter,
            "judul_kolom": pattern, "kolom": [{"cocok": k, "nama": n} for k, n in zip(keys, names)],
        }
    return None


def merge_settings(existing: dict[str, Any], captions: list[dict[str, Any]]) -> tuple[dict[str, Any], list[str]]:
    """Tambahkan aturan penomoran dari referensi yang belum ada di settings.json yang tersimpan; aturan yang sudah
    ada (termasuk yang diubah lewat menu) tidak diubah."""
    derived = profile_store.default_settings(captions)["penomoran"]
    known = {(r["bagian_cocok"].lower(), r["awalan"].lower()) for r in existing.get("penomoran", [])}
    added = [r for r in derived if (r["bagian_cocok"].lower(), r["awalan"].lower()) not in known]
    merged = {**existing, "penomoran": [*existing.get("penomoran", []), *added]}
    return merged, [f"{r['awalan']} @ {r['bagian_cocok']}" for r in added]


def _count_by(items: list[dict[str, Any]], key: str) -> dict[str, int]:
    return dict(Counter(item.get(key, "?") for item in items))


def _diff_summary(old: dict[str, Any], new: dict[str, Any]) -> dict[str, Any]:
    old_prefix = {(row["awalan"]) for row in old.get("numbering") or []}
    new_prefix = {row["awalan"] for row in new["numbering"]}
    return {
        "layout": [len(old.get("layouts") or []), len(new["layouts"])],
        "caption": [len(old.get("captions") or []), len(new["captions"])],
        "blok_sampel": [len(old.get("style_samples") or []), len(new["style_samples"])],
        "awalan_penomoran_baru": sorted(new_prefix - old_prefix),
        "awalan_penomoran_hilang": sorted(old_prefix - new_prefix),
        "jenis_caption": _count_by(new["captions"], "type"),
        "sampel_berubah": json.dumps(old.get("style_samples"), sort_keys=True) != json.dumps(new["style_samples"], sort_keys=True),
        "statistik_berubah": json.dumps(old.get("style_stats"), sort_keys=True) != json.dumps(new["style_stats"], sort_keys=True),
    }


def _read_old_profile(profile_dir: Path) -> dict[str, Any]:
    keys = PROFILE_FILES + OPTIONAL_PROFILE_FILES
    return {key: _read_json(profile_dir / f"{key}.json") for key in keys}


def update_profile(
    name: str, html_paths: list[Path], pdf_paths: list[Path], *, root: Path, title: str = "",
    remove: list[str] | None = None, replace_all: bool = False, force: bool = False, dry_run: bool = False,
) -> dict[str, Any]:
    """Buat profil tipe dokumen `name` bila belum ada di `root`, atau perbarui bila sudah ada.

    Pembaruan menggabungkan dokumen referensi baru dengan dokumen yang sudah dianalisis (tersimpan di sumber/):
    dokumen yang sama persis (nama dan hash) dilewati, yang isinya berubah diganti, yang baru ditambah, yang tidak
    disebut dipertahankan (kecuali `remove`, atau `replace_all` yang hanya memakai dokumen yang disebut). Semua berkas
    profil lalu dirakit ulang dari kumpulan itu: hasil analisis, struktur bab (chapters.json), pengaturan section
    bawaan (default_section.json), dan penambahan aturan penomoran baru ke settings.json bila ada.
    Mengembalikan laporan perubahan."""
    _validate_name(name, root)
    profile_dir = root / name
    exists = document_type_exists(name, root)
    if not exists and (remove or replace_all):
        raise ProfileBuildError(f"Tipe dokumen '{name}' belum ada; --hapus dan --ganti hanya untuk pembaruan.")
    cached = load_cached_sources(profile_dir) if exists else {}
    old_meta = _read_json(profile_dir / "meta.json") if exists else None

    provided: dict[tuple[str, str], tuple[Path, str, str | None]] = {}  # kunci -> (jalur, sha, isi html)
    for path in html_paths:
        raw = read_html(path)
        provided[("html", path.name)] = (path, html_sha(raw), raw)
    for path in pdf_paths:
        provided[("pdf", path.name)] = (path, pdf_sha(path), None)
    if len(provided) != len(html_paths) + len(pdf_paths):
        raise ProfileBuildError("Ada dua dokumen dengan nama berkas yang sama; ganti nama salah satunya.")
    remove_names = set(remove or [])
    unknown = remove_names - {file for _, file in cached} - {s["file"] for s in (old_meta or {}).get("sources", [])} \
        - {s["file"] for s in (old_meta or {}).get("pdf_sources", [])}
    if unknown:
        raise ProfileBuildError(f"Dokumen yang akan dihapus tidak ada di profil: {', '.join(sorted(unknown))}.")

    legacy = {}  # dokumen yang tercatat di meta lama tetapi belum punya hasil analisis tersimpan
    for kind, entries in (("html", (old_meta or {}).get("sources", [])), ("pdf", (old_meta or {}).get("pdf_sources", []))):
        for entry in entries:
            if (kind, entry["file"]) not in cached:
                legacy[(kind, entry["file"])] = entry["sha256"]

    report: dict[str, Any] = {"mode": "update" if exists else "baru", "tipe": name, "ditambah": [], "diganti": [],
                              "tetap": [], "dipertahankan": [], "dihapus": [], "dicatat_ke_cache": []}
    kept: dict[tuple[str, str], dict[str, Any]] = {}
    for key, source in cached.items():
        if key[1] in remove_names:
            report["dihapus"].append(key[1])
        elif key not in provided:
            if replace_all:
                report["dihapus"].append(key[1])
            else:
                kept[key] = source
                report["dipertahankan"].append(key[1])
    missing_legacy = sorted(
        file for (kind, file) in legacy if (kind, file) not in provided and file not in remove_names)
    if missing_legacy and not replace_all:
        raise ProfileBuildError(
            f"Profil '{name}' dibuat sebelum hasil analisis dokumen disimpan, sehingga dokumen berikut harus disertakan "
            f"sekali lagi agar tidak hilang dari profil: {', '.join(missing_legacy)}. "
            "Sebutkan semua dokumen referensi (HTML dan PDF); atau pakai --ganti untuk membangun ulang hanya dari "
            "dokumen yang disebut.")

    sources: dict[tuple[str, str], dict[str, Any]] = dict(kept)
    for key, (path, sha, raw) in provided.items():
        old = cached.get(key)
        if old is not None and old["sha256"] == sha:
            sources[key] = old
            report["tetap"].append(path.name)
            continue
        sources[key] = analyze_html_source(path, raw) if key[0] == "html" else analyze_pdf_source(path)
        if old is not None:
            report["diganti"].append(path.name)
        elif key in legacy and legacy[key] == sha:
            report["dicatat_ke_cache"].append(path.name)
        else:
            report["ditambah"].append(path.name)

    html_sources = sorted((s for s in sources.values() if s["jenis"] == "html"), key=lambda s: s["file"])
    pdf_sources = sorted((s for s in sources.values() if s["jenis"] == "pdf"), key=lambda s: s["file"])
    if not html_sources:
        raise ProfileBuildError("Profil membutuhkan sedikitnya satu dokumen referensi HTML.")
    changed = bool(report["ditambah"] or report["diganti"] or report["dihapus"] or report["dicatat_ke_cache"])
    report["ada_perubahan"] = changed or not exists
    if exists and not changed and not force:
        report["berkas"] = {"ditulis": [], "dibiarkan": "semua (tidak ada dokumen baru atau berubah; pakai --paksa untuk merakit ulang)"}
        return report

    old_profile = _read_old_profile(profile_dir) if exists else {}
    keep_title = (old_meta or {}).get("title") or name
    profile = assemble_profile(html_sources, pdf_sources, title or keep_title)
    history = [*((old_meta or {}).get("riwayat") or []), {
        "waktu": profile["meta"]["created_at"], "mode": report["mode"],
        "ditambah": report["ditambah"], "diganti": report["diganti"], "dihapus": report["dihapus"],
    }][-MAX_HISTORY:]
    profile["meta"]["riwayat"] = history
    if old_meta:
        profile["meta"]["created_at"] = old_meta.get("created_at", profile["meta"]["created_at"])
        profile["meta"]["diperbarui_pada"] = history[-1]["waktu"]
    report["perubahan"] = _diff_summary(old_profile, profile)

    chapters_existing = None
    if (profile_dir / "chapters.json").is_file():
        try:
            chapters_existing = profile_store.validate_chapters(_read_json(profile_dir / "chapters.json"))
        except profile_store.ProfileError:
            report["catatan"] = report.get("catatan", []) + ["chapters.json lama rusak dan diganti."]
    chapters = merge_chapters(chapters_existing, derive_chapter_structures(html_sources))
    chapters_clean = profile_store.validate_chapters(chapters)
    default_section = derive_default_section(html_sources)
    ds_path = profile_dir / "default_section.json"
    ds_old = _read_json(ds_path) if ds_path.is_file() else None
    if default_section is None:
        ds_status = "dibiarkan (tidak ada bab pertama dengan sub-bab di referensi terbaru)"
    elif ds_old is not None and ds_old.get("dibuat_oleh") != GENERATED_BY:
        ds_status = "dibiarkan (berkas lama tidak bertanda dibuat_oleh; hapus berkasnya agar dibuat ulang oleh skrip)"
        default_section = None
    else:
        ds_status = "dibuat" if ds_old is None else ("diperbarui" if ds_old != default_section else "tidak berubah")
    settings_status = "tidak ada settings.json (aturan penomoran diturunkan otomatis dari caption saat profil dimuat)"
    settings_new = None
    if (profile_dir / "settings.json").is_file():
        try:
            settings_old = profile_store.validate_settings(_read_json(profile_dir / "settings.json"))
            merged, added = merge_settings(settings_old, profile["captions"])
            settings_new = profile_store.validate_settings(merged) if added else None
            settings_status = f"aturan baru ditambahkan: {', '.join(added)}" if added else "tidak berubah"
        except profile_store.ProfileError:
            settings_status = "dibiarkan (settings.json lama tidak valid)"
    report["berkas"] = {
        "dirakit_ulang": [f"{k}.json" for k in PROFILE_FILES + (("page",) if "page" in profile else ())],
        "page.json": "dirakit dari PDF yang tersimpan" if "page" in profile else "dihapus/tidak dibuat (tidak ada PDF)",
        "chapters.json": "struktur per periode ditambahkan atau diganti",
        "default_section.json": ds_status,
        "settings.json": settings_status,
        "dibiarkan": ["style_guide.json (butuh model; pakai --panduan-gaya atau tombol di menu)"],
    }
    report["struktur_bab"] = [item["nama"] for item in chapters_clean["daftar"]]
    if dry_run:
        report["dry_run"] = True
        return report

    save_profile(profile, profile_dir)
    _write_source_cache(profile_dir, [*html_sources, *pdf_sources])
    _write_json(profile_dir / "chapters.json", chapters_clean)
    if default_section is not None:
        _write_json(ds_path, default_section)
    if settings_new is not None:
        _write_json(profile_dir / "settings.json", settings_new)
    guide_path = profile_dir / "style_guide.json"
    if guide_path.is_file():
        stale = profile_store._style_summary(profile_dir)["usang"]
        report["panduan_gaya"] = "usang (sampel berubah); buat ulang dengan --panduan-gaya" if stale else "masih sesuai sampel"
    else:
        report["panduan_gaya"] = "belum ada; buat dengan --panduan-gaya atau tombol di menu"
    return report


def generate_style_guide_for(name: str, root: Path, model_id: str | None, region: str | None) -> dict[str, Any]:
    """Buat style_guide.json lewat Bedrock dari sampel dan statistik profil (sama dengan tombol Buat Panduan Gaya).
    Sampel referensi dikirim ke model, jadi hanya dijalankan atas permintaan eksplisit."""
    from bedrock_service import BedrockDocumentGenerator  # impor lambat: butuh boto3 dan kredensial AWS

    profile_store.PROFILES_DIR = root
    samples, statistics_ = profile_store.style_inputs(name)
    generator = BedrockDocumentGenerator(model_id=model_id, region=region)
    guide, dropped = generator.generate_style_guide(samples, statistics_)
    return profile_store.save_style_guide(name, guide, generator.model_id, dropped)


def _print_report(report: dict[str, Any]) -> None:
    mode = "TIPE DOKUMEN BARU DIBUAT" if report["mode"] == "baru" else "TIPE DOKUMEN SUDAH ADA, DIPERBARUI"
    print(f"{mode}: {report['tipe']}")
    for label, key in (("ditambah", "ditambah"), ("diganti (isi berubah)", "diganti"), ("dicatat ke cache (sudah ada di profil)", "dicatat_ke_cache"),
                       ("tidak berubah", "tetap"), ("dipertahankan dari hasil lama", "dipertahankan"), ("dihapus", "dihapus")):
        if report.get(key):
            print(f"  dokumen {label}: {', '.join(report[key])}")
    if report.get("perubahan"):
        change = report["perubahan"]
        print(f"  layout {change['layout'][0]} -> {change['layout'][1]}, caption {change['caption'][0]} -> {change['caption'][1]}, "
              f"blok sampel {change['blok_sampel'][0]} -> {change['blok_sampel'][1]}")
        if change["awalan_penomoran_baru"]:
            print(f"  awalan penomoran baru: {', '.join(change['awalan_penomoran_baru'])}")
        if change["awalan_penomoran_hilang"]:
            print(f"  awalan penomoran yang tidak lagi ada: {', '.join(change['awalan_penomoran_hilang'])}")
        print(f"  sampel gaya {'berubah' if change['sampel_berubah'] else 'sama'}, statistik gaya {'berubah' if change['statistik_berubah'] else 'sama'}")
    files = report.get("berkas") or {}
    for key, value in files.items():
        print(f"  {key}: {', '.join(value) if isinstance(value, list) else value}")
    if report.get("struktur_bab"):
        print(f"  struktur bab tersimpan: {', '.join(report['struktur_bab'])}")
    if report.get("panduan_gaya"):
        print(f"  panduan gaya: {report['panduan_gaya']}")
    for note in report.get("catatan", []):
        print(f"  catatan: {note}")
    if report.get("dry_run"):
        print("  (hanya pemeriksaan: tidak ada berkas yang ditulis)")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Buat atau perbarui profil tipe dokumen dari dokumen referensi. Tipe yang belum ada di folder "
                    "profiles dibuat; yang sudah ada diperbarui dengan menggabungkan dokumen baru.")
    parser.add_argument("html", nargs="*", type=Path, help="HTML referensi (baru, atau semua bila profil lama belum punya cache sumber)")
    parser.add_argument("--name", required=True, help="Tipe dokumen, juga nama folder di profiles/ (mis. dokumen-2d)")
    parser.add_argument("--title", default="", help="Nama tampilan tipe dokumen (mis. 'Dokumen 2d'); bawaan: judul lama atau nama")
    parser.add_argument("--pdf", nargs="*", type=Path, default=[], help="PDF referensi untuk properti halaman")
    parser.add_argument("--out", type=Path, default=Path("profiles"))
    parser.add_argument("--hapus", nargs="*", default=[], metavar="BERKAS", help="Nama berkas referensi yang dikeluarkan dari profil")
    parser.add_argument("--ganti", action="store_true", help="Bangun ulang hanya dari dokumen yang disebut (dokumen lain dibuang)")
    parser.add_argument("--paksa", action="store_true", help="Rakit ulang semua berkas walau tidak ada dokumen baru")
    parser.add_argument("--cek", action="store_true", help="Hanya periksa dan tampilkan perubahan, tanpa menulis berkas")
    parser.add_argument("--panduan-gaya", action="store_true", help="Buat ulang style_guide.json lewat Bedrock (mengirim sampel referensi ke model)")
    parser.add_argument("--model-id", default="", help="Model ID Bedrock untuk --panduan-gaya")
    parser.add_argument("--region", default="", help="Region AWS untuk --panduan-gaya")
    args = parser.parse_args(argv)

    types = existing_document_types(args.out)
    print(f"Tipe dokumen yang sudah ada di {args.out}: {', '.join(types) or 'belum ada'}")
    try:
        report = update_profile(
            args.name, args.html, args.pdf, root=args.out, title=args.title, remove=args.hapus,
            replace_all=args.ganti, force=args.paksa, dry_run=args.cek)
        if args.panduan_gaya and not args.cek:
            if not (args.out / args.name / "meta.json").is_file():
                raise ProfileBuildError("Profil belum terbentuk; panduan gaya tidak bisa dibuat.")
            generate_style_guide_for(args.name, args.out, args.model_id or None, args.region or None)
            report["panduan_gaya"] = "dibuat ulang"
    except ProfileBuildError as exc:
        print(f"GAGAL: {exc}")
        return 1
    _print_report(report)
    if args.cek or not (args.out / args.name / "meta.json").is_file():
        return 0
    profile = {key: _read_json(args.out / args.name / f"{key}.json") for key in PROFILE_FILES + OPTIONAL_PROFILE_FILES}
    for pdf in (profile.get("page") or {}).get("pdf", []):
        margin = pdf["margin_pt"]
        print(f"halaman {pdf['file'][:34]}: {pdf['ukuran']['nama']} margin K{margin['kiri']}/R{margin['kanan']}"
              f"/A{margin['atas']}/B{margin['bawah']} watermark={(pdf['watermark'] or {}).get('teks')}")
    for scale in (profile.get("page") or {}).get("skala_cetak", []):
        print(f"skala cetak {scale['dokumen']}: HTML {scale['ukuran_html_pt']}pt, PDF {scale['ukuran_pdf_pt']}pt "
              f"-> {scale['skala']}")
    for source in profile["meta"]["sources"]:
        ok = source["images_in_html"] == source["images_in_layouts"]
        print(f"gambar {source['file'][:40]}: html={source['images_in_html']} "
              f"layout={source['images_in_layouts']} {'OK' if ok else 'TIDAK COCOK'}")
    for layout in profile["layouts"]:
        print(f"{layout['count']:>4}  {layout['id']:<54} caption: {layout['caption']['counts']}")
    for row in profile["numbering"]:
        print(f"nomor {row['source'][5:30]} {row['awalan']}: {row['first']}-{row['last']} "
              f"({row['count']}) berurutan={row['sequential']} ganda={row['duplicates']} ulang={row['restarts']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
