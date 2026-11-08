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

import style_guide
from pdf_profile import analyze_pdf


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


def build_page_profile(pdf_paths: list[Path], typography: list[dict[str, Any]]) -> dict[str, Any]:
    pdfs = [analyze_pdf(path) for path in pdf_paths]
    return {
        "pdf": pdfs,
        "html": typography,
        "skala_cetak": _print_scale(pdfs, typography),
        "catatan": "Skala cetak di bawah 1 berarti browser menyusutkan halaman saat mencetak; "
                   "ukuran font dokumen mengikuti HTML, bukan PDF.",
    }


def build_profile(paths: list[Path], title: str = "", pdf_paths: list[Path] | None = None) -> dict[str, Any]:
    items: list[dict[str, Any]] = []
    sources = []
    typography = []
    blocks: list[dict[str, str]] = []
    italics: list[str] = []
    for path in paths:
        raw = path.read_text(encoding="utf-8", errors="replace")
        soup = _load_html(raw)
        blocks += style_guide.extract_blocks(soup, path.name)
        italics += style_guide.italic_terms(soup)
        typography.append(html_typography(raw, path.name))
        found = analyze_html(raw, path.name)
        sources.append({
            "file": path.name,
            "sha256": hashlib.sha256(raw.encode()).hexdigest(),
            "images_in_html": count_images(raw),
            "images_in_layouts": sum(item["images"] for item in found),
        })
        items.extend(found)

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
    if pdf_paths:
        profile["meta"]["pdf_sources"] = [
            {"file": path.name, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()} for path in pdf_paths
        ]
        profile["page"] = build_page_profile(pdf_paths, typography)
    return profile


def save_profile(profile: dict[str, Any], out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for key in PROFILE_FILES + tuple(k for k in OPTIONAL_PROFILE_FILES if k in profile):
        (out_dir / f"{key}.json").write_text(
            json.dumps(profile[key], ensure_ascii=False, indent=2), encoding="utf-8"
        )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Bangun profil format dari HTML referensi.")
    parser.add_argument("html", nargs="+", type=Path, help="Satu atau lebih berkas HTML referensi")
    parser.add_argument("--name", required=True, help="Jenis dokumen, juga nama folder di profiles/ (mis. dokumen-2d)")
    parser.add_argument("--title", default="", help="Nama tampilan jenis dokumen (mis. 'Dokumen 2d')")
    parser.add_argument("--pdf", nargs="*", type=Path, default=[], help="PDF referensi untuk properti halaman")
    parser.add_argument("--out", type=Path, default=Path("profiles"))
    args = parser.parse_args(argv)

    profile = build_profile(args.html, args.title or args.name, args.pdf)
    save_profile(profile, args.out / args.name)
    print(f"Profil disimpan di {args.out / args.name}")
    for pdf in profile.get("page", {}).get("pdf", []):
        margin = pdf["margin_pt"]
        print(f"halaman {pdf['file'][:34]}: {pdf['ukuran']['nama']} margin K{margin['kiri']}/R{margin['kanan']}"
              f"/A{margin['atas']}/B{margin['bawah']} watermark={(pdf['watermark'] or {}).get('teks')}")
    for scale in profile.get("page", {}).get("skala_cetak", []):
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
