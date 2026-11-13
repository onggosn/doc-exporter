"""Tinjauan format, urutan butir, dan penomoran pada halaman yang memuat butir atau daftar bernomor.

Hanya untuk blok daftar (`numbered`, `bullets`, dan label tebal di antaranya) yang dibawa apa adanya dari slide
(`doc_composer.verbatim_blocks`). Teks tidak diubah: temuan dipasang sebagai komentar Word pada butir yang bersangkutan
(`comments` pada entri `numbered`, `item_comments` pada blok `bullets`, `comments` pada label).

Pemeriksaan (semuanya deterministik):
- isi: urutan dan isi baris di dokumen sama dengan baris slide.
- urutan: nomor berurutan pada kedalaman yang sama (1, 2, 3 / a, b, c / i, ii, iii); nomor ganda, melompat, mundur,
  atau mulai lagi dari awal setelah disela butir tanpa nomor.
- format: gaya penanda berubah pada kedalaman yang sama, tingkat butir melompat lebih dari satu, atau kedalaman slide
  melebihi batas dokumen sehingga diratakan.

Penanda pertama yang bukan "1."/"a." dan daftar yang dimulai dari tingkat dalam tidak ditandai: halaman lanjutan memang
bisa meneruskan daftar dari halaman sebelumnya.
"""
from __future__ import annotations

import re
from typing import Any

MAX_DEPTH = 2  # sama dengan batas kedalaman di doc_composer.verbatim_blocks dan docx_builder
_ROMAN = re.compile(r"^[ivxlcdm]+$", re.I)
_ROMAN_VALUES = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}
_TAG = re.compile(r"</?[bi]>")


def _norm(text: str) -> str:
    return " ".join(re.sub(r"[<>]", "", text).split())


def _roman(core: str) -> int | None:
    if not _ROMAN.match(core):
        return None
    total, previous = 0, 0
    for char in reversed(core.lower()):
        value = _ROMAN_VALUES[char]
        total += -value if value < previous else value
        previous = max(previous, value)
    return total


def _alpha(core: str) -> int | None:
    return ord(core.lower()) - 96 if len(core) == 1 and core.isalpha() else None


def parse_marker(marker: str, previous_style: str | None = None) -> dict[str, Any] | None:
    """Pecah penanda ("1.", "a)", "(i)") menjadi {"gaya", "nilai", "bentuk"}; None bila tak dikenali.
    Huruf tunggal yang bisa angka romawi atau abjad ("i", "c") mengikuti gaya penanda sebelumnya."""
    match = re.fullmatch(r"(\(?)([A-Za-z]+|\d+)([.)])", marker.strip())
    if not match:
        return None
    opener, core, closer = match.groups()
    shape = f"{opener}{closer}"
    if core.isdigit():
        return {"gaya": "angka", "nilai": int(core), "bentuk": shape}
    case = "besar" if core.isupper() else "kecil"
    roman, alpha = _roman(core), _alpha(core)
    if roman is not None and alpha is not None:  # ambigu: "i", "v", "x", "l", "c", "d", "m"
        # Tanpa konteks, "i", "v", "x" dibaca romawi; "l", "c", "d", "m" dibaca abjad (romawi sebesar itu jarang di slide).
        want_alpha = previous_style == f"abjad-{case}" or (previous_style is None and core.lower() in "lcdm")
        style, value = ("abjad", alpha) if want_alpha else ("romawi", roman)
        other = ("romawi", roman) if want_alpha else ("abjad", alpha)
        return {"gaya": f"{style}-{case}", "nilai": value, "bentuk": shape,
                "alt": {"gaya": f"{other[0]}-{case}", "nilai": other[1]}}
    elif roman is not None:
        style, value = "romawi", roman
    elif alpha is not None:
        style, value = "abjad", alpha
    else:
        return None
    return {"gaya": f"{style}-{case}", "nilai": value, "bentuk": shape}


def _rows(blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for b, block in enumerate(blocks):
        kind = block.get("type")
        if kind == "numbered":
            for i, entry in enumerate(block["items"]):
                rows.append({"jenis": "nomor", "b": b, "i": i, "penanda": entry["marker"],
                             "teks": entry["text"], "depth": int(entry.get("depth", 0)),
                             "keluar": _norm(f"{entry['marker']} {entry['text']}")})
        elif kind == "bullets":
            for i, text in enumerate(block["items"]):
                depth = min((len(text) - len(text.lstrip(" "))) // 2, MAX_DEPTH)
                rows.append({"jenis": "butir", "b": b, "i": i, "penanda": "", "teks": text.strip(),
                             "depth": depth, "keluar": _norm(text)})
        elif kind == "paragraph":  # label tebal
            rows.append({"jenis": "label", "b": b, "i": None, "penanda": "", "teks": block["text"],
                         "depth": 0, "keluar": _norm(_TAG.sub("", block["text"]))})
    return rows


def _check_content(source: list[str], rows: list[dict[str, Any]]) -> list[tuple[int, str, str]]:
    lines = [(line, _norm(line)) for line in source if _norm(line)]
    expected = [text for _, text in lines]
    produced = [row["keluar"] for row in rows]
    if expected == produced:
        return []
    for index in range(min(len(expected), len(produced))):
        if expected[index] != produced[index]:
            return [(index, "isi", f"Baris ke-{index + 1} di dokumen berbeda dari slide. Slide: \"{expected[index][:100]}\". "
                                   f"Dokumen: \"{produced[index][:100]}\".")]
    shorter = min(len(expected), len(produced))
    return [(max(shorter - 1, 0), "isi",
             f"Jumlah baris berbeda: slide {len(expected)} baris, dokumen {len(produced)} baris.")]


def _check_depth_limit(source: list[str], rows: list[dict[str, Any]]) -> list[tuple[int, str, str]]:
    lines = [line for line in source if _norm(line)]
    if len(lines) != len(rows):
        return []
    found = []
    for index, line in enumerate(lines):
        level = (len(line) - len(line.lstrip(" "))) // 2
        if level > MAX_DEPTH:
            found.append((index, "format",
                          f"Di slide butir ini bertingkat {level + 1}; dokumen hanya mendukung {MAX_DEPTH + 1} tingkat, jadi diratakan."))
    return found


def _check_structure(rows: list[dict[str, Any]]) -> list[tuple[int, str, str]]:
    found: list[tuple[int, str, str]] = []
    last: dict[int, dict[str, Any]] = {}
    interrupted: dict[int, bool] = {}
    previous_depth: int | None = None
    for index, row in enumerate(rows):
        depth = row["depth"]
        if previous_depth is not None and depth > previous_depth + 1:
            found.append((index, "format", f"Tingkat butir melompat dari {previous_depth + 1} ke {depth + 1}."))
        previous_depth = depth
        if row["jenis"] == "label":
            last.clear()
            interrupted.clear()
            continue
        for deeper in [d for d in last if d > depth]:
            del last[deeper], interrupted[deeper]
        if row["jenis"] == "butir":
            if depth in last:
                interrupted[depth] = True
            continue
        before = last.get(depth)
        marker = parse_marker(row["penanda"], before["gaya"] if before else None)
        if marker is not None and before and marker["nilai"] != before["nilai"] + 1:
            # Huruf tunggal ambigu ("i", "c"): bila bukan lanjutan dari penanda sebelumnya, baca sebagai romawi
            # supaya "i." lalu "ii." tidak dianggap pindah gaya.
            alternative = parse_marker(row["penanda"], None)
            if alternative is not None and alternative["nilai"] != marker["nilai"] or (
                    alternative and alternative["gaya"] != marker["gaya"] and alternative["nilai"] == before["nilai"] + 1):
                marker = alternative
        if marker is not None and before and before.get("alt") and marker["gaya"] == before["alt"]["gaya"] \
                and marker["nilai"] == before["alt"]["nilai"] + 1 and marker["bentuk"] == before["bentuk"]:
            before = {**before, **before["alt"]}  # "i." lalu "j." = abjad, bukan romawi lalu abjad
        if marker is None:
            found.append((index, "format", f"Penanda \"{row['penanda']}\" tidak dikenali sebagai nomor atau huruf."))
            continue
        if not row["teks"].strip():
            found.append((index, "format", f"Butir bernomor \"{row['penanda']}\" tidak berisi teks."))
        if before is not None:  # penanda pertama boleh bukan "1." (daftar berlanjut dari slide sebelumnya)
            if marker["gaya"] != before["gaya"] or marker["bentuk"] != before["bentuk"]:
                found.append((index, "format", f"Gaya penomoran berubah dari \"{before['penanda']}\" menjadi \"{row['penanda']}\" pada tingkat yang sama."))
            elif marker["nilai"] == before["nilai"] + 1:
                pass
            elif marker["nilai"] == 1 and interrupted.get(depth):
                found.append((index, "urutan", f"Penomoran mulai lagi dari \"{row['penanda']}\" setelah \"{before['penanda']}\" karena disela butir tanpa nomor; "
                                               "cocokkan dengan slide apakah seharusnya berlanjut."))
            elif marker["nilai"] == 1 and before["nilai"] > 1:
                pass  # daftar baru (kotak teks lain) yang wajar mulai dari awal
            elif marker["nilai"] == before["nilai"]:
                found.append((index, "urutan", f"Nomor ganda: \"{row['penanda']}\" muncul lagi setelah \"{before['penanda']}\"."))
            else:
                found.append((index, "urutan", f"Urutan nomor tidak berurutan: \"{row['penanda']}\" setelah \"{before['penanda']}\"."))
        last[depth] = {**marker, "penanda": row["penanda"]}
        interrupted[depth] = False
    return found


def review_blocks(source: list[str], blocks: list[dict[str, Any]]) -> list[dict[str, str]]:
    """Periksa blok daftar satu slide dan pasang komentar pada butirnya. Mengembalikan temuan
    [{"jenis", "pesan", "teks"}]; kosong (dan tanpa komentar) bila slide tidak memuat butir atau nomor."""
    if not any(block.get("type") in ("numbered", "bullets") for block in blocks):
        return []
    rows = _rows(blocks)
    if not rows:
        return []
    raw = [*_check_content(source, rows), *_check_depth_limit(source, rows), *_check_structure(rows)]
    findings = []
    for index, kind, message in raw:
        row = rows[min(index, len(rows) - 1)]
        text = f"[Cek format: {kind}] {message}"
        block = blocks[row["b"]]
        if row["jenis"] == "nomor":
            block["items"][row["i"]].setdefault("comments", []).append(text)
        elif row["jenis"] == "butir":
            block.setdefault("item_comments", {}).setdefault(row["i"], []).append(text)
        else:
            block.setdefault("comments", []).append(text)
        findings.append({"jenis": kind, "pesan": message, "teks": (row["penanda"] + " " + row["teks"]).strip()[:160]})
    return findings
