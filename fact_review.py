"""Tinjauan fakta khusus ringkasan gambar (caption slide atau panel): membandingkan ringkasan dengan fakta slide asalnya.

Narasi tidak ditinjau di sini. Hasilnya hanya catatan: teks tidak pernah diubah; catatan dipasang sebagai komentar Word
pada paragraf ringkasan yang bersangkutan (lihat docx_builder). Dua lapis:

1. Aturan deterministik (`check_text`): angka, periode (tahun, bulan, triwulan), dan arah perubahan ("naik" padahal
   angkanya turun). Tanpa panggilan model. Nama dan singkatan tidak diperiksa: ringkasan gambar wajar menyebut nama
   yang hanya terbaca dari gambarnya.
2. Model (`BedrockDocumentGenerator.review_slide`): membandingkan ringkasan dengan fakta tertulis dan gambarnya.
"""
from __future__ import annotations

import os
import re
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Callable

import doc_composer
from bedrock_service import run_parallel

AUTHOR = "Peninjau Fakta"
INITIALS = "PF"
MIN_NUMBERS_ENV = "FACT_REVIEW_MIN_NUMBERS"
DEFAULT_MIN_NUMBERS = 4  # slide dengan angka sebanyak ini (atau lebih) di teks tertulisnya ikut ditinjau model

MODES = ("tidak", "aturan", "model")

_TAG = re.compile(r"</?[a-zA-Z][^>]*>")
_YEAR = re.compile(r"\b(?:19|20)\d{2}\b")
_PLAIN_YEAR = re.compile(r"(?:19|20)\d{2}")
_QUARTER = re.compile(r"\b(?:TW|Q|triwulan|kuartal)\s*-?\s*(IV|III|II|I|[1-4])\b", re.I)
_QUARTER_NUMBER = {"I": "1", "II": "2", "III": "3", "IV": "4"}
_FROM_TO = re.compile(
    r"\bdari\b[^\d.;]{0,30}(\d+(?:[.,]\d+)*)[^\d;]{0,50}?\b(?:menjadi|ke|hingga|sampai|menuju)\b[^\d;]{0,30}(\d+(?:[.,]\d+)*)",
    re.I,
)
_UP = re.compile(r"\b(?:naik|meningkat|menguat|tumbuh|melonjak|bertambah|menanjak|menaikkan|meningkatkan)\b", re.I)
_DOWN = re.compile(r"\b(?:turun|menurun|melemah|menyusut|terkoreksi|berkurang|merosot|anjlok|menurunkan)\b", re.I)
_MONTHS = (
    "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober",
    "November", "Desember",
)
_MONTH = re.compile(r"\b(" + "|".join(_MONTHS) + r")\b", re.I)
_SENTENCE = re.compile(r"(?<=[.!?])\s+")


def _plain(text: str) -> str:
    return _TAG.sub("", text or "")


def _float(token: str) -> float | None:
    try:
        return float(doc_composer.normalize_number(token))
    except ValueError:
        return None


def _supported_numbers(facts: str) -> set[str]:
    """Angka fakta, ditambah versi bulatnya: 4,85 di slide boleh ditulis 4,9 atau 5 di narasi."""
    base = doc_composer.extract_numbers(_YEAR.sub("", facts))
    expanded = set(base)
    for number in base:
        if "." not in number:
            continue
        value, decimals = Decimal(number), len(number.split(".")[1])
        for places in range(decimals):
            rounded = value.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
            expanded.add(doc_composer.normalize_number(f"{rounded:f}"))
    return expanded


def _check_numbers(text: str, facts: str) -> list[dict[str, str]]:
    supported = _supported_numbers(facts)
    missing: dict[str, str] = {}
    for match in doc_composer._NUMBER.finditer(text):
        token = match.group()
        if _PLAIN_YEAR.fullmatch(token):  # tahun diperiksa di _check_periods
            continue
        normalized = doc_composer.normalize_number(token)
        if len(normalized) == 1 and not text[match.end():match.end() + 1] == "%":
            continue  # bilangan satu digit ("3 faktor", "2 skenario") terlalu sering hitungan, bukan data
        if normalized not in supported:
            missing.setdefault(normalized, token)
    if not missing:
        return []
    shown = ", ".join(missing.values())
    return [{"jenis": "angka", "pesan": (
        f"Angka {shown} tidak ditemukan di teks slide. Mungkin hasil hitungan atau terbaca dari gambar; "
        "pastikan benar.")}]


def _check_periods(text: str, facts: str) -> list[dict[str, str]]:
    findings = []
    facts_years = set(_YEAR.findall(facts))
    years = sorted({y for y in _YEAR.findall(text) if y not in facts_years})
    if years:
        findings.append({"jenis": "periode", "pesan": f"Tahun {', '.join(years)} tidak ada di teks slide."})
    months = []
    for month in dict.fromkeys(m.capitalize() for m in _MONTH.findall(text)):
        if not re.search(rf"\b{month[:3]}", facts, re.I):
            months.append(month)
    if months:
        findings.append({"jenis": "periode", "pesan": f"Bulan {', '.join(months)} tidak ada di teks slide."})
    facts_quarters = {_QUARTER_NUMBER.get(q.upper(), q) for q in _QUARTER.findall(facts)}
    quarters = []
    for match in _QUARTER.finditer(text):
        number = _QUARTER_NUMBER.get(match.group(1).upper(), match.group(1))
        if number not in facts_quarters and match.group().strip() not in quarters:
            quarters.append(match.group().strip())
    if quarters:
        findings.append({"jenis": "periode", "pesan": f"Triwulan {', '.join(quarters)} tidak ada di teks slide."})
    return findings


def _check_direction(text: str) -> list[dict[str, str]]:
    findings = []
    for sentence in _SENTENCE.split(text):
        match = _FROM_TO.search(sentence)
        if not match:
            continue
        before, after = _float(match.group(1)), _float(match.group(2))
        if before is None or after is None or before == after:
            continue
        up, down = bool(_UP.search(sentence)), bool(_DOWN.search(sentence))
        if up == down:
            continue
        if (up and after < before) or (down and after > before):
            findings.append({"jenis": "arah", "pesan": (
                f"Kalimat menyebut {'kenaikan' if up else 'penurunan'}, tetapi angkanya berubah dari "
                f"{match.group(1)} menjadi {match.group(2)}: \"{sentence.strip()[:120]}\".")})
    return findings


def check_text(text: str, facts: str) -> list[dict[str, str]]:
    """Temuan aturan deterministik untuk satu paragraf: daftar {"jenis", "pesan"}."""
    plain = " ".join(_plain(text).split())
    if not plain:
        return []
    facts = " ".join((facts or "").split())
    return [*_check_numbers(plain, facts), *_check_periods(plain, facts),
            *_check_direction(plain)]


def collect_paragraphs(deck: dict[str, Any]) -> list[dict[str, Any]]:
    """Ringkasan gambar per slide: ringkasan tiap panel (slide yang dipecah) atau ringkasan slide utuh, masing-masing
    dengan gambarnya. Fakta = teks tertulis slide beserta deskripsi visualnya bila ada."""
    result = []
    for slide in deck["items"]:
        facts = doc_composer.all_facts(slide)
        panels = slide.get("panel_images") or []
        if len(panels) >= 2:
            sources = [(p.get("summary"), p.get("image_path")) for p in panels]
        else:
            sources = [(slide.get("summary"), slide.get("image_path"))]
        paragraphs = [{"jenis": "ringkasan", "teks": text, "fakta": facts, "gambar": image}
                      for text, image in sources if text]
        if paragraphs:
            result.append({"slide": slide, "paragraf": paragraphs})
    return result


def min_numbers() -> int:
    try:
        return max(0, int(os.environ.get(MIN_NUMBERS_ENV, DEFAULT_MIN_NUMBERS)))
    except ValueError:
        return DEFAULT_MIN_NUMBERS


def _comment(jenis: str, pesan: str) -> str:
    return f"[Cek fakta: {jenis}] {pesan}"


def review_deck(
    deck: dict[str, Any], generator: Any | None, progress: Callable[..., None] | None = None,
) -> dict[str, Any]:
    """Tinjau ringkasan gambar. `generator` None = hanya aturan deterministik. Mengembalikan:
    {"komentar": {(ref, teks): [str]}, "temuan": [{ref, jenis, teks, pesan}], "ditinjau_model": [ref],
     "gagal_model": [{ref, alasan}]}."""
    comments: dict[tuple[str, str], list[str]] = {}
    findings: list[dict[str, str]] = []

    def add(ref: str, text: str, jenis: str, pesan: str, source: str) -> None:
        comments.setdefault((ref, text), []).append(_comment(jenis, pesan))
        findings.append({"ref": ref, "sumber": source, "jenis": jenis, "teks": text[:160], "pesan": pesan})

    flagged: set[str] = set()
    for entry in collect_paragraphs(deck):
        ref = entry["slide"]["ref"]
        for paragraph in entry["paragraf"]:
            for finding in check_text(paragraph["teks"], paragraph["fakta"]):
                add(ref, paragraph["teks"], finding["jenis"], finding["pesan"], "aturan")
                flagged.add(ref)

    reviewed: list[str] = []
    failed: list[dict[str, str]] = []
    if generator is not None:
        threshold = min_numbers()
        selected = [
            entry for entry in collect_paragraphs(deck)
            if entry["slide"]["ref"] in flagged
            or sum(len(doc_composer.extract_numbers(p["teks"])) for p in entry["paragraf"]) >= threshold
        ]
        done = 0
        total = len(selected)
        if progress:
            progress("review", "running", "", 0, total)

        def run(entry: dict[str, Any]) -> list[dict[str, Any]]:
            return generator.review_slide(entry["slide"], entry["paragraf"])

        def advance(_index: int, _outcome: Any) -> None:
            nonlocal done
            done += 1
            if progress:
                progress("review", "running", "", done, total)

        outcomes = run_parallel(selected, run, generator.workers, stop_on_error=False, on_done=advance)
        for entry, (ok, value) in zip(selected, outcomes):
            ref = entry["slide"]["ref"]
            if not ok:
                failed.append({"ref": ref, "alasan": str(value)})
                continue
            reviewed.append(ref)
            for item in value:
                paragraph = entry["paragraf"][item["paragraf"] - 1]
                quote = f" Kutipan: \"{item['kutipan']}\"." if item.get("kutipan") else ""
                add(ref, paragraph["teks"], "model", f"{item['alasan']}{quote}", "model")
    return {"komentar": comments, "temuan": findings, "ditinjau_model": reviewed, "gagal_model": failed}


def summarize(result: dict[str, Any], mode: str) -> dict[str, Any]:
    """Ringkasan untuk laporan job (tanpa kunci tuple)."""
    return {
        "mode": mode,
        "jumlah_temuan": len(result["temuan"]),
        "temuan": result["temuan"],
        "slide_ditinjau_model": result["ditinjau_model"],
        "gagal_model": result["gagal_model"],
    }
