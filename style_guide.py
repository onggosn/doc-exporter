"""Statistik gaya penulisan, sampel teks, dan validasi panduan gaya dari HTML referensi.

Bagian terukur (statistik) dihitung dengan aturan tetap. Model bahasa hanya merangkum
gayanya menjadi panduan, dan panduan itu divalidasi: kutipan contoh harus ada persis
di sampel, dan isinya tidak boleh memuat fakta dari dokumen referensi.
"""
from __future__ import annotations

import json
import re
import statistics
from collections import Counter
from typing import Any

from bs4 import BeautifulSoup, Tag


SAMPLE_MAX_CHARS = 24_000
BLOCK_MIN_CHARS = 40
BLOCK_MAX_CHARS = 600
MAX_BLOCKS_PER_GROUP = 25
ITALIC_TOP = 10
ITALIC_MAX_CHARS = 40
MAX_GUIDE_FIELD_CHARS = 1200
MAX_EXAMPLES = 6
MAX_PROHIBITIONS = 10
EXAMPLE_CONTEXTS = ("paragraf", "poin", "caption")
CAPTION_TYPES_FOR_SAMPLES = ("headline", "label_topik", "label_bernomor")

# Singkatan tidak baku yang lazim di tulisan kerja berbahasa Indonesia; yang dihitung
# adalah seberapa sering dipakai dibanding bentuk lengkapnya di referensi.
ABBREVIATIONS = {
    "yg": "yang", "dg": "dengan", "dgn": "dengan", "utk": "untuk", "krn": "karena",
    "shg": "sehingga", "scr": "secara", "tdk": "tidak", "sbg": "sebagai", "dlm": "dalam",
    "thd": "terhadap", "tsb": "tersebut", "jg": "juga", "dpt": "dapat", "spt": "seperti",
    "hrs": "harus", "mjd": "menjadi",
}
TECHNICAL_TERMS = ("bps", "ppt", "yoy", "mtm", "ytd", "qtq")
GUIDE_TEXT_FIELDS = (
    "nada", "sudut_pandang", "struktur_paragraf", "panjang_kalimat", "gaya_poin",
    "istilah_dan_singkatan", "format_angka", "istilah_asing",
)
GUIDE_CAPTION_FIELDS = ("headline", "label")

_SENTENCE_END = re.compile(r"(?<=[.!?])\s+(?=[A-Z0-9“\"(])")
_SOURCE_NOTE = re.compile(r"^(source|sumber)\s*:", re.IGNORECASE)
_DECIMAL_COMMA = re.compile(r"(?<![\d,.])\d+,\d+(?![\d,])")
_DECIMAL_POINT = re.compile(r"(?<![\d.,])\d{1,3}\.\d{1,2}(?![\d.,])")
_PERCENT_ATTACHED = re.compile(r"\d%")
_PERCENT_SPACED = re.compile(r"\d\s+%")
_RUPIAH_ATTACHED = re.compile(r"Rp\d")
_RUPIAH_SPACED = re.compile(r"Rp\s+\d")


class StyleGuideError(ValueError):
    """Panduan gaya dari model tidak memenuhi bentuk yang diminta."""


def _squash(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def _own_text(item: Tag) -> str:
    """Teks sebuah <li> tanpa teks daftar bersarang di dalamnya."""
    return _squash(" ".join(s for s in item.find_all(string=True) if s.find_parent("li") is item))


def extract_blocks(soup: BeautifulSoup, source: str) -> list[dict[str, str]]:
    """Paragraf dan poin berisi teks dokumen, tanpa judul bagian, caption, dan catatan sumber."""
    blocks: list[dict[str, str]] = []
    seen: set[str] = set()
    for tag in soup.find_all(["p", "li"]):
        if tag.find_parent("figcaption") is not None:
            continue
        if tag.find_parent("div", class_=re.compile(r"^title-level-\d+$")) is not None:
            continue
        if tag.name == "p" and tag.find_parent("li") is not None:
            continue
        text = _own_text(tag) if tag.name == "li" else _squash(tag.get_text(" "))
        if len(text) < BLOCK_MIN_CHARS or _SOURCE_NOTE.match(text) or text in seen:
            continue
        seen.add(text)
        blocks.append({"source": source, "kind": "poin" if tag.name == "li" else "paragraf", "text": text})
    return blocks


def italic_terms(soup: BeautifulSoup) -> list[str]:
    terms = []
    for tag in soup.find_all(["i", "em"]):
        text = _squash(tag.get_text(" "))
        if any(c.isalpha() for c in text) and len(text) <= ITALIC_MAX_CHARS:
            terms.append(text)
    return terms


def _sentence_lengths(texts: list[str]) -> dict[str, Any]:
    lengths = [
        len(sentence.split())
        for text in texts for sentence in _SENTENCE_END.split(text) if sentence.strip()
    ]
    if not lengths:
        return {"jumlah": 0, "rata_kata": None, "median_kata": None}
    return {
        "jumlah": len(lengths),
        "rata_kata": round(statistics.mean(lengths), 1),
        "median_kata": statistics.median(lengths),
    }


def _count(pattern: re.Pattern[str], text: str) -> int:
    return len(pattern.findall(text))


def compute_statistics(
    blocks: list[dict[str, str]], captions: list[dict[str, Any]], italics: list[str]
) -> dict[str, Any]:
    by_kind = {kind: [b["text"] for b in blocks if b["kind"] == kind] for kind in ("paragraf", "poin")}
    text = "\n".join(b["text"] for b in blocks)

    def word_count(word: str) -> int:
        return len(re.findall(rf"\b{word}\b", text, flags=re.IGNORECASE))

    abbreviations = []
    for full in dict.fromkeys(ABBREVIATIONS.values()):
        short = {a: word_count(a) for a, f in ABBREVIATIONS.items() if f == full}
        entry = {"lengkap": full, "singkatan": short,
                 "jumlah_singkat": sum(short.values()), "jumlah_lengkap": word_count(full)}
        if entry["jumlah_singkat"] or entry["jumlah_lengkap"]:
            abbreviations.append(entry)
    abbreviations.sort(key=lambda e: -e["jumlah_singkat"])

    caption_types = Counter(c["type"] for c in captions if c.get("text"))
    return {
        "blok": {**{kind: len(texts) for kind, texts in by_kind.items()}, "caption": sum(caption_types.values())},
        "kalimat": {kind: _sentence_lengths(texts) for kind, texts in by_kind.items()},
        "singkatan": abbreviations,
        "angka": {
            "desimal_koma": _count(_DECIMAL_COMMA, text),
            "desimal_titik": _count(_DECIMAL_POINT, text),
            "persen_rapat": _count(_PERCENT_ATTACHED, text),
            "persen_spasi": _count(_PERCENT_SPACED, text),
            "rupiah_rapat": _count(_RUPIAH_ATTACHED, text),
            "rupiah_spasi": _count(_RUPIAH_SPACED, text),
        },
        "istilah_teknis": {term: n for term in TECHNICAL_TERMS if (n := word_count(term))},
        "miring": {"jumlah": len(italics), "teratas": [t for t, _ in Counter(italics).most_common(ITALIC_TOP)]},
    }


def select_samples(
    blocks: list[dict[str, str]], captions: list[dict[str, Any]], max_chars: int = SAMPLE_MAX_CHARS
) -> list[dict[str, str]]:
    """Pilih sampel merata di seluruh dokumen; hasil sama untuk masukan yang sama."""
    groups: dict[tuple[str, str], list[dict[str, str]]] = {}
    for block in blocks:
        groups.setdefault((block["source"], block["kind"]), []).append(block)
    for caption in captions:
        if caption.get("text") and caption["type"] in CAPTION_TYPES_FOR_SAMPLES:
            groups.setdefault((caption["source_file"], "caption"), []).append(
                {"source": caption["source_file"], "kind": "caption", "text": caption["text"]}
            )
    if not groups:
        return []
    budget = max_chars // len(groups)
    samples: list[dict[str, str]] = []
    for key in sorted(groups):
        items = groups[key]
        chosen, used = [], 0
        average = max(1, sum(min(len(i["text"]), BLOCK_MAX_CHARS) for i in items) // len(items))
        take = max(1, min(len(items), budget // average, MAX_BLOCKS_PER_GROUP))
        for index in sorted({int(i * len(items) / take) for i in range(take)}):
            text = items[index]["text"][:BLOCK_MAX_CHARS]
            if used + len(text) > budget:
                break
            used += len(text)
            chosen.append({**items[index], "text": text})
        samples += chosen
    return samples


def build_prompt(samples: list[dict[str, str]], statistics_: dict[str, Any]) -> str:
    listing = "\n".join(f"[{s['kind']}] {s['text']}" for s in samples)
    return (
        "Anda menganalisis GAYA PENULISAN sebuah jenis dokumen kerja berbahasa Indonesia dari sampel teks "
        "referensi di bawah. Tugas Anda hanya merangkum gaya (nada, struktur, pilihan kata, format angka, "
        "cara menulis caption) agar penulis lain bisa meniru gayanya pada materi yang berbeda.\n\n"
        "ATURAN KETAT:\n"
        "- Jangan menulis fakta, angka, nama, atau kesimpulan dari sampel ke dalam panduan; panduan hanya "
        "tentang cara menulis, bukan isi.\n"
        "- Statistik terukur di bawah adalah kebenaran. Jangan bertentangan dengannya. Bila pemakaian "
        "tidak konsisten (mis. desimal koma dan titik sama-sama muncul), katakan itu apa adanya.\n"
        f"- Tiap bidang teks maksimal 3 kalimat dan {MAX_GUIDE_FIELD_CHARS} karakter.\n"
        f"- 'contoh' berisi paling banyak {MAX_EXAMPLES} kutipan pendek yang SALIN PERSIS dari sampel "
        "(kata per kata), mewakili gaya; jangan mengubah satu huruf pun. 'konteks' salah satu dari: "
        + ", ".join(EXAMPLE_CONTEXTS) + ".\n"
        f"- 'larangan' berisi paling banyak {MAX_PROHIBITIONS} hal yang dihindari gaya ini.\n"
        "- Kembalikan HANYA satu objek JSON tanpa markdown, dengan bentuk:\n"
        '{"nada": "", "sudut_pandang": "", "struktur_paragraf": "", "panjang_kalimat": "", "gaya_poin": "", '
        '"istilah_dan_singkatan": "", "format_angka": "", "istilah_asing": "", '
        '"gaya_caption": {"headline": "", "label": ""}, "larangan": [""], '
        '"contoh": [{"konteks": "", "teks": ""}]}\n\n'
        f"Statistik terukur:\n{json.dumps(statistics_, ensure_ascii=False)}\n\n"
        f"Sampel teks referensi:\n{listing}"
    )


def _text_field(guide: dict[str, Any], key: str, label: str) -> str:
    value = guide.get(key)
    if not isinstance(value, str) or not value.strip():
        raise StyleGuideError(f"bidang '{label}' kosong atau bukan teks")
    if len(value.strip()) > MAX_GUIDE_FIELD_CHARS:
        raise StyleGuideError(f"bidang '{label}' lebih dari {MAX_GUIDE_FIELD_CHARS} karakter")
    return value.strip()


def validate_guide(guide: dict[str, Any], samples: list[dict[str, str]]) -> tuple[dict[str, Any], int]:
    """Kembalikan (panduan bersih, jumlah contoh yang dibuang karena bukan kutipan persis)."""
    cleaned: dict[str, Any] = {key: _text_field(guide, key, key) for key in GUIDE_TEXT_FIELDS}

    captions = guide.get("gaya_caption")
    if not isinstance(captions, dict):
        raise StyleGuideError("bidang 'gaya_caption' harus berupa objek")
    cleaned["gaya_caption"] = {
        key: _text_field(captions, key, f"gaya_caption.{key}") for key in GUIDE_CAPTION_FIELDS
    }

    prohibitions = guide.get("larangan")
    if not isinstance(prohibitions, list) or not all(isinstance(p, str) for p in prohibitions):
        raise StyleGuideError("bidang 'larangan' harus berupa daftar teks")
    cleaned["larangan"] = [p.strip() for p in prohibitions if p.strip()][:MAX_PROHIBITIONS]

    examples = guide.get("contoh")
    if not isinstance(examples, list):
        raise StyleGuideError("bidang 'contoh' harus berupa daftar")
    corpus = _squash("\n".join(s["text"] for s in samples))
    kept, dropped = [], 0
    for example in examples:
        text = _squash(example.get("teks", "")) if isinstance(example, dict) else ""
        context = example.get("konteks") if isinstance(example, dict) else None
        if text and context in EXAMPLE_CONTEXTS and text in corpus and len(kept) < MAX_EXAMPLES:
            kept.append({"konteks": context, "teks": text})
        else:
            dropped += 1
    cleaned["contoh"] = kept
    return cleaned, dropped
