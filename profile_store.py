"""Akses profil format yang dibuat profile_builder, beserta pengaturan penomorannya.

Profil dibaca dari `profiles/<nama>/*.json`. Pengaturan yang bisa diubah pengguna
disimpan terpisah di `settings.json` agar tidak hilang saat profil dibangun ulang.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PROFILES_DIR = Path(os.environ.get("PROFILES_DIR", Path(__file__).resolve().parent / "profiles"))
SETTINGS_FILE = "settings.json"
CHAPTERS_FILE = "chapters.json"
STYLE_GUIDE_FILE = "style_guide.json"
STYLE_SAMPLES_FILE = "style_samples.json"
PROFILE_NAME = re.compile(r"[a-z0-9][a-z0-9-]{0,63}")
SECTION_NUMBER_PREFIX = re.compile(r"^\d+(?:\.\d+)*\.?\s*")
LANJUT_MODES = ("per_bagian", "berlanjut")
CAPTION_DEFAULTS = ("headline", "kosong")
DEFAULT_CAPTION = "kosong"
MAX_RULES = 50
MAX_STRUCTURES = 10
MAX_STRUCTURE_NAME_CHARS = 60
MAX_CHAPTERS = 20  # sama dengan batas bab per dokumen di form
MAX_CHAPTER_CHARS = 200
MAX_KEYWORD_CHARS = 200
MAX_PREFIX_CHARS = 20
MAX_START_NUMBER = 9999
EXAMPLE_CHARS = 90


class ProfileError(ValueError):
    """Permintaan profil tidak valid; pesannya aman ditampilkan ke UI."""


class ProfileNotFound(LookupError):
    """Profil dengan nama tersebut tidak ada."""


def _profile_path(name: str) -> Path:
    if not PROFILE_NAME.fullmatch(name or ""):
        raise ProfileError("Nama profil tidak valid.")
    path = PROFILES_DIR / name
    if not (path / "meta.json").is_file():
        raise ProfileNotFound(f"Profil '{name}' tidak ditemukan.")
    return path


def _read(path: Path, key: str) -> Any:
    return json.loads((path / f"{key}.json").read_text(encoding="utf-8"))


def list_profiles() -> list[dict[str, Any]]:
    if not PROFILES_DIR.is_dir():
        return []
    profiles = []
    for path in sorted(PROFILES_DIR.iterdir()):
        if not (path.is_dir() and PROFILE_NAME.fullmatch(path.name) and (path / "meta.json").is_file()):
            continue
        meta = _read(path, "meta")
        profiles.append({
            "name": path.name,
            "title": meta.get("title") or path.name,
            "sources": [source["file"] for source in meta.get("sources", [])],
            "layouts": len(_read(path, "layouts")),
            "created_at": meta.get("created_at", ""),
        })
    return profiles


def _section_keyword(section: str) -> str:
    """Kata kunci bagian dari judul terdalam, tanpa nomor bab ("2.1 Pokok …" jadi "Pokok …")."""
    leaf = section.split(" > ")[-1].strip()
    return SECTION_NUMBER_PREFIX.sub("", leaf).strip() or leaf


def numbering_evidence(captions: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Ringkasan penomoran yang ditemukan di referensi, dikelompokkan per bagian dan awalan."""
    groups: dict[tuple[str, str], dict[str, Any]] = {}
    for caption in captions:
        numbering = caption.get("penomoran")
        if not numbering:
            continue
        key = (_section_keyword(caption["section"]), numbering["awalan"])
        group = groups.setdefault(key, {
            "bagian": key[0], "awalan": key[1], "jumlah": 0,
            "pertama": numbering["nomor"], "terakhir": numbering["nomor"],
            "contoh": caption["text"][:EXAMPLE_CHARS],
        })
        group["jumlah"] += 1
        group["pertama"] = min(group["pertama"], numbering["nomor"])
        group["terakhir"] = max(group["terakhir"], numbering["nomor"])
    return list(groups.values())


def _prefixes_continuing_across_sections(captions: list[dict[str, Any]]) -> set[str]:
    sequences: dict[tuple[str, str], list[tuple[str, int]]] = defaultdict(list)
    for caption in captions:
        numbering = caption.get("penomoran")
        if numbering:
            sequences[(caption["source_file"], numbering["awalan"])].append(
                (caption["section"], numbering["nomor"])
            )
    continuing = set()
    for (_source, prefix), sequence in sequences.items():
        numbers = [number for _section, number in sequence]
        spans_sections = len({section for section, _number in sequence}) > 1
        if spans_sections and all(b >= a for a, b in zip(numbers, numbers[1:])):
            continuing.add(prefix)
    return continuing


def default_settings(captions: list[dict[str, Any]]) -> dict[str, Any]:
    """Aturan penomoran awal yang diturunkan dari apa yang dipakai referensi."""
    continuing = _prefixes_continuing_across_sections(captions)
    return {
        "penomoran": [
            {
                "bagian_cocok": group["bagian"],
                "awalan": group["awalan"],
                "mulai": group["pertama"],
                "lanjut": "berlanjut" if group["awalan"] in continuing else "per_bagian",
                "aktif": True,
            }
            for group in numbering_evidence(captions)
        ],
        "caption_default": DEFAULT_CAPTION,
    }


def _validate_rule(rule: Any, index: int) -> dict[str, Any]:
    label = f"Aturan penomoran #{index}"
    if not isinstance(rule, dict):
        raise ProfileError(f"{label} tidak valid.")
    keyword = rule.get("bagian_cocok")
    prefix = rule.get("awalan")
    start = rule.get("mulai")
    if not isinstance(keyword, str) or not keyword.strip() or len(keyword.strip()) > MAX_KEYWORD_CHARS:
        raise ProfileError(f"{label}: kata kunci bagian wajib diisi (maksimal {MAX_KEYWORD_CHARS} karakter).")
    if not isinstance(prefix, str) or not prefix.strip() or len(prefix.strip()) > MAX_PREFIX_CHARS:
        raise ProfileError(f"{label}: awalan wajib diisi (maksimal {MAX_PREFIX_CHARS} karakter).")
    if isinstance(start, bool) or not isinstance(start, int) or not 0 <= start <= MAX_START_NUMBER:
        raise ProfileError(f"{label}: nomor awal harus bilangan bulat 0-{MAX_START_NUMBER}.")
    if rule.get("lanjut") not in LANJUT_MODES:
        raise ProfileError(f"{label}: mode penomoran harus salah satu dari {', '.join(LANJUT_MODES)}.")
    if not isinstance(rule.get("aktif"), bool):
        raise ProfileError(f"{label}: status aktif harus benar atau salah.")
    return {
        "bagian_cocok": keyword.strip(), "awalan": prefix.strip(), "mulai": start,
        "lanjut": rule["lanjut"], "aktif": rule["aktif"],
    }


def validate_settings(data: Any) -> dict[str, Any]:
    rules = data.get("penomoran") if isinstance(data, dict) else None
    if not isinstance(rules, list):
        raise ProfileError("Pengaturan harus berisi daftar 'penomoran'.")
    if len(rules) > MAX_RULES:
        raise ProfileError(f"Maksimal {MAX_RULES} aturan penomoran.")
    caption_default = data.get("caption_default", DEFAULT_CAPTION)
    if caption_default not in CAPTION_DEFAULTS:
        raise ProfileError(f"Caption bawaan harus salah satu dari {', '.join(CAPTION_DEFAULTS)}.")
    return {
        "penomoran": [_validate_rule(rule, i) for i, rule in enumerate(rules, start=1)],
        "caption_default": caption_default,
    }


def validate_chapters(data: Any) -> dict[str, Any]:
    """Struktur bab tersimpan per profil: {'bawaan': nama, 'daftar': [{'nama', 'bab': [judul, ...]}]}."""
    if not isinstance(data, dict) or not isinstance(data.get("daftar"), list):
        raise ProfileError("Struktur bab harus berisi daftar 'daftar'.")
    if len(data["daftar"]) > MAX_STRUCTURES:
        raise ProfileError(f"Maksimal {MAX_STRUCTURES} struktur bab per profil.")
    structures, names = [], set()
    for index, item in enumerate(data["daftar"], start=1):
        label = f"Struktur bab #{index}"
        name = item.get("nama") if isinstance(item, dict) else None
        if not isinstance(name, str) or not name.strip() or len(name.strip()) > MAX_STRUCTURE_NAME_CHARS:
            raise ProfileError(f"{label}: nama wajib diisi (maksimal {MAX_STRUCTURE_NAME_CHARS} karakter).")
        name = " ".join(name.split())
        if name.lower() in names:
            raise ProfileError(f"{label}: nama '{name}' sudah dipakai struktur lain.")
        names.add(name.lower())
        chapters = item.get("bab")
        if not isinstance(chapters, list) or not all(isinstance(c, str) for c in chapters):
            raise ProfileError(f"{label} ('{name}'): 'bab' harus berupa daftar judul.")
        chapters = [" ".join(c.split()) for c in chapters if c.strip()]
        if not chapters:
            raise ProfileError(f"{label} ('{name}'): isi minimal satu judul bab.")
        if len(chapters) > MAX_CHAPTERS:
            raise ProfileError(f"{label} ('{name}'): maksimal {MAX_CHAPTERS} bab.")
        if any(len(c) > MAX_CHAPTER_CHARS for c in chapters):
            raise ProfileError(f"{label} ('{name}'): judul bab maksimal {MAX_CHAPTER_CHARS} karakter.")
        structures.append({"nama": name, "bab": chapters})
    default = data.get("bawaan", "")
    if not isinstance(default, str):
        raise ProfileError("Struktur bab bawaan harus berupa nama.")
    default = " ".join(default.split())
    if structures:
        by_name = {item["nama"].lower(): item["nama"] for item in structures}
        if default and default.lower() not in by_name:
            raise ProfileError(f"Struktur bab bawaan '{default}' tidak ada di daftar.")
        default = by_name.get(default.lower(), structures[0]["nama"])
    else:
        default = ""
    return {"bawaan": default, "daftar": structures}


def load_chapters(name: str) -> dict[str, Any]:
    path = _profile_path(name)
    stored = _read_optional(path, CHAPTERS_FILE)
    if stored is None:
        return {"bawaan": "", "daftar": []}
    try:
        return validate_chapters(stored)
    except ProfileError as exc:
        raise ProfileError(f"Berkas struktur bab profil rusak: {exc}") from exc


def save_chapters(name: str, data: Any) -> dict[str, Any]:
    path = _profile_path(name)
    chapters = validate_chapters(data)
    _write_json(path / CHAPTERS_FILE, chapters)
    return chapters


def _write_json(path: Path, data: Any) -> None:
    temporary = path.with_name(f"{path.name}.tmp")
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)


def _read_optional(path: Path, filename: str) -> Any:
    target = path / filename
    return json.loads(target.read_text(encoding="utf-8")) if target.is_file() else None


def _samples_sha256(path: Path) -> str | None:
    target = path / STYLE_SAMPLES_FILE
    return hashlib.sha256(target.read_bytes()).hexdigest() if target.is_file() else None


def _style_summary(path: Path) -> dict[str, Any]:
    samples = _read_optional(path, STYLE_SAMPLES_FILE)
    saved = _read_optional(path, STYLE_GUIDE_FILE)
    return {
        "statistik": _read_optional(path, "style_stats.json"),
        "sampel": (
            {"blok": len(samples), "karakter": sum(len(s["text"]) for s in samples)}
            if samples is not None else None
        ),
        "panduan": saved,
        # Panduan dibuat dari sampel yang sudah berbeda dengan sampel profil saat ini.
        "usang": bool(saved) and saved.get("sampel_sha256") != _samples_sha256(path),
    }


def style_inputs(name: str) -> tuple[list[dict[str, str]], dict[str, Any]]:
    """Sampel teks dan statistik gaya yang akan dikirim ke model untuk membuat panduan."""
    path = _profile_path(name)
    samples = _read_optional(path, STYLE_SAMPLES_FILE)
    statistics = _read_optional(path, "style_stats.json")
    if not samples or statistics is None:
        raise ProfileError("Profil belum memiliki sampel gaya. Bangun ulang profil dengan profile_builder.py.")
    return samples, statistics


def save_style_guide(name: str, guide: dict[str, Any], model_id: str, dropped_examples: int) -> dict[str, Any]:
    path = _profile_path(name)
    saved = {
        "panduan": guide,
        "model_id": model_id,
        "dibuat": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "contoh_dibuang": dropped_examples,
        "sampel_sha256": _samples_sha256(path),
    }
    _write_json(path / STYLE_GUIDE_FILE, saved)
    return saved


def load_profile(name: str) -> dict[str, Any]:
    path = _profile_path(name)
    captions = _read(path, "captions")
    settings_path = path / SETTINGS_FILE
    saved = settings_path.is_file()
    if saved:
        try:
            settings = validate_settings(json.loads(settings_path.read_text(encoding="utf-8")))
        except (json.JSONDecodeError, ProfileError) as exc:
            raise ProfileError(f"Berkas pengaturan profil rusak: {exc}") from exc
    else:
        settings = default_settings(captions)
    return {
        "name": name,
        "meta": _read(path, "meta"),
        "layouts": _read(path, "layouts"),
        "row_patterns": _read(path, "row_patterns"),
        "sections": _read(path, "sections"),
        "numbering": _read(path, "numbering"),
        # Profil lama dibuat tanpa PDF, jadi page.json bisa tidak ada.
        "page": _read_optional(path, "page.json"),
        "gaya": _style_summary(path),
        "caption_counts": dict(Counter(c["type"] for c in captions)),
        "evidence": numbering_evidence(captions),
        "settings": settings,
        "settings_saved": saved,
        "struktur_bab": load_chapters(name),
    }


def save_settings(name: str, data: Any) -> dict[str, Any]:
    path = _profile_path(name)
    settings = validate_settings(data)
    _write_json(path / SETTINGS_FILE, settings)
    return settings


def reset_settings(name: str) -> dict[str, Any]:
    path = _profile_path(name)
    (path / SETTINGS_FILE).unlink(missing_ok=True)
    return default_settings(_read(path, "captions"))
