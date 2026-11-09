"""Kontrak OpenAPI untuk API konversi PPTX ke DOCX (disajikan di /api/docx/openapi.json)."""
from __future__ import annotations

from typing import Any

_ERROR = {"$ref": "#/components/schemas/Error"}
_SESSION = {"name": "session_id", "in": "path", "required": True, "schema": {"type": "string"},
            "description": "ID sesi dari /api/docx/start."}


def _json(schema: dict[str, Any], description: str) -> dict[str, Any]:
    return {"description": description, "content": {"application/json": {"schema": schema}}}


def _error(description: str) -> dict[str, Any]:
    return _json(_ERROR, description)


DOCX_OPENAPI: dict[str, Any] = {
    "openapi": "3.0.3",
    "info": {
        "title": "API konversi PPTX ke DOCX",
        "version": "1.0.0",
        "description": (
            "Alur: POST /api/docx/start (unggah PPTX) lalu ulangi GET /api/docx/status/{session_id} sampai "
            "status 'done', lalu GET /api/docx/download/{session_id}. Pembuatan berjalan di latar belakang dan "
            "bisa memakan beberapa menit; sesi dan berkasnya kedaluwarsa sesuai 'sesi_kedaluwarsa_detik' di "
            "/api/docx/options. Tanpa autentikasi: layanan ini ditujukan untuk jaringan lokal tepercaya."
        ),
    },
    "servers": [{"url": "/"}],
    "paths": {
        "/api/docx/options": {"get": {
            "summary": "Nilai bawaan dan batas yang dipakai /api/docx/start",
            "responses": {"200": _json({"$ref": "#/components/schemas/Options"}, "Opsi")},
        }},
        "/api/docx/start": {"post": {
            "summary": "Mulai konversi PPTX ke DOCX",
            "description": (
                "Mengembalikan segera setelah job berjalan di latar belakang. Bila 'chapters' tidak dikirim, "
                "struktur bab diambil dari profil: struktur bernama 'struktur', atau struktur bawaan profil, "
                "atau bawaan aplikasi. 'chapters' yang dikirim tetapi kosong ditolak."
            ),
            "requestBody": {"required": True, "content": {"multipart/form-data": {"schema": {
                "type": "object",
                "required": ["presentation", "profile"],
                "properties": {
                    "presentation": {"type": "string", "format": "binary", "description": "Berkas .pptx."},
                    "profile": {"type": "string", "description": "Nama profil jenis dokumen, mis. dokumen-2d."},
                    "chapters": {"type": "string", "description": "Judul bab, satu per baris (1 sampai 20)."},
                    "struktur": {"type": "string", "description": "Nama struktur bab tersimpan di profil."},
                    "mapping_mode": {"type": "string", "enum": ["batch", "per_slide"], "default": "batch"},
                    "split_panels": {"type": "string", "enum": ["ya", "tidak"], "default": "ya"},
                    "workers": {"type": "integer", "minimum": 1, "maximum": 8, "description": "Worker paralel."},
                    "model_id": {"type": "string"},
                    "region": {"type": "string"},
                },
            }}}},
            "responses": {
                "200": _json({"$ref": "#/components/schemas/Started"}, "Job dimulai"),
                "400": _error("Permintaan tidak valid (berkas, bab, mode, nama profil)"),
                "404": _error("Profil tidak ditemukan"),
                "413": _error("Berkas melebihi batas unggah"),
            },
        }},
        "/api/docx/status/{session_id}": {"get": {
            "summary": "Status dan progres per langkah",
            "parameters": [_SESSION],
            "responses": {
                "200": _json({"$ref": "#/components/schemas/Status"}, "Status job"),
                "404": _error("Sesi tidak ditemukan atau sudah kedaluwarsa"),
            },
        }},
        "/api/docx/report/{session_id}": {"get": {
            "summary": "Laporan verifikasi dan token job yang sudah selesai",
            "parameters": [_SESSION],
            "responses": {
                "200": _json({"$ref": "#/components/schemas/Result"}, "Laporan"),
                "404": _error("Sesi tidak ditemukan"),
                "409": _error("Job belum selesai, gagal, atau dibatalkan"),
            },
        }},
        "/api/docx/download/{session_id}": {"get": {
            "summary": "Unduh berkas DOCX",
            "parameters": [_SESSION],
            "responses": {
                "200": {"description": "Berkas DOCX", "content": {
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": {
                        "schema": {"type": "string", "format": "binary"}}}},
                "404": _error("Dokumen tidak ditemukan atau sudah kedaluwarsa"),
            },
        }},
        "/api/docx/cancel/{session_id}": {"post": {
            "summary": "Minta job berhenti",
            "description": (
                "Berlaku di antara slide, bab, dan langkah; panggilan model yang sedang berjalan diselesaikan dulu. "
                "Status berubah menjadi 'cancelled'. Token yang sudah terpakai tetap dicatat."
            ),
            "parameters": [_SESSION],
            "responses": {
                "202": _json({"type": "object", "properties": {"status": {"type": "string", "example": "cancelling"}}},
                             "Permintaan pembatalan diterima"),
                "404": _error("Sesi tidak ditemukan"),
                "409": _error("Job sudah berakhir"),
            },
        }},
        "/api/docx/{session_id}": {"delete": {
            "summary": "Hapus sesi dan berkas hasilnya",
            "parameters": [_SESSION],
            "responses": {
                "200": _json({"type": "object", "properties": {"dihapus": {"type": "boolean"}}}, "Terhapus"),
                "404": _error("Sesi tidak ditemukan"),
                "409": _error("Job masih berjalan; batalkan dulu"),
            },
        }},
        "/api/docx/docs": {"get": {
            "summary": "Halaman Swagger UI untuk mencoba API ini",
            "responses": {"200": {"description": "Halaman HTML", "content": {"text/html": {"schema": {"type": "string"}}}}},
        }},
        "/api/docx/openapi.json": {"get": {
            "summary": "Kontrak API ini",
            "responses": {"200": {"description": "Dokumen OpenAPI 3"}},
        }},
    },
    "components": {"schemas": {
        "Error": {"type": "object", "required": ["error"], "properties": {
            "error": {"type": "string"}, "status": {"type": "string"}}},
        "Step": {"type": "object", "properties": {
            "status": {"type": "string", "enum": ["pending", "running", "done", "skipped", "error", "cancelled"]},
            "detail": {"type": "string"}, "current": {"type": "integer"}, "total": {"type": "integer"}}},
        "Started": {"type": "object", "required": ["session_id", "steps", "initial_steps"], "properties": {
            "session_id": {"type": "string"},
            "steps": {"type": "array", "items": {"type": "object", "properties": {
                "key": {"type": "string"}, "label": {"type": "string"}}}},
            "initial_steps": {"type": "object", "additionalProperties": {"$ref": "#/components/schemas/Step"}}}},
        "Status": {"type": "object", "required": ["status", "steps", "elapsed_seconds"], "properties": {
            "status": {"type": "string", "enum": ["running", "done", "error", "cancelled"]},
            "steps": {"type": "object", "additionalProperties": {"$ref": "#/components/schemas/Step"}},
            "elapsed_seconds": {"type": "number"},
            "result": {"$ref": "#/components/schemas/Result"},
            "error": {"type": "string", "description": "Ada bila status 'error' atau 'cancelled'."}}},
        "Result": {"type": "object", "properties": {
            "filename": {"type": "string"},
            "download_url": {"type": "string"},
            "laporan": {"$ref": "#/components/schemas/Report"},
            "token": {"type": "object", "description": "Token per langkah model.", "additionalProperties": {
                "type": "object", "properties": {
                    "input_tokens": {"type": "integer"}, "output_tokens": {"type": "integer"}}}}}},
        "Report": {"type": "object", "properties": {
            "slide": {"type": "integer", "description": "Slide yang diproses (setelah sampul dan lampiran keluar)."},
            "gambar": {"type": "integer"}, "narasi": {"type": "integer"}, "gambar_diringkas": {"type": "integer"},
            "slide_dikecualikan": {"type": "array", "items": {"type": "string"}, "description": "Mis. 'slide-1 (sampul)'."},
            "slide_dilewati": {"type": "array", "items": {"type": "string"}},
            "slide_dipecah": {"type": "array", "items": {"type": "string"}},
            "slide_teks": {"type": "array", "items": {"type": "string"}},
            "slide_teks_asli": {"type": "array", "items": {"type": "string"}},
            "bab_kosong": {"type": "array", "items": {"type": "string"}},
            "aturan_bab": {"type": "object", "additionalProperties": {"type": "array", "items": {"type": "string"}}},
            "gambar_ganda": {"type": "array", "items": {"type": "string"}},
            "angka_tidak_didukung": {"type": "array", "items": {"type": "object", "properties": {
                "ref": {"type": "string"}, "angka": {"type": "array", "items": {"type": "string"}}}},
                "description": "Angka di narasi atau ringkasan yang tidak ada di teks slide; perlu diperiksa."},
            "bab_cadangan": {"type": "array", "items": {"type": "object", "properties": {
                "bab": {"type": "string"}, "alasan": {"type": "string"}}},
                "description": "Bab yang narasinya gagal sehingga memakai teks slide apa adanya."}}},
        "Options": {"type": "object", "properties": {
            "profil": {"type": "array", "items": {"type": "object", "properties": {
                "name": {"type": "string"}, "title": {"type": "string"}}}},
            "mapping_mode": {"type": "object"}, "split_panels": {"type": "object"},
            "workers": {"type": "object"}, "model_id": {"type": "string"},
            "region": {"type": "string", "nullable": True}, "bab": {"type": "object"},
            "batas_unggah_mb": {"type": "integer"}, "sesi_kedaluwarsa_detik": {"type": "integer"},
            "langkah": {"type": "array", "items": {"type": "object"}}}},
    }},
}
