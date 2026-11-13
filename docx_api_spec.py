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
                "required": ["presentation"],
                "properties": {
                    "presentation": {"type": "string", "format": "binary", "description": "Berkas .pptx."},
                    "profile": {"type": "string", "description": "Nama profil jenis dokumen. Bila kosong dipakai profil bawaan server (env DOCX_PROFILE, bawaan dokumen-2d)."},
                    "chapters": {"type": "string", "description": "Judul bab, satu per baris (1 sampai 20). Hanya dipakai bila server diatur DOCX_CHAPTERS_MODE=daftar; bawaannya (slide) mengikuti section PowerPoint dan mengabaikannya."},
                    "struktur": {"type": "string", "description": "Nama struktur bab tersimpan di profil. Hanya dipakai bila DOCX_CHAPTERS_MODE=daftar."},
                    "mapping_mode": {"type": "string", "enum": ["batch", "per_slide"], "default": "batch"},
                    "style_guide": {"type": "string", "enum": ["ya", "tidak"], "default": "ya",
                                    "description": "Pakai panduan gaya dari model (style_guide.json profil) di prompt narasi."},
                    "fact_review": {"type": "string", "enum": ["tidak", "aturan", "model"], "default": "aturan",
                                    "description": "Tinjauan fakta khusus ringkasan gambar (bukan narasi); hasilnya komentar Word, teks tidak diubah. aturan: angka, periode, arah perubahan tanpa model. model: aturan ditambah tinjauan model per slide yang ikut melihat gambarnya (menambah token)."},
                    "format_review": {"type": "string", "enum": ["ya", "tidak"], "default": "ya",
                                      "description": "Periksa format, urutan butir, dan penomoran pada halaman berbutir atau bernomor yang dibawa apa adanya dari slide; hasilnya komentar Word, teks tidak diubah."},
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
            "slide": {"type": "integer", "description": "Slide yang diproses (setelah sampul, lampiran, dan agenda keluar)."},
            "gambar": {"type": "integer"}, "narasi": {"type": "integer"}, "gambar_diringkas": {"type": "integer"},
            "slide_dipecah_kolom": {"type": "array", "items": {"type": "string"}, "description": "Slide berkolom di section bawaan yang dipecah per kolom menurut default_section.json profil, mis. 'slide-2 (3 kolom)'."},
            "slide_dikecualikan": {"type": "array", "items": {"type": "string"}, "description": "Mis. 'slide-1 (sampul)', 'slide-3 (agenda)'."},
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
                "description": "Bab yang narasinya gagal sehingga memakai teks slide apa adanya."},
            "tinjauan_format": {"type": "array", "description": "Catatan format, urutan butir, dan penomoran; tiap catatan juga menjadi komentar Word.", "items": {"type": "object", "properties": {
                "ref": {"type": "string"}, "jenis": {"type": "string", "enum": ["isi", "urutan", "format"]},
                "pesan": {"type": "string"}, "teks": {"type": "string"}}}},
            "tinjauan_fakta": {"type": "object", "description": "Catatan tinjauan fakta; tiap catatan juga menjadi komentar Word.", "properties": {
                "mode": {"type": "string", "enum": ["tidak", "aturan", "model"]},
                "jumlah_temuan": {"type": "integer"},
                "temuan": {"type": "array", "items": {"type": "object", "properties": {
                    "ref": {"type": "string"}, "sumber": {"type": "string", "enum": ["aturan", "model"]},
                    "jenis": {"type": "string", "description": "angka, periode, arah, atau model"},
                    "teks": {"type": "string"}, "pesan": {"type": "string"}}}},
                "slide_ditinjau_model": {"type": "array", "items": {"type": "string"}},
                "gagal_model": {"type": "array", "items": {"type": "object", "properties": {
                    "ref": {"type": "string"}, "alasan": {"type": "string"}}}}}}}},
        "Options": {"type": "object", "properties": {
            "profil": {"type": "array", "items": {"type": "object", "properties": {
                "name": {"type": "string"}, "title": {"type": "string"}}}},
            "mapping_mode": {"type": "object"}, "style_guide": {"type": "object"}, "fact_review": {"type": "object"}, "format_review": {"type": "object"}, "chapters_mode": {"type": "object"},
            "workers": {"type": "object"}, "model_id": {"type": "string"},
            "region": {"type": "string", "nullable": True}, "bab": {"type": "object"},
            "batas_unggah_mb": {"type": "integer"}, "sesi_kedaluwarsa_detik": {"type": "integer"},
            "langkah": {"type": "array", "items": {"type": "object"}}}},
    }},
}
