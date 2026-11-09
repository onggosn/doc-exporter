from __future__ import annotations

import json
import logging
import os
import re
import shutil
import tempfile
import threading
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path

import psycopg2
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

import doc_composer
from docx_api_spec import DOCX_OPENAPI
import profile_store
from bedrock_service import (
    DEFAULT_MODEL_ID, MAX_WORKERS, BedrockDocumentGenerator, BedrockError, NarrationPartial, run_parallel,
)
from docx_builder import DocxBuildError, build_docx, page_setup_from_profile
from ppt_adapter import ingest_pptx, render_slides_png

load_dotenv()


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", BASE_DIR / "uploads")).resolve()
ALLOWED_EXTENSIONS = {".pptx"}
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "1000"))
DEFAULT_KNOWLEDGE_BASE_ID = "2R7PFHPUO6"
DEFAULT_FAST_MODEL_ID = "moonshotai.kimi-k2.5"
MAPPING_SESSION_TTL_SECONDS = 3600
LOG_FILE = BASE_DIR / "proses.log"

_file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
_file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s"))
_file_handler.setLevel(logging.DEBUG)
logging.getLogger("bedrock_service").addHandler(_file_handler)
logging.getLogger("bedrock_service").setLevel(logging.DEBUG)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
# Default Werkzeug (500 KB) terlalu kecil untuk field teks biasa seperti
# html_fragment di KB Correction, yang membawa HTML + gambar base64 inline.
app.config["MAX_FORM_MEMORY_SIZE"] = 100 * 1024 * 1024
app.config["UPLOAD_DIR"] = UPLOAD_DIR
app.logger.addHandler(_file_handler)
app.logger.setLevel(logging.DEBUG)

# Sesi klasifikasi per-slide yang masih berjalan, menyimpan deck (termasuk gambar
# render sementara), progres mapping, dan konfigurasi generator sampai difinalisasi.
MAPPING_SESSIONS: dict[str, dict] = {}

# Sesi generate mode batch yang berjalan di background thread, menyimpan progres
# tiap langkah supaya frontend bisa polling status + progress bar.
GENERATE_SESSIONS: dict[str, dict] = {}
GENERATE_STEPS = [
    ("ingest", "Ekstrak slide dari presentasi"),
    ("visuals", "Mendeskripsikan visual slide"),
    ("mapping", "Memetakan slide ke bab"),
    ("kb", "Mengambil contoh dari Knowledge Base"),
    ("writing", "Menyusun isi tiap bab"),
]

# Pembuatan DOCX dari PPT berjalan di background thread; tiap sesi menyimpan progres,
# laporan, dan berkas hasil sampai kedaluwarsa.
DOCX_SESSIONS: dict[str, dict] = {}
DOCX_STEPS = [
    ("ingest", "Ekstrak slide dan render gambar"),
    ("visuals", "Mendeskripsikan visual slide"),
    ("mapping", "Memetakan slide ke bab"),
    ("captions", "Meringkas gambar untuk caption"),
    ("writing", "Menulis narasi tiap bab"),
    ("build", "Merakit dokumen DOCX"),
]
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

PG_DSN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": os.environ.get("PGPORT", "5432"),
    "dbname": os.environ.get("PGDATABASE", "ingestdb"),
    "user": os.environ.get("PGUSER", "onggonugroho"),
    "password": os.environ.get("PGPASSWORD", ""),
}


def _log_processing_times(
    document_name: str,
    status: str,
    step_timing: dict,
    total_seconds: float,
    usage_totals: dict[str, dict[str, int]] | None = None,
) -> None:
    """Simpan durasi + token per step generate ke PostgreSQL. Kegagalan di sini
    tidak boleh menggagalkan proses generate, cukup dicatat ke log aplikasi."""
    usage_totals = usage_totals or {}

    def tokens(step: str, key: str) -> int | None:
        return usage_totals.get(step, {}).get(key)

    total_input_tokens = sum(t.get("input_tokens", 0) for t in usage_totals.values())
    total_output_tokens = sum(t.get("output_tokens", 0) for t in usage_totals.values())

    try:
        with psycopg2.connect(**PG_DSN) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO document_processing_log (
                        document_name, status,
                        ingest_seconds, ingest_at, ingest_input_tokens, ingest_output_tokens,
                        visuals_seconds, visuals_at, visuals_input_tokens, visuals_output_tokens,
                        mapping_seconds, mapping_at, mapping_input_tokens, mapping_output_tokens,
                        kb_seconds, kb_at, kb_input_tokens, kb_output_tokens,
                        writing_seconds, writing_at, writing_input_tokens, writing_output_tokens,
                        total_seconds, total_input_tokens, total_output_tokens
                    ) VALUES (
                        %s, %s,
                        %s, %s, %s, %s,
                        %s, %s, %s, %s,
                        %s, %s, %s, %s,
                        %s, %s, %s, %s,
                        %s, %s, %s, %s,
                        %s, %s, %s
                    )
                    """,
                    (
                        document_name, status,
                        step_timing["ingest"]["seconds"], step_timing["ingest"]["at"],
                        tokens("ingest", "input_tokens"), tokens("ingest", "output_tokens"),
                        step_timing["visuals"]["seconds"], step_timing["visuals"]["at"],
                        tokens("visuals", "input_tokens"), tokens("visuals", "output_tokens"),
                        step_timing["mapping"]["seconds"], step_timing["mapping"]["at"],
                        tokens("mapping", "input_tokens"), tokens("mapping", "output_tokens"),
                        step_timing["kb"]["seconds"], step_timing["kb"]["at"],
                        tokens("kb", "input_tokens"), tokens("kb", "output_tokens"),
                        step_timing["writing"]["seconds"], step_timing["writing"]["at"],
                        tokens("writing", "input_tokens"), tokens("writing", "output_tokens"),
                        total_seconds, total_input_tokens, total_output_tokens,
                    ),
                )
    except Exception:
        app.logger.exception("Gagal menyimpan log waktu proses ke PostgreSQL.")


# Dokumen yang sudah selesai digenerate, disimpan untuk fitur "Part Document"
# (koreksi sebagian) — menyimpan deck versi teks saja (tanpa gambar) + hasil bab
# + referensi + config generator, supaya koreksi belakangan masih bisa cross-check
# fakta ke materi slide asli tanpa perlu upload ulang PPT.
DOCUMENT_SESSIONS: dict[str, dict] = {}
DOCUMENT_SESSION_TTL_SECONDS = 6 * 3600


def _slide_without_image(slide: dict) -> dict:
    return {key: value for key, value in slide.items() if key != "image_path"}


def _store_document_session(
    deck: dict, chapters_result: list[dict], references: list[dict], config: dict
) -> str:
    document_id = uuid.uuid4().hex
    DOCUMENT_SESSIONS[document_id] = {
        "deck": {
            "filename": deck["filename"],
            "items": [_slide_without_image(slide) for slide in deck["items"]],
        },
        "chapters_result": chapters_result,
        "references": references,
        "config": config,
        "created_at": time.time(),
    }
    return document_id


def _purge_expired_document_sessions() -> None:
    now = time.time()
    expired = [
        did for did, d in DOCUMENT_SESSIONS.items()
        if now - d["created_at"] > DOCUMENT_SESSION_TTL_SECONDS
    ]
    for did in expired:
        DOCUMENT_SESSIONS.pop(did, None)


def _cleanup_mapping_session(session_id: str) -> None:
    session = MAPPING_SESSIONS.pop(session_id, None)
    if session and session.get("render_dir"):
        shutil.rmtree(session["render_dir"], ignore_errors=True)


def _purge_expired_mapping_sessions() -> None:
    now = time.time()
    expired = [
        sid for sid, s in MAPPING_SESSIONS.items()
        if now - s["created_at"] > MAPPING_SESSION_TTL_SECONDS
    ]
    for sid in expired:
        _cleanup_mapping_session(sid)


def _purge_expired_generate_sessions() -> None:
    now = time.time()
    expired = [
        sid for sid, s in GENERATE_SESSIONS.items()
        if now - s["created_at"] > MAPPING_SESSION_TTL_SECONDS
    ]
    for sid in expired:
        GENERATE_SESSIONS.pop(sid, None)


def _save_upload(upload) -> Path:
    filename = secure_filename(upload.filename or "")
    if not filename or Path(filename).suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError("Unggah file PowerPoint dengan format .pptx.")
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    path = UPLOAD_DIR / f"{uuid.uuid4().hex}-{filename}"
    upload.save(path)
    return path


def _read_chapters(raw: str) -> list[str]:
    chapters = [line.strip() for line in raw.splitlines() if line.strip()]
    if not chapters:
        raise ValueError("Isi struktur bab minimal satu judul bab.")
    if len(chapters) > 20:
        raise ValueError("Maksimal 20 bab per dokumen.")
    return chapters


def _read_reference_pairs() -> list[dict[str, str]]:
    reference_ppts = request.files.getlist("reference_ppt")
    reference_docs = request.files.getlist("reference_doc")
    if len(reference_ppts) != len(reference_docs):
        raise ValueError("Lengkapi kedua file untuk setiap pasangan referensi.")
    pairs = []
    for ppt, doc in zip(reference_ppts, reference_docs):
        if not ppt or not ppt.filename:
            continue
        if not doc or not doc.filename:
            raise ValueError("Setiap PPT referensi harus memiliki dokumen final pasangannya.")
        ppt_path = _save_upload(ppt)
        try:
            reference = ingest_pptx(ppt_path, f"reference-{len(pairs) + 1}")
            final_text = doc.read().decode("utf-8", errors="replace").strip()
            if not final_text:
                raise ValueError(f"Dokumen final untuk {ppt.filename} kosong.")
            pairs.append(
                {
                    "ppt_text": json.dumps(reference["items"], ensure_ascii=False)[:30000],
                    "final_text": final_text[:30000],
                }
            )
        finally:
            ppt_path.unlink(missing_ok=True)
    if len(pairs) > 2:
        raise ValueError("Maksimal dua pasangan referensi untuk contoh gaya.")
    return pairs


@app.get("/")
def index():
    return render_template(
        "index.html",
        max_upload_mb=MAX_UPLOAD_MB,
        model_id=os.environ.get("BEDROCK_MODEL_ID") or DEFAULT_MODEL_ID,
        region=os.environ.get("AWS_REGION") or os.environ.get("AWS_DEFAULT_REGION", ""),
        knowledge_base_id=os.environ.get("BEDROCK_KNOWLEDGE_BASE_ID") or DEFAULT_KNOWLEDGE_BASE_ID,
        fast_model_id=os.environ.get("BEDROCK_FAST_MODEL_ID") or DEFAULT_FAST_MODEL_ID,
        docx_default_chapters=DOCX_DEFAULT_CHAPTERS,
    )


def _run_generate_job(session_id: str, ppt_path: Path, render_dir: str, config: dict) -> None:
    active_step = {"key": "ingest"}
    step_start: dict[str, float] = {}
    step_timing = {key: {"seconds": None, "at": None} for key, _ in GENERATE_STEPS}
    job_started_at = time.time()
    status_for_log = "error"
    generator: BedrockDocumentGenerator | None = None

    def progress_callback(step: str, status: str, detail: str = "", current: int = 0, total: int = 0) -> None:
        now = time.time()
        if status == "running":
            active_step["key"] = step
            step_start.setdefault(step, now)
        if status in ("done", "skipped", "error"):
            start = step_start.get(step, now)
            step_timing[step]["seconds"] = round(now - start, 2)
            step_timing[step]["at"] = datetime.now(timezone.utc)
        GENERATE_SESSIONS[session_id]["steps"][step] = {
            "status": status,
            "detail": detail,
            "current": current,
            "total": total,
        }

    try:
        progress_callback("ingest", "running", "", 0, 1)
        can_render = bool(
            (shutil.which("soffice") or shutil.which("libreoffice"))
            and shutil.which("pdftoppm")
        )
        ingest_options = {"render_dir": render_dir} if can_render else {}
        deck = ingest_pptx(ppt_path, source_id=uuid.uuid4().hex, **ingest_options)
        if not deck["items"]:
            raise ValueError("Presentasi tidak memiliki slide.")
        progress_callback("ingest", "done", f"{len(deck['items'])} slide", 1, 1)

        generator = BedrockDocumentGenerator(
            model_id=config["model_id"],
            region=config["region"],
            knowledge_base_id=config["knowledge_base_id"],
            fast_model_id=config.get("fast_model_id"),
        )
        result = generator.generate(
            deck,
            config["chapters"],
            config["references"],
            mapping_mode=config["mapping_mode"],
            progress_callback=progress_callback,
        )
        result["document_id"] = _store_document_session(
            deck, result["chapters"], config["references"], config
        )
        GENERATE_SESSIONS[session_id]["result"] = result
        GENERATE_SESSIONS[session_id]["status"] = "done"
        GENERATE_SESSIONS[session_id]["finished_at"] = time.time()
        status_for_log = "done"
    except (ValueError, BedrockError) as exc:
        GENERATE_SESSIONS[session_id]["steps"][active_step["key"]]["status"] = "error"
        GENERATE_SESSIONS[session_id]["status"] = "error"
        GENERATE_SESSIONS[session_id]["error"] = str(exc)
        GENERATE_SESSIONS[session_id]["finished_at"] = time.time()
    except Exception:
        app.logger.exception("Gagal membuat dokumen dari presentasi.")
        GENERATE_SESSIONS[session_id]["steps"][active_step["key"]]["status"] = "error"
        GENERATE_SESSIONS[session_id]["status"] = "error"
        GENERATE_SESSIONS[session_id]["error"] = "Terjadi kesalahan saat memproses file. Periksa log aplikasi."
        GENERATE_SESSIONS[session_id]["finished_at"] = time.time()
    finally:
        ppt_path.unlink(missing_ok=True)
        shutil.rmtree(render_dir, ignore_errors=True)
        total_seconds = round(time.time() - job_started_at, 2)
        usage_totals = generator.usage_totals if generator else {}
        _log_processing_times(
            config["document_name"], status_for_log, step_timing, total_seconds, usage_totals
        )


@app.post("/api/generate-start")
def generate_start():
    _purge_expired_generate_sessions()
    ppt_path = None
    render_dir = None
    try:
        ppt_upload = request.files.get("presentation")
        if not ppt_upload or not ppt_upload.filename:
            raise ValueError("Pilih file PowerPoint yang akan diolah.")
        ppt_path = _save_upload(ppt_upload)
        chapters = _read_chapters(request.form.get("chapters", ""))
        references = _read_reference_pairs()

        model_id = request.form.get("model_id", "").strip() or None
        fast_model_id = request.form.get("fast_model_id", "").strip() or None
        region = request.form.get("region", "").strip() or None
        knowledge_base_id = request.form.get("knowledge_base_id", "").strip() or None
        mapping_mode = request.form.get("mapping_mode", "batch").strip()
        if mapping_mode not in ("batch", "per_slide"):
            raise ValueError("Mode pemetaan slide tidak dikenali.")

        render_dir = tempfile.mkdtemp(prefix="ppt-slides-")
        session_id = uuid.uuid4().hex
        default_totals = {
            "ingest": 1,
            "visuals": 0,  # baru diketahui setelah ekstraksi slide selesai
            "mapping": (
                BedrockDocumentGenerator.MAP_SLIDES_MAX_ATTEMPTS
                if mapping_mode == "batch"
                else 0  # jumlah slide baru diketahui setelah ekstraksi selesai
            ),
            "kb": 1 if knowledge_base_id else 0,
            "writing": len(chapters),
        }
        GENERATE_SESSIONS[session_id] = {
            "status": "running",
            "steps": {
                key: {"status": "pending", "detail": "", "current": 0, "total": default_totals[key]}
                for key, _ in GENERATE_STEPS
            },
            "result": None,
            "error": None,
            "created_at": time.time(),
            "finished_at": None,
        }
        config = {
            "chapters": chapters,
            "references": references,
            "model_id": model_id,
            "fast_model_id": fast_model_id,
            "region": region,
            "knowledge_base_id": knowledge_base_id,
            "mapping_mode": mapping_mode,
            "document_name": ppt_upload.filename,
        }
        thread = threading.Thread(
            target=_run_generate_job, args=(session_id, ppt_path, render_dir, config), daemon=True
        )
        thread.start()
        ppt_path = None  # kepemilikan dialihkan ke thread, jangan dihapus di finally
        render_dir = None
        return jsonify(
            {
                "session_id": session_id,
                "steps": [{"key": key, "label": label} for key, label in GENERATE_STEPS],
                "initial_steps": GENERATE_SESSIONS[session_id]["steps"],
            }
        )
    except (ValueError, BedrockError) as exc:
        return jsonify({"error": str(exc)}), 400
    except RequestEntityTooLarge as exc:
        return handle_large_upload(exc)
    except Exception:
        app.logger.exception("Gagal memulai pembuatan dokumen.")
        return jsonify({"error": "Terjadi kesalahan saat memproses file. Periksa log aplikasi."}), 500
    finally:
        if ppt_path:
            ppt_path.unlink(missing_ok=True)
        if render_dir:
            shutil.rmtree(render_dir, ignore_errors=True)


@app.get("/api/generate-status/<session_id>")
def generate_status(session_id: str):
    session = GENERATE_SESSIONS.get(session_id)
    if not session:
        return jsonify({"error": "Sesi pembuatan dokumen tidak ditemukan atau sudah berakhir."}), 404
    reference_time = session["finished_at"] or time.time()
    response = {
        "status": session["status"],
        "steps": session["steps"],
        "elapsed_seconds": round(reference_time - session["created_at"], 1),
    }
    if session["status"] == "done":
        response["result"] = session["result"]
        GENERATE_SESSIONS.pop(session_id, None)
    elif session["status"] == "error":
        response["error"] = session["error"]
        GENERATE_SESSIONS.pop(session_id, None)
    return jsonify(response)


@app.post("/api/prepare-mapping")
def prepare_mapping():
    _purge_expired_mapping_sessions()
    prepare_started = time.time()
    ppt_path = None
    render_dir = None
    try:
        ppt_upload = request.files.get("presentation")
        if not ppt_upload or not ppt_upload.filename:
            raise ValueError("Pilih file PowerPoint yang akan diolah.")
        ppt_path = _save_upload(ppt_upload)
        chapters = _read_chapters(request.form.get("chapters", ""))
        references = _read_reference_pairs()

        model_id = request.form.get("model_id", "").strip() or None
        fast_model_id = request.form.get("fast_model_id", "").strip() or None
        region = request.form.get("region", "").strip() or None
        knowledge_base_id = request.form.get("knowledge_base_id", "").strip() or None

        render_dir = tempfile.mkdtemp(prefix="ppt-slides-")
        can_render = bool(
            (shutil.which("soffice") or shutil.which("libreoffice"))
            and shutil.which("pdftoppm")
        )
        ingest_options = {"render_dir": render_dir} if can_render else {}
        deck = ingest_pptx(ppt_path, source_id=uuid.uuid4().hex, **ingest_options)
        if not deck["items"]:
            raise ValueError("Presentasi tidak memiliki slide.")

        generator = BedrockDocumentGenerator(
            model_id=model_id, region=region, knowledge_base_id=knowledge_base_id,
            fast_model_id=fast_model_id,
        )
        visuals_started = time.time()
        generator.describe_visuals(deck)
        visuals_seconds = time.time() - visuals_started

        session_id = uuid.uuid4().hex
        MAPPING_SESSIONS[session_id] = {
            "document_name": ppt_upload.filename,
            "timing": {"ingest": visuals_started - prepare_started, "visuals": visuals_seconds, "mapping": 0.0},
            "deck": deck,
            "chapters": chapters,
            "references": references,
            "generator": generator,
            "render_dir": render_dir,
            "pending": {slide["ref"] for slide in deck["items"]},
            "mapping": {chapter: [] for chapter in chapters},
            "by_ref": {slide["ref"]: slide for slide in deck["items"]},
            "created_at": time.time(),
        }
        render_dir = None  # kepemilikan dialihkan ke sesi, jangan dihapus di finally
        return jsonify(
            {
                "session_id": session_id,
                "chapters": chapters,
                "slides": [
                    {"ref": slide["ref"], "title": slide.get("title") or ""}
                    for slide in deck["items"]
                ],
            }
        )
    except (ValueError, BedrockError) as exc:
        return jsonify({"error": str(exc)}), 400
    except RequestEntityTooLarge as exc:
        return handle_large_upload(exc)
    except Exception:
        app.logger.exception("Gagal menyiapkan slide untuk klasifikasi.")
        return jsonify({"error": "Terjadi kesalahan saat memproses file. Periksa log aplikasi."}), 500
    finally:
        if ppt_path:
            ppt_path.unlink(missing_ok=True)
        if render_dir:
            shutil.rmtree(render_dir, ignore_errors=True)


@app.post("/api/classify-slide")
def classify_slide_route():
    body = request.get_json(silent=True) or {}
    session_id = body.get("session_id", "")
    ref = body.get("ref", "")
    session = MAPPING_SESSIONS.get(session_id)
    if not session:
        return jsonify({"error": "Sesi klasifikasi tidak ditemukan atau sudah berakhir."}), 404
    if ref not in session["pending"]:
        return jsonify({"error": "Slide sudah diklasifikasikan atau tidak ditemukan."}), 400
    try:
        slide = session["by_ref"][ref]
        generator: BedrockDocumentGenerator = session["generator"]
        classify_started = time.time()
        chapter = generator.classify_slide(slide, session["chapters"])
        session["timing"]["mapping"] += time.time() - classify_started
        session["mapping"][chapter].append(ref)
        session["pending"].discard(ref)
        return jsonify(
            {
                "ref": ref,
                "title": slide.get("title") or "",
                "chapter": chapter,
                "remaining": len(session["pending"]),
                "done": len(session["pending"]) == 0,
            }
        )
    except BedrockError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception:
        app.logger.exception("Gagal mengklasifikasikan slide.")
        return jsonify({"error": "Terjadi kesalahan saat mengklasifikasikan slide."}), 500


@app.post("/api/finalize")
def finalize():
    body = request.get_json(silent=True) or {}
    session_id = body.get("session_id", "")
    session = MAPPING_SESSIONS.get(session_id)
    if not session:
        return jsonify({"error": "Sesi klasifikasi tidak ditemukan atau sudah berakhir."}), 404
    status = "error"
    logged = False
    try:
        if session["pending"]:
            return jsonify({"error": "Masih ada slide yang belum diklasifikasikan."}), 400
        logged = True
        writing_started = time.time()
        generator: BedrockDocumentGenerator = session["generator"]
        result = generator.generate_from_mapping(
            session["deck"], session["chapters"], session["references"], session["mapping"]
        )
        result["document_id"] = _store_document_session(
            session["deck"],
            result["chapters"],
            session["references"],
            {
                "model_id": generator.model_id,
                "region": generator.region,
                "knowledge_base_id": generator.knowledge_base_id,
            },
        )
        result["elapsed_seconds"] = round(time.time() - session["created_at"], 1)
        status = "done"
        return jsonify(result)
    except BedrockError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception:
        app.logger.exception("Gagal membuat dokumen dari pemetaan.")
        return jsonify({"error": "Terjadi kesalahan saat membuat dokumen."}), 500
    finally:
        if logged:  # satu baris per sesi: persiapan, klasifikasi tiap slide, dan penulisan digabung
            timing = {**session["timing"], "writing": time.time() - writing_started}
            _log_feature(f"{session['document_name']} [Mapping manual]", status, session["created_at"],
                         session["generator"], timing)
        _cleanup_mapping_session(session_id)


@app.post("/api/correct-selection")
def correct_selection():
    """Fitur 'Part Document': koreksi satu potongan HTML yang diseleksi user.
    Independen dari endpoint generate/mapping lama — hanya membaca DOCUMENT_SESSIONS."""
    _purge_expired_document_sessions()
    body = request.get_json(silent=True) or {}
    document_id = body.get("document_id", "")
    chapter_title = body.get("chapter", "")
    selected_html = (body.get("selected_html") or "").strip()
    doc = DOCUMENT_SESSIONS.get(document_id)
    if not doc:
        return jsonify({"error": "Dokumen tidak ditemukan atau sesi sudah berakhir."}), 404
    if not selected_html:
        return jsonify({"error": "Tidak ada teks yang diseleksi."}), 400
    chapter_entry = next((c for c in doc["chapters_result"] if c["title"] == chapter_title), None)
    if not chapter_entry:
        return jsonify({"error": "Bab tidak ditemukan di dokumen ini."}), 400
    generator = None
    started_at = time.time()
    status = "error"
    try:
        generator = BedrockDocumentGenerator(
            model_id=doc["config"]["model_id"],
            region=doc["config"]["region"],
            knowledge_base_id=doc["config"]["knowledge_base_id"],
        )
        by_ref = {slide["ref"]: slide for slide in doc["deck"]["items"]}
        materials = "\n\n".join(
            generator._slide_text(by_ref[ref])
            for ref in chapter_entry["slide_refs"]
            if ref in by_ref
        )
        knowledge_base_examples = generator._retrieve_examples(doc["deck"])
        reference_text = ""
        if doc["references"]:
            reference_text = "\n\nContoh pasangan PPT dan dokumen final untuk meniru gaya/struktur:\n"
            reference_text += "\n\n".join(
                f"Contoh {i} - isi PPT:\n{ref['ppt_text']}\nDokumen final:\n{ref['final_text']}"
                for i, ref in enumerate(doc["references"], start=1)
            )
        if knowledge_base_examples:
            reference_text += "\n\nContoh paling relevan dari Knowledge Base:\n"
            reference_text += "\n\n".join(
                f"Contoh {i}:\n{text}" for i, text in enumerate(knowledge_base_examples, start=1)
            )
        corrected_html = generator.correct_fragment(chapter_title, materials, reference_text, selected_html)
        status = "done"
        return jsonify({"corrected_html": corrected_html})
    except BedrockError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception:
        app.logger.exception("Gagal mengoreksi seleksi.")
        return jsonify({"error": "Terjadi kesalahan saat mengoreksi seleksi."}), 500
    finally:
        if generator is not None:
            _log_feature("[Part Document: koreksi bagian]", status, started_at, generator,
                         {"writing": time.time() - started_at})


@app.post("/api/correct-from-kb")
def correct_from_kb():
    """Menu baru 'KB Correction': user mengisi potongan konten di CKEditor, sistem
    koreksi berdasarkan materi PPT yang sudah ada di Knowledge Base. Independen dari
    endpoint generate/mapping/correct-selection yang sudah ada."""
    selected_html = request.form.get("html_fragment", "").strip()
    if not selected_html:
        return jsonify({"error": "Tulis atau tempel konten yang ingin dikoreksi."}), 400

    model_id = request.form.get("model_id", "").strip() or None
    region = request.form.get("region", "").strip() or None
    knowledge_base_id = request.form.get("knowledge_base_id", "").strip() or None

    generator = None
    started_at = None
    attempted = False
    status = "error"
    try:
        generator = BedrockDocumentGenerator(
            model_id=model_id, region=region, knowledge_base_id=knowledge_base_id
        )
        if not generator.knowledge_base_client or not generator.knowledge_base_id:
            return jsonify({"error": "Isi Knowledge Base ID — fitur ini butuh materi PPT dari KB."}), 400

        query_text = re.sub(r"<[^>]+>", " ", selected_html).strip()
        if not query_text:
            return jsonify({
                "error": "Konten hanya berisi gambar tanpa teks. Pencarian Knowledge Base "
                "berbasis teks — tambahkan sedikit teks (misal judul atau caption) di "
                "samping gambar supaya sistem bisa mencari materi relevan.",
            }), 400
        started_at = time.time()
        kb_materials = generator.retrieve_kb_materials(query_text)
        if not kb_materials:
            return jsonify({"error": "Tidak ada materi relevan ditemukan di Knowledge Base."}), 400
        materials = "\n\n".join(f"Materi {i}:\n{text}" for i, text in enumerate(kb_materials, start=1))

        attempted = True  # model dipanggil: pemakaian token layak dicatat
        result = generator.correct_fragment_from_kb(materials, selected_html)
        status = "done"
        return jsonify({"corrected_html": result["html"], "changes": result["changes"]})
    except BedrockError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception:
        app.logger.exception("Gagal mengoreksi HTML dari Knowledge Base.")
        return jsonify({"error": "Terjadi kesalahan saat mengoreksi HTML."}), 500
    finally:
        if attempted:
            _log_feature("[KB Correction]", status, started_at, generator, {"kb": time.time() - started_at})


def _purge_expired_docx_sessions() -> None:
    now = time.time()
    for sid in [s for s, d in DOCX_SESSIONS.items() if now - d["created_at"] > MAPPING_SESSION_TTL_SECONDS]:
        session = DOCX_SESSIONS.pop(sid)
        shutil.rmtree(session["work_dir"], ignore_errors=True)


def _attach_slide_images(deck: dict, ppt_path: Path, render_dir: Path) -> int:
    """Render slide yang harus jadi gambar (visual, chart, diagram, tabel) dan pasang image_path."""
    wanted = [i for i, slide in enumerate(deck["items"], start=1)
              if not slide.get("excluded") and doc_composer.slide_needs_image(slide)]
    if not wanted:
        return 0
    if not ((shutil.which("soffice") or shutil.which("libreoffice")) and shutil.which("pdftoppm")):
        raise ValueError("Slide bergambar perlu dirender: pasang LibreOffice (soffice) dan Poppler (pdftoppm).")
    pngs = render_slides_png(ppt_path, render_dir)
    attached = 0
    for index in wanted:
        if index in pngs:
            item = deck["items"][index - 1]
            if item.get("panels"):  # dipotong dari render utuh, sebelum margin slide dipangkas
                crops = doc_composer.crop_panels(str(pngs[index]), item["panels"], render_dir, item["ref"])
                if len(crops) >= 2:
                    item["panel_images"] = crops
            doc_composer.trim_margins(str(pngs[index]))
            item["image_path"] = str(pngs[index])
            attached += 1
    return attached


# Langkah yang punya kolom di tabel document_processing_log. "build" tidak punya kolom sendiri, tetapi
# durasinya sudah termasuk di total_seconds; "kb" tidak dipakai alur DOCX dan tersimpan kosong.
LOGGED_STEPS = ("ingest", "visuals", "mapping", "kb", "writing")
# Membedakan baris alur PPT → DOCX dari alur dokumen HTML pada tabel yang sama (tidak ada kolom jenis alur).
DOCX_LOG_SUFFIX = "[PPT→DOCX]"


def _usage_for_log(usage_totals: dict[str, dict[str, int]]) -> dict[str, dict[str, int]]:
    """Token per langkah model untuk tabel log. Panggilan narasi dan pembuatan panduan gaya dicatat di kolom
    writing_*, setara penulisan isi bab; langkah dengan nama sama dijumlahkan."""
    merged: dict[str, dict[str, int]] = {}
    for step, totals in usage_totals.items():
        key = "writing" if step in ("narration", "style_guide") else step
        if key not in LOGGED_STEPS:
            continue
        entry = merged.setdefault(key, {"input_tokens": 0, "output_tokens": 0})
        entry["input_tokens"] += totals.get("input_tokens", 0)
        entry["output_tokens"] += totals.get("output_tokens", 0)
    return merged


def _log_feature(
    name: str, status: str, started_at: float, generator: "BedrockDocumentGenerator | None" = None,
    seconds: dict[str, float] | None = None,
) -> None:
    """Catat satu pemakaian fitur (durasi dan token) ke document_processing_log, seperti alur generate.
    `seconds` memberi durasi per langkah; total dihitung dari started_at. Tidak pernah melempar error."""
    timing = {key: {"seconds": None, "at": None} for key in LOGGED_STEPS}
    for step, value in (seconds or {}).items():
        timing[step] = {"seconds": round(value, 2), "at": datetime.now(timezone.utc)}
    _log_processing_times(
        name, status, timing, round(time.time() - started_at, 2),
        _usage_for_log(generator.usage_totals if generator else {}),
    )


# Berapa bab berturut-turut boleh gagal sebelum proses dihentikan (tanda masalah kredensial atau jaringan).
DOCX_MAX_CHAPTER_FAILURES = 3


DOCX_DEFAULT_WORKERS = 4

# Isi awal kolom "Struktur bab" di form PPT → DOCX; pengguna bebas mengubahnya sebelum membuat dokumen.
DOCX_DEFAULT_CHAPTERS = """1. Kesimpulan Diskusi Hari 1 dan Arahan ADG
1.1 Kesimpulan dan Arahan Moneter
1.2 Kesimpulan dan Arahan Makroprudensial
1.3 Kesimpulan dan Arahan SP
2 Ringkasan Asesmen Evaluasi dan Rekomendasi Kebijakan
2.1 Ringkasan Asesmen dan Rekomendasi Kebijakan Moneter
2.2 Ringkasan Asesmen dan Rekomendasi Kebijakan Makroprudensial
2.3 Ringkasan Asesmen dan Rekomendasi Kebijakan SP
3. Proyeksi
4. Butir Keputusan
4.1 Kebijakan Moneter
4.2 Kebijakan Makroprudensial
4.3 Kebijakan Sistem Pembayaran
4.4 Kebijakan Pendukung
4.5 Koordinasi Kebijakan
4.6 Rekomendasi Tema Komunikasi Data
4.7 Pokok-Pokok Komunikasi Kebijakan"""


def _read_workers(raw: str) -> int:
    """Jumlah worker paralel dari form, atau DOCX_WORKERS, atau bawaan; dibatasi 1 sampai MAX_WORKERS."""
    try:
        value = int(raw or os.environ.get("DOCX_WORKERS") or DOCX_DEFAULT_WORKERS)
    except ValueError:
        value = DOCX_DEFAULT_WORKERS
    return max(1, min(value, MAX_WORKERS))


def _summary_warnings(deck: dict) -> dict[str, list[str]]:
    """Angka pada ringkasan gambar yang tidak ada di teks slide, per slide atau panel."""
    found: dict[str, list[str]] = {}
    for slide in deck["items"]:
        if slide.get("summary_flagged"):
            found[f"{slide['ref']} (ringkasan gambar)"] = slide["summary_flagged"]
        for number, panel in enumerate(slide.get("panel_images") or [], start=1):
            if panel.get("summary_flagged"):
                found[f"{slide['ref']} panel {number} (ringkasan gambar)"] = panel["summary_flagged"]
    return found


class JobCancelled(Exception):
    """Pengguna meminta job DOCX dibatalkan; dilempar dari callback progres di antara langkah dan item."""


class _Skipped(Exception):
    """Bab tidak dikerjakan karena proses sudah dihentikan oleh kegagalan beruntun."""


def _write_narrations(generator, chapters, mapping, by_ref, style_text, workers, progress):
    """Tulis narasi semua bab, paralel dengan `workers` thread. Bab yang gagal memakai teks slide apa adanya
    (seluruh bab, atau hanya slide yang gagal pada kegagalan sebagian). Bila DOCX_MAX_CHAPTER_FAILURES bab
    gagal total sebelum ada satu pun yang berhasil, proses dihentikan dengan error aslinya, karena itu tanda
    masalah kredensial atau jaringan. Mengembalikan (narasi, angka tak didukung, bab cadangan -> alasan)."""
    slides_of = {
        chapter: [] if doc_composer.is_verbatim_chapter(chapter) else [
            by_ref[ref] for ref in mapping[chapter]
            if doc_composer.has_substance(by_ref[ref]) and not doc_composer.is_filler(by_ref[ref])
        ]
        for chapter in chapters
    }
    work = [chapter for chapter in chapters if slides_of[chapter]]
    total = len(chapters)
    finished = total - len(work)
    state = {"hard_failures": 0, "successes": 0, "first_error": None}
    abort = threading.Event()

    def write(chapter):
        if abort.is_set():
            raise _Skipped()
        return generator.write_chapter_narration(chapter, slides_of[chapter], style_text)

    def done(index, outcome):
        nonlocal finished
        finished += 1
        ok, value = outcome
        if ok or (isinstance(value, NarrationPartial) and value.partial):
            state["successes"] += 1
        elif not isinstance(value, _Skipped):
            state["hard_failures"] += 1
            state["first_error"] = state["first_error"] or value
            if state["successes"] == 0 and state["hard_failures"] >= DOCX_MAX_CHAPTER_FAILURES:
                abort.set()
        progress("writing", "running", work[index], finished, total)

    progress("writing", "running", "", finished, total)
    outcomes = run_parallel(work, write, workers, stop_on_error=False, on_done=done)
    if abort.is_set():
        raise state["first_error"]

    narration: dict[str, list[str]] = {}
    warnings: dict[str, list[str]] = {}
    fallback: dict[str, str] = {}
    for chapter, (ok, value) in zip(work, outcomes):  # digabung sesuai urutan bab
        slides = slides_of[chapter]
        if ok:
            narration.update(value[0])
            warnings.update(value[1])
            continue
        if not isinstance(value, BedrockError):
            raise value
        partial = value if isinstance(value, NarrationPartial) else None
        if partial and partial.partial:  # sebagian slide berhasil: simpan, cadangan hanya untuk sisanya
            narration.update(partial.partial)
            warnings.update(partial.flagged)
        needing = partial.failed_slides if partial else slides
        app.logger.warning("Narasi bab '%s' gagal untuk %s slide, memakai teks slide: %s",
                           chapter, len(needing), value)
        narration.update(doc_composer.fallback_narration(needing))
        fallback[chapter] = (f"{len(needing)} dari {len(slides)} slide: " if partial else "") + str(value)
    progress("writing", "done", "", total, total)
    return narration, warnings, fallback


def _run_docx_job(session_id: str, ppt_path: Path, config: dict) -> None:
    session = DOCX_SESSIONS[session_id]
    work_dir = Path(session["work_dir"])
    active = {"key": "ingest"}
    generator: BedrockDocumentGenerator | None = None
    step_start: dict[str, float] = {}
    step_timing = {key: {"seconds": None, "at": None} for key in LOGGED_STEPS}
    job_started_at = time.time()

    def progress(step: str, status: str, detail: str = "", current: int = 0, total: int = 0) -> None:
        cancel = session.get("cancel")
        if cancel is not None and cancel.is_set():  # berlaku di antara slide, bab, dan langkah
            raise JobCancelled("Dibatalkan oleh pengguna.")
        now = time.time()
        if status == "running":
            active["key"] = step
            step_start.setdefault(step, now)
        logged = "visuals" if step == "captions" else step  # ringkasan gambar dihitung bersama langkah visual
        if status in ("done", "skipped", "error") and logged in step_timing:
            elapsed = now - step_start.get(step, now)
            step_timing[logged]["seconds"] = round((step_timing[logged]["seconds"] or 0) + elapsed, 2)
            step_timing[logged]["at"] = datetime.now(timezone.utc)
        session["steps"][step] = {"status": status, "detail": detail, "current": current, "total": total}

    try:
        progress("ingest", "running", "", 0, 1)
        deck = ingest_pptx(ppt_path, source_id=uuid.uuid4().hex)
        if not deck["items"]:
            raise ValueError("Presentasi tidak memiliki slide.")
        excluded = doc_composer.excluded_slides(deck["items"])  # sampul dan lampiran: keluar dari semua proses
        for item in deck["items"]:
            if item["ref"] in excluded:
                item["excluded"] = excluded[item["ref"]]
        images = _attach_slide_images(deck, ppt_path, work_dir / "slides")
        if not config.get("split_panels", True):
            for item in deck["items"]:
                item.pop("panel_images", None)
        for item in deck["items"]:
            doc_composer.split_verbatim(item)  # teks "Hasil Pembahasan Dewan Gubernur" dan "Key Takeaways" tetap asli
        skipped = [f"{item['ref']} ({item['excluded']})" for item in deck["items"] if item.get("excluded")]
        deck = {**deck, "items": [item for item in deck["items"] if not item.get("excluded")]}
        if not deck["items"]:
            raise ValueError("Semua slide dikecualikan (sampul dan lampiran); tidak ada yang bisa diproses.")
        progress("ingest", "done", f"{len(deck['items'])} slide, {images} gambar, {len(skipped)} dikecualikan", 1, 1)

        profile = profile_store.load_profile(config["profile"])
        generator = BedrockDocumentGenerator(model_id=config["model_id"], region=config["region"])
        generator.workers = config.get("workers", 1)
        generator.describe_split_slides = False  # slide yang dipecah cukup diringkas per panel
        generator.describe_visuals(deck, progress)
        chapters = config["chapters"]
        mapping = generator._map_slides(deck, chapters, config["mapping_mode"], progress)

        # Slide di bab kesimpulan/arahan dibawa apa adanya: tanpa ringkasan gambar dan tanpa narasi model.
        verbatim_refs = {ref for chapter in chapters if doc_composer.is_verbatim_chapter(chapter)
                         for ref in mapping[chapter]}
        generator.summarize_images(deck, progress, verbatim_refs)
        summary_warnings = _summary_warnings(deck)

        by_ref = {slide["ref"]: slide for slide in deck["items"]}
        style_text = doc_composer.style_context(profile["gaya"])
        narration, warnings, fallback = _write_narrations(
            generator, chapters, mapping, by_ref, style_text, config.get("workers", 1), progress
        )
        warnings.update(summary_warnings)

        progress("build", "running", "", 0, 1)
        spec, composed = doc_composer.compose(
            deck, chapters, mapping, narration, profile["settings"], profile["row_patterns"]
        )
        output = work_dir / "dokumen.docx"
        build_docx(spec, page_setup_from_profile(profile["page"]), output)
        progress("build", "done", "", 1, 1)

        composed["slide_dikecualikan"] = skipped
        session["result"] = {
            "filename": f"{Path(config['document_name']).stem}.docx",
            "laporan": doc_composer.build_report(spec, composed, warnings, fallback),
            "token": generator.usage_totals,
        }
        session["file_path"] = str(output)
        session["status"] = "done"
    except JobCancelled as exc:
        session["steps"][active["key"]]["status"] = "cancelled"
        session["status"], session["error"] = "cancelled", str(exc)
    except (ValueError, BedrockError, DocxBuildError, profile_store.ProfileNotFound) as exc:
        app.logger.exception("Pembuatan DOCX berhenti di langkah '%s': %s", active["key"], exc)
        session["steps"][active["key"]]["status"] = "error"
        session["status"], session["error"] = "error", str(exc)
    except Exception:
        app.logger.exception("Gagal membuat DOCX dari presentasi.")
        session["steps"][active["key"]]["status"] = "error"
        session["status"] = "error"
        session["error"] = "Terjadi kesalahan saat memproses file. Periksa log aplikasi."
    finally:
        session["finished_at"] = time.time()
        ppt_path.unlink(missing_ok=True)
        shutil.rmtree(work_dir / "slides", ignore_errors=True)  # gambar sudah tertanam di DOCX
        _log_processing_times(
            f"{config['document_name']} {DOCX_LOG_SUFFIX}",
            session["status"] if session["status"] in ("done", "cancelled") else "error",
            step_timing,
            round(time.time() - job_started_at, 2),
            _usage_for_log(generator.usage_totals if generator else {}),
        )


def _profile_chapters_text(profile_name: str, structure_name: str) -> str:
    """Struktur bab untuk permintaan tanpa 'chapters': struktur bernama (bila 'struktur' diisi), atau bawaan
    profil, atau bawaan aplikasi."""
    stored = profile_store.load_chapters(profile_name)
    wanted = " ".join(structure_name.split())
    if wanted:
        found = next((item for item in stored["daftar"] if item["nama"].lower() == wanted.lower()), None)
        if found is None:
            names = ", ".join(item["nama"] for item in stored["daftar"]) or "belum ada"
            raise ValueError(f"Struktur bab '{wanted}' tidak ada di profil (tersedia: {names}).")
        return "\n".join(found["bab"])
    default = next((item for item in stored["daftar"] if item["nama"] == stored["bawaan"]), None)
    return "\n".join(default["bab"]) if default else DOCX_DEFAULT_CHAPTERS


@app.post("/api/docx/start")
def docx_start():
    _purge_expired_docx_sessions()
    ppt_path = None
    work_dir = None
    try:
        upload = request.files.get("presentation")
        if not upload or not upload.filename:
            raise ValueError("Pilih file PowerPoint yang akan diolah.")
        if Path(upload.filename).suffix.lower() not in ALLOWED_EXTENSIONS:
            raise ValueError("File harus berformat .pptx.")
        profile_name = request.form.get("profile", "").strip()
        profile_store.load_profile(profile_name)  # memastikan profil ada sebelum menyimpan unggahan
        if "chapters" in request.form:
            chapters = _read_chapters(request.form["chapters"])
        else:  # klien API boleh tidak mengirim bab: struktur tersimpan profil, lalu bawaan aplikasi
            chapters = _read_chapters(_profile_chapters_text(profile_name, request.form.get("struktur", "")))
        mapping_mode = request.form.get("mapping_mode", "batch").strip()
        if mapping_mode not in ("batch", "per_slide"):
            raise ValueError("Mode pemetaan slide tidak dikenali.")
        ppt_path = _save_upload(upload)
        work_dir = tempfile.mkdtemp(prefix="docx-job-")

        session_id = uuid.uuid4().hex
        totals = {"ingest": 1, "visuals": 0, "mapping": 0, "captions": 0, "writing": len(chapters), "build": 1}
        DOCX_SESSIONS[session_id] = {
            "status": "running",
            "steps": {key: {"status": "pending", "detail": "", "current": 0, "total": totals[key]}
                      for key, _ in DOCX_STEPS},
            "result": None, "error": None, "file_path": None, "work_dir": work_dir,
            "created_at": time.time(), "finished_at": None, "cancel": threading.Event(),
        }
        config = {
            "chapters": chapters, "profile": profile_name, "mapping_mode": mapping_mode,
            "model_id": request.form.get("model_id", "").strip() or None,
            "region": request.form.get("region", "").strip() or None,
            "document_name": upload.filename,
            "split_panels": request.form.get("split_panels", "ya").strip() != "tidak",
            "workers": _read_workers(request.form.get("workers", "")),
        }
        threading.Thread(target=_run_docx_job, args=(session_id, ppt_path, config), daemon=True).start()
        ppt_path = work_dir = None  # kepemilikan dialihkan ke thread
        return jsonify({
            "session_id": session_id,
            "steps": [{"key": key, "label": label} for key, label in DOCX_STEPS],
            "initial_steps": DOCX_SESSIONS[session_id]["steps"],
        })
    except profile_store.ProfileNotFound as exc:
        return jsonify({"error": str(exc)}), 404
    except (ValueError, profile_store.ProfileError) as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception:
        app.logger.exception("Gagal memulai pembuatan DOCX.")
        return jsonify({"error": "Terjadi kesalahan saat memproses file. Periksa log aplikasi."}), 500
    finally:
        if ppt_path:
            ppt_path.unlink(missing_ok=True)
        if work_dir:
            shutil.rmtree(work_dir, ignore_errors=True)


@app.get("/api/docx/status/<session_id>")
def docx_status(session_id: str):
    session = DOCX_SESSIONS.get(session_id)
    if not session:
        return jsonify({"error": "Sesi pembuatan DOCX tidak ditemukan atau sudah berakhir."}), 404
    response = {
        "status": session["status"],
        "steps": session["steps"],
        "elapsed_seconds": round((session["finished_at"] or time.time()) - session["created_at"], 1),
    }
    if session["status"] == "done":
        response["result"] = {**session["result"], "download_url": f"/api/docx/download/{session_id}"}
    elif session["status"] in ("error", "cancelled"):
        response["error"] = session["error"]
    return jsonify(response)


@app.get("/api/docx/download/<session_id>")
def docx_download(session_id: str):
    session = DOCX_SESSIONS.get(session_id)
    if not session or session["status"] != "done":
        return jsonify({"error": "Dokumen tidak ditemukan atau sudah kedaluwarsa."}), 404
    return send_file(
        session["file_path"], mimetype=DOCX_MIME, as_attachment=True, download_name=session["result"]["filename"]
    )


@app.get("/api/docx/report/<session_id>")
def docx_report(session_id: str):
    """Laporan verifikasi dan pemakaian token sebuah job yang sudah selesai (tanpa unduhan)."""
    session = DOCX_SESSIONS.get(session_id)
    if not session:
        return jsonify({"error": "Sesi pembuatan DOCX tidak ditemukan atau sudah berakhir."}), 404
    if session["status"] != "done":
        message = session["error"] if session["status"] in ("error", "cancelled") else "Proses belum selesai."
        return jsonify({"error": message, "status": session["status"]}), 409
    return jsonify({**session["result"], "download_url": f"/api/docx/download/{session_id}"})


@app.post("/api/docx/cancel/<session_id>")
def docx_cancel(session_id: str):
    """Minta job berhenti. Berlaku di antara slide, bab, dan langkah; panggilan model yang sedang berjalan
    diselesaikan dulu. Cek hasilnya lewat status (akan menjadi 'cancelled')."""
    session = DOCX_SESSIONS.get(session_id)
    if not session:
        return jsonify({"error": "Sesi pembuatan DOCX tidak ditemukan atau sudah berakhir."}), 404
    if session["status"] != "running":
        return jsonify({"error": f"Proses sudah berakhir (status: {session['status']}).", "status": session["status"]}), 409
    session["cancel"].set()
    return jsonify({"status": "cancelling"}), 202


@app.delete("/api/docx/<session_id>")
def docx_delete(session_id: str):
    """Hapus sesi beserta berkas DOCX hasilnya tanpa menunggu kedaluwarsa."""
    session = DOCX_SESSIONS.get(session_id)
    if not session:
        return jsonify({"error": "Sesi pembuatan DOCX tidak ditemukan atau sudah berakhir."}), 404
    if session["status"] == "running":
        return jsonify({"error": "Proses masih berjalan; batalkan dulu lewat /api/docx/cancel.", "status": "running"}), 409
    DOCX_SESSIONS.pop(session_id, None)
    shutil.rmtree(session["work_dir"], ignore_errors=True)
    return jsonify({"dihapus": True})


@app.get("/api/docx/options")
def docx_options():
    """Nilai bawaan dan batas yang dipakai /api/docx/start, supaya klien tidak menebak."""
    return jsonify({
        "profil": [{"name": p["name"], "title": p["title"]} for p in profile_store.list_profiles()],
        "mapping_mode": {"nilai": ["batch", "per_slide"], "bawaan": "batch"},
        "split_panels": {"nilai": ["ya", "tidak"], "bawaan": "ya"},
        "workers": {"bawaan": _read_workers(""), "min": 1, "maks": MAX_WORKERS},
        "model_id": os.environ.get("BEDROCK_MODEL_ID") or DEFAULT_MODEL_ID,
        "region": os.environ.get("AWS_REGION") or os.environ.get("AWS_DEFAULT_REGION") or None,
        "bab": {"maksimal": 20, "bawaan_aplikasi": DOCX_DEFAULT_CHAPTERS.splitlines()},
        "batas_unggah_mb": MAX_UPLOAD_MB,
        "sesi_kedaluwarsa_detik": MAPPING_SESSION_TTL_SECONDS,
        "langkah": [{"key": key, "label": label} for key, label in DOCX_STEPS],
    })


@app.get("/api/docx/openapi.json")
def docx_openapi():
    return jsonify(DOCX_OPENAPI)


SWAGGER_UI_HTML = """<!doctype html>
<html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>API PPTX ke DOCX</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/swagger-ui/5.17.14/swagger-ui.min.css">
</head><body><div id="swagger-ui"></div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/swagger-ui/5.17.14/swagger-ui-bundle.min.js"></script>
<script>window.ui = SwaggerUIBundle({url: "/api/docx/openapi.json", dom_id: "#swagger-ui", deepLinking: true});</script>
</body></html>"""


@app.get("/api/docx/docs")
def docx_swagger_ui():
    return SWAGGER_UI_HTML, 200, {"Content-Type": "text/html; charset=utf-8"}


def _profile_response(action):
    try:
        return jsonify(action())
    except profile_store.ProfileNotFound as exc:
        return jsonify({"error": str(exc)}), 404
    except (profile_store.ProfileError, BedrockError) as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception:
        app.logger.exception("Gagal memproses profil format.")
        return jsonify({"error": "Terjadi kesalahan saat memproses profil format."}), 500


@app.get("/api/profiles")
def profiles_list():
    return _profile_response(lambda: {"profiles": profile_store.list_profiles()})


@app.get("/api/profiles/<name>")
def profile_detail(name: str):
    return _profile_response(lambda: profile_store.load_profile(name))


@app.put("/api/profiles/<name>/settings")
def profile_settings_save(name: str):
    body = request.get_json(silent=True)
    return _profile_response(
        lambda: {"settings": profile_store.save_settings(name, body), "settings_saved": True}
    )


@app.post("/api/profiles/<name>/style-guide")
def profile_style_guide(name: str):
    """Buat panduan gaya dari sampel teks profil lewat Bedrock. Sampel dikirim ke model."""
    body = request.get_json(silent=True) or {}

    def action():
        samples, statistics = profile_store.style_inputs(name)
        generator = BedrockDocumentGenerator(
            model_id=str(body.get("model_id", "")).strip() or None,
            region=str(body.get("region", "")).strip() or None,
        )
        started_at, status = time.time(), "error"
        try:
            guide, dropped = generator.generate_style_guide(samples, statistics)
            saved = profile_store.save_style_guide(name, guide, generator.model_id, dropped)
            status = "done"
        finally:
            _log_feature(f"{name} [Panduan gaya]", status, started_at, generator,
                         {"writing": time.time() - started_at})
        return {"panduan": saved, "usang": False}

    return _profile_response(action)


@app.put("/api/profiles/<name>/chapters")
def profile_chapters_save(name: str):
    body = request.get_json(silent=True)
    return _profile_response(lambda: {"struktur_bab": profile_store.save_chapters(name, body)})


@app.delete("/api/profiles/<name>/settings")
def profile_settings_reset(name: str):
    return _profile_response(
        lambda: {"settings": profile_store.reset_settings(name), "settings_saved": False}
    )


@app.errorhandler(RequestEntityTooLarge)
def handle_large_upload(_error):
    return jsonify({"error": f"Ukuran unggahan melebihi batas {MAX_UPLOAD_MB} MB."}), 413


if __name__ == "__main__":
    app.run(
        host=os.environ.get("HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", "5003")),
        debug=os.environ.get("FLASK_DEBUG", "").lower() == "true",
    )
