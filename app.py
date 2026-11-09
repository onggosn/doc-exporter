from __future__ import annotations

import logging
import os
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
from ppt_adapter import can_render_slides, ingest_pptx, render_slides_png

load_dotenv()


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", BASE_DIR / "uploads")).resolve()
ALLOWED_EXTENSIONS = {".pptx"}
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "1000"))
MAPPING_SESSION_TTL_SECONDS = 3600
LOG_FILE = Path(os.environ.get("LOG_FILE", BASE_DIR / "proses.log"))
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

_file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
_file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s"))
_file_handler.setLevel(logging.DEBUG)
logging.getLogger("bedrock_service").addHandler(_file_handler)
logging.getLogger("bedrock_service").setLevel(logging.DEBUG)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["UPLOAD_DIR"] = UPLOAD_DIR
app.logger.addHandler(_file_handler)
app.logger.setLevel(logging.DEBUG)

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


# Saklar penyimpanan log waktu proses ke PostgreSQL. Nilai awal dari env
# LOG_PROCESSING_ENABLED (default: aktif); bisa diubah saat runtime lewat
# set_log_processing_enabled() atau endpoint /api/log-processing.
# Catatan: disimpan di memori proses, kembali ke nilai env saat restart.
_LOG_PROCESSING_ENABLED = os.environ.get("LOG_PROCESSING_ENABLED", "true").strip().lower() not in (
    "0", "false", "off", "no",
)


def is_log_processing_enabled() -> bool:
    return _LOG_PROCESSING_ENABLED


def set_log_processing_enabled(enabled: bool) -> bool:
    """Aktif/nonaktifkan penyimpanan _log_processing_times. Return status baru."""
    global _LOG_PROCESSING_ENABLED
    _LOG_PROCESSING_ENABLED = bool(enabled)
    app.logger.info("Penyimpanan log waktu proses: %s", "ON" if _LOG_PROCESSING_ENABLED else "OFF")
    return _LOG_PROCESSING_ENABLED


def _log_processing_times(
    document_name: str,
    status: str,
    step_timing: dict,
    total_seconds: float,
    usage_totals: dict[str, dict[str, int]] | None = None,
) -> None:
    """Simpan durasi + token per step generate ke PostgreSQL. Kegagalan di sini
    tidak boleh menggagalkan proses generate, cukup dicatat ke log aplikasi.
    Dilewati bila saklar log dinonaktifkan."""
    if not _LOG_PROCESSING_ENABLED:
        return
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



@app.get("/")
def index():
    return render_template(
        "index.html",
        max_upload_mb=MAX_UPLOAD_MB,
        model_id=os.environ.get("BEDROCK_MODEL_ID") or DEFAULT_MODEL_ID,
        region=os.environ.get("AWS_REGION") or os.environ.get("AWS_DEFAULT_REGION", ""),
        docx_default_chapters=DOCX_DEFAULT_CHAPTERS,
    )




@app.route("/api/log-processing", methods=["GET", "POST"])
def api_log_processing():
    """GET: status saklar. POST: JSON {"enabled": true|false}, atau tanpa body untuk toggle."""
    if request.method == "POST":
        payload = request.get_json(silent=True) or {}
        enabled = payload.get("enabled")
        set_log_processing_enabled(not _LOG_PROCESSING_ENABLED if enabled is None else bool(enabled))
    return jsonify({"enabled": _LOG_PROCESSING_ENABLED})








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
    if not can_render_slides():
        raise ValueError("Slide bergambar perlu dirender: pasang LibreOffice (soffice) dan PyMuPDF (pip install pymupdf).")
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
# durasinya sudah termasuk di total_seconds; "kb" tidak dipakai dan tersimpan kosong (kolom tabel dipertahankan).
LOGGED_STEPS = ("ingest", "visuals", "mapping", "kb", "writing")
# Penanda baris alur PPT → DOCX pada tabel log (tidak ada kolom jenis alur).
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
    except RequestEntityTooLarge:
        raise  # ditangani handler 413 di bawah, bukan dijadikan 500
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
