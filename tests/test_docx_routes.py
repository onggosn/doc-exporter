import io
import json
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from PIL import Image

import app
import profile_store
from bedrock_service import BedrockError, NarrationPartial


def write_profile(root: Path) -> None:
    path = root / "dokumen-2d"
    path.mkdir()
    files = {
        "meta": {"title": "Dokumen 2d", "sources": []}, "layouts": [], "sections": [], "numbering": [],
        "captions": [], "row_patterns": [{"tokens": "I,T", "width_buckets": {"40/60": 3}}],
    }
    for key, value in files.items():
        (path / f"{key}.json").write_text(json.dumps(value), encoding="utf-8")


def slide(ref, title, text=(), reasons=(), image=None):
    item = {"ref": ref, "title": title, "text": list(text), "tables": [], "charts": [], "notes": "",
            "visual_desc": None, "vision_reasons": list(reasons)}
    return item


class FakeGenerator:
    usage_totals = {"narration": {"input_tokens": 1, "output_tokens": 2}}
    instances: list = []

    def __init__(self, **kwargs):
        self.kwargs = kwargs
        self.narrated: list[tuple[str, list[str]]] = []
        FakeGenerator.instances.append(self)

    describe_split_slides = True

    def describe_visuals(self, deck, progress):
        progress("visuals", "done", "", 1, 1)

    def summarize_images(self, deck, progress, skip_refs=frozenset()):
        self.summarized_skip = set(skip_refs)
        progress("captions", "done", "", 1, 1)

    def _map_slides(self, deck, chapters, mode, progress):
        progress("mapping", "done", "", 1, 1)
        return {chapters[0]: ["slide-1", "slide-2"]}

    def write_chapter_narration(self, chapter, slides, style_text):
        self.narrated.append((chapter, [s["ref"] for s in slides]))
        return {s["ref"]: [f"Narasi {s['ref']} 4,85%."] for s in slides}, {"slide-2": ["9.9"]}


class DocxJobTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        profiles = self.tmp / "profiles"
        profiles.mkdir()
        write_profile(profiles)
        patcher = patch.object(profile_store, "PROFILES_DIR", profiles)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(app.DOCX_SESSIONS.clear)
        FakeGenerator.instances.clear()
        log_patcher = patch.object(app, "_log_processing_times")  # jangan menulis ke PostgreSQL sungguhan
        self.log = log_patcher.start()
        self.addCleanup(log_patcher.stop)
        self.client = app.app.test_client()

    def run_job(self, generator=FakeGenerator, chapters=("Ringkasan",), appendix=False, pre_cancel=False):
        image = self.tmp / "slide-2.png"
        Image.new("RGB", (400, 300), "teal").save(image)
        deck = {"items": [slide("cover", "Judul Dokumen", ["Rapat Dewan Gubernur"]),
                          slide("slide-1", "Pengantar", ["PDB 4.85%", "PDB 4.85%", "Inflasi <3%"]),
                          slide("slide-2", "Grafik Kredit", ["Kredit 4.85%"], reasons=["chart"]),
                          slide("slide-3", "Penutup", ["Tindak lanjut dilakukan segera."]),
                          slide("slide-4", "Data Tambahan", ["Data tambahan disajikan di lampiran."])]}
        if appendix:
            deck["items"] += [slide("lamp-1", "", ["LAMPIRAN KEBIJAKAN EKONOMI DAN MONETER"]),
                              slide("lamp-2", "Data Lampiran", ["Isi lampiran yang cukup panjang."])]

        def attach(deck_, ppt, render_dir):
            next(item for item in deck_["items"] if item["ref"] == "slide-2")["image_path"] = str(image)
            return 1

        ppt = self.tmp / "contoh.pptx"
        ppt.write_bytes(b"x")
        work = self.tmp / "work"
        work.mkdir()
        (work / "slides").mkdir()
        app.DOCX_SESSIONS["sid"] = {
            "status": "running", "steps": {k: {"status": "pending", "detail": "", "current": 0, "total": 0}
                                           for k, _ in app.DOCX_STEPS},
            "result": None, "error": None, "file_path": None, "work_dir": str(work),
            "created_at": time.time(), "finished_at": None, "cancel": threading.Event(),
        }
        if pre_cancel:
            app.DOCX_SESSIONS["sid"]["cancel"].set()
        config = {"chapters": list(chapters), "profile": "dokumen-2d", "mapping_mode": "batch",
                  "model_id": "model-uji", "region": None, "document_name": "contoh.pptx"}
        with patch.object(app, "ingest_pptx", return_value=deck), \
             patch.object(app, "_attach_slide_images", side_effect=attach), \
             patch.object(app, "BedrockDocumentGenerator", generator):
            app._run_docx_job("sid", ppt, config)
        return ppt, work

    def test_successful_job_builds_docx_reports_and_cleans_up(self):
        ppt, work = self.run_job()
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        self.assertFalse(ppt.exists())
        self.assertFalse((work / "slides").exists())
        self.assertEqual(FakeGenerator.instances[0].kwargs, {"model_id": "model-uji", "region": None})

        document = Document(session["file_path"])
        self.assertEqual(len(document.inline_shapes), 1)
        texts = [p.text for p in document.paragraphs] + [c.text for t in document.tables for r in t.rows for c in r.cells]
        for expected in ("Ringkasan", "Narasi slide-1 4,85%.", "Narasi slide-2 4,85%."):
            self.assertTrue(any(expected in t for t in texts), expected)
        # Bawaan tanpa caption: gambar slide sudah memuat judulnya sendiri.
        self.assertFalse(any("Grafik Kredit" in t for t in texts))

        report = session["result"]["laporan"]
        self.assertEqual((report["gambar"], report["narasi"], report["gambar_ganda"]), (1, 2, []))
        self.assertEqual(report["angka_tidak_didukung"], [{"ref": "slide-2", "angka": ["9.9"]}])
        self.assertEqual(session["result"]["filename"], "contoh.docx")

    def test_duration_and_tokens_are_logged_like_the_existing_flow(self):
        class Tokens(FakeGenerator):
            usage_totals = {"visuals": {"input_tokens": 10, "output_tokens": 2},
                            "mapping": {"input_tokens": 30, "output_tokens": 4},
                            "narration": {"input_tokens": 500, "output_tokens": 90},
                            "style_guide": {"input_tokens": 1, "output_tokens": 1}}

            def describe_visuals(self, deck, progress):
                progress("visuals", "running", "", 0, 1)
                progress("visuals", "done", "", 1, 1)

        self.run_job(Tokens)
        name, status, timing, total_seconds, usage = self.log.call_args.args
        self.assertEqual((name, status), ("contoh.pptx [PPT→DOCX]", "done"))
        self.assertEqual(set(timing), {"ingest", "visuals", "mapping", "kb", "writing"})
        self.assertIsNotNone(timing["ingest"]["seconds"])
        self.assertIsNotNone(timing["visuals"]["at"])
        self.assertIsNone(timing["kb"]["seconds"])
        self.assertGreaterEqual(total_seconds, 0)
        self.assertEqual(usage, {"visuals": {"input_tokens": 10, "output_tokens": 2},
                                 "mapping": {"input_tokens": 30, "output_tokens": 4},
                                 "writing": {"input_tokens": 501, "output_tokens": 91}})  # narasi + panduan gaya

    def test_failed_job_is_logged_with_error_status(self):
        class Broken(FakeGenerator):
            def _map_slides(self, deck, chapters, mode, progress):
                raise BedrockError("gagal")

        self.run_job(Broken)
        self.assertEqual(self.log.call_args.args[1], "error")
        self.assertEqual(app.DOCX_SESSIONS["sid"]["status"], "error")

    def test_cover_and_appendix_are_excluded_from_every_step_and_reported(self):
        seen = {}

        class Spy(FakeGenerator):
            def describe_visuals(self, deck, progress):
                seen["visuals"] = [s["ref"] for s in deck["items"]]
                super().describe_visuals(deck, progress)

            def _map_slides(self, deck, chapters, mode, progress):
                seen["mapped"] = [s["ref"] for s in deck["items"]]
                return super()._map_slides(deck, chapters, mode, progress)

        self.run_job(Spy, appendix=True)
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        for step in ("visuals", "mapped"):
            self.assertFalse({"cover", "lamp-1", "lamp-2"} & set(seen[step]), step)
        report = session["result"]["laporan"]
        self.assertEqual(report["slide_dikecualikan"], ["cover (sampul)", "lamp-1 (lampiran)", "lamp-2 (lampiran)"])
        self.assertEqual(report["slide"], 4)

    def test_failed_chapter_falls_back_to_slide_text_and_is_reported(self):
        class Flaky(FakeGenerator):
            def _map_slides(self, deck, chapters, mode, progress):
                progress("mapping", "done", "", 1, 1)
                return dict(zip(chapters, (["slide-1"], ["slide-3"])))

            def write_chapter_narration(self, chapter, slides, style_text):
                if chapter == "Bab A":
                    raise BedrockError("Narasi bab 'Bab A' tidak valid")
                return super().write_chapter_narration(chapter, slides, style_text)

        self.run_job(Flaky, chapters=("Bab A", "Bab B"))
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        report = session["result"]["laporan"]
        self.assertEqual([item["bab"] for item in report["bab_cadangan"]], ["Bab A"])
        texts = [p.text for p in Document(session["file_path"]).paragraphs]
        # baris kembar dibuang, tanda < dibuang agar lolos pemeriksaan tag, baris pendek digabung satu butir
        self.assertIn("PDB 4.85%; Inflasi 3%", texts)
        self.assertIn("Narasi slide-3 4,85%.", texts)  # bab lain tetap ditulis model

    def test_partial_chapter_keeps_good_slides_and_falls_back_only_for_the_rest(self):
        class Partial(FakeGenerator):
            def _map_slides(self, deck, chapters, mode, progress):
                progress("mapping", "done", "", 1, 1)
                return {chapters[0]: ["slide-1", "slide-3"]}

            def write_chapter_narration(self, chapter, slides, style_text):
                good, bad = slides[0], slides[1]
                raise NarrationPartial("JSON terpotong", {good["ref"]: ["Narasi model slide-1 4,85%."]}, {}, [bad])

        self.run_job(Partial, chapters=("Bab A",))
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        report = session["result"]["laporan"]
        self.assertEqual(report["bab_cadangan"], [{"bab": "Bab A", "alasan": "1 dari 2 slide: JSON terpotong"}])
        texts = [p.text for p in Document(session["file_path"]).paragraphs]
        self.assertIn("Narasi model slide-1 4,85%.", texts)
        self.assertIn("Tindak lanjut dilakukan segera.", texts)

    def test_chapters_are_written_in_parallel_and_merged_in_chapter_order(self):
        import threading
        barrier = threading.Barrier(2, timeout=3)

        class Parallel(FakeGenerator):
            def _map_slides(self, deck, chapters, mode, progress):
                progress("mapping", "done", "", 1, 1)
                return dict(zip(chapters, (["slide-1"], ["slide-3"])))

            def write_chapter_narration(self, chapter, slides, style_text):
                barrier.wait()  # lolos hanya bila kedua bab berjalan bersamaan
                return super().write_chapter_narration(chapter, slides, style_text)

        original = app._run_docx_job

        def with_workers(session_id, ppt, config):
            return original(session_id, ppt, {**config, "workers": 2})

        with patch.object(app, "_run_docx_job", with_workers):
            self.run_job(Parallel, chapters=("Bab A", "Bab B"))
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        self.assertEqual(Parallel.instances[-1].workers, 2)
        texts = [p.text for p in Document(session["file_path"]).paragraphs]
        self.assertLess(texts.index("Bab A"), texts.index("Bab B"))
        self.assertEqual(session["steps"]["writing"]["status"], "done")

    def test_parallel_failures_still_stop_when_no_chapter_ever_succeeds(self):
        class Broken(FakeGenerator):
            def _map_slides(self, deck, chapters, mode, progress):
                progress("mapping", "done", "", 1, 1)
                return dict(zip(chapters, (["slide-1"], ["slide-3"], ["slide-4"], ["slide-2"])))

            def write_chapter_narration(self, chapter, slides, style_text):
                raise BedrockError("kredensial ditolak")

        original = app._run_docx_job
        with patch.object(app, "_run_docx_job", lambda sid, ppt, config: original(sid, ppt, {**config, "workers": 3})):
            self.run_job(Broken, chapters=("A", "B", "C", "D"))
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual((session["status"], session["error"]), ("error", "kredensial ditolak"))

    def test_job_stops_after_three_chapters_in_a_row_fail(self):
        class Broken(FakeGenerator):
            def _map_slides(self, deck, chapters, mode, progress):
                progress("mapping", "done", "", 1, 1)
                return dict(zip(chapters, (["slide-1"], ["slide-3"], ["slide-4"])))

            def write_chapter_narration(self, chapter, slides, style_text):
                raise BedrockError("kredensial ditolak")

        self.run_job(Broken, chapters=("A", "B", "C"))
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual((session["status"], session["error"]), ("error", "kredensial ditolak"))

    def test_job_stops_between_steps_when_cancelled_and_logs_the_tokens_spent_so_far(self):
        ran = []

        class Cancelling(FakeGenerator):
            usage_totals = {"visuals": {"input_tokens": 10, "output_tokens": 2}}

            def _map_slides(self, deck, chapters, mode, progress):
                progress("mapping", "running", "", 0, 1)
                app.DOCX_SESSIONS["sid"]["cancel"].set()  # pengguna menekan batal saat pemetaan
                progress("mapping", "running", "", 1, 1)  # callback progres berikutnya menghentikan job
                return super()._map_slides(deck, chapters, mode, progress)

            def summarize_images(self, deck, progress, skip_refs=frozenset()):
                ran.append("captions")

            def write_chapter_narration(self, chapter, slides, style_text):
                ran.append("writing")
                return super().write_chapter_narration(chapter, slides, style_text)

        self.run_job(Cancelling)
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual((session["status"], session["error"]), ("cancelled", "Dibatalkan oleh pengguna."))
        self.assertEqual(session["steps"]["mapping"]["status"], "cancelled")
        self.assertEqual(ran, [])  # langkah sesudah pembatalan tidak dijalankan
        self.assertIsNone(session["file_path"])
        name, status, _timing, _total, usage = self.log.call_args.args
        self.assertEqual(status, "cancelled")
        self.assertEqual(usage, {"visuals": {"input_tokens": 10, "output_tokens": 2}})
        status_response = self.client.get("/api/docx/status/sid").json
        self.assertEqual((status_response["status"], status_response["error"]), ("cancelled", "Dibatalkan oleh pengguna."))

    def test_job_cancelled_before_it_starts_stops_at_the_first_progress_call(self):
        self.run_job(pre_cancel=True)
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "cancelled")
        self.assertEqual(session["steps"]["ingest"]["status"], "cancelled")
        self.assertEqual(FakeGenerator.instances, [])  # model tidak pernah dipanggil
        self.assertEqual(self.log.call_args.args[1], "cancelled")

    def test_status_and_download_routes(self):
        self.run_job()
        status = self.client.get("/api/docx/status/sid").json
        self.assertEqual(status["status"], "done")
        self.assertEqual(status["result"]["download_url"], "/api/docx/download/sid")
        download = self.client.get("/api/docx/download/sid")
        self.assertEqual(download.status_code, 200)
        self.assertEqual(download.mimetype, app.DOCX_MIME)
        self.assertIn("contoh.docx", download.headers["Content-Disposition"])
        self.assertTrue(download.data.startswith(b"PK"))
        download.close()

    def test_error_in_job_is_reported_on_active_step(self):
        ppt = self.tmp / "x.pptx"
        ppt.write_bytes(b"x")
        work = self.tmp / "w"
        work.mkdir()
        app.DOCX_SESSIONS["sid"] = {
            "status": "running", "steps": {k: {"status": "pending", "detail": "", "current": 0, "total": 0}
                                           for k, _ in app.DOCX_STEPS},
            "result": None, "error": None, "file_path": None, "work_dir": str(work),
            "created_at": time.time(), "finished_at": None,
        }
        with patch.object(app, "ingest_pptx", return_value={"items": []}):
            app._run_docx_job("sid", ppt, {"chapters": ["A"], "profile": "dokumen-2d", "document_name": "x.pptx"})
        status = self.client.get("/api/docx/status/sid").json
        self.assertEqual((status["status"], status["error"]), ("error", "Presentasi tidak memiliki slide."))
        self.assertEqual(status["steps"]["ingest"]["status"], "error")
        self.assertFalse(ppt.exists())

    def test_unknown_session_and_unfinished_download(self):
        self.assertEqual(self.client.get("/api/docx/status/nope").status_code, 404)
        self.assertEqual(self.client.get("/api/docx/download/nope").status_code, 404)
        app.DOCX_SESSIONS["run"] = {"status": "running", "created_at": time.time()}
        self.assertEqual(self.client.get("/api/docx/download/run").status_code, 404)


class DocxApiTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.tmp = Path(tmp.name)
        profiles = self.tmp / "profiles"
        profiles.mkdir()
        write_profile(profiles)
        patcher = patch.object(profile_store, "PROFILES_DIR", profiles)
        patcher.start()
        self.addCleanup(patcher.stop)
        uploads = self.tmp / "uploads"
        uploads.mkdir()
        upload_patcher = patch.object(app, "UPLOAD_DIR", uploads)
        upload_patcher.start()
        self.addCleanup(upload_patcher.stop)
        self.addCleanup(app.DOCX_SESSIONS.clear)
        self.client = app.app.test_client()

    def session(self, status="running", **extra):
        work = self.tmp / f"work-{status}"
        work.mkdir(exist_ok=True)
        (work / "dokumen.docx").write_bytes(b"docx")
        app.DOCX_SESSIONS["s1"] = {
            "status": status, "steps": {}, "result": None, "error": None, "file_path": str(work / "dokumen.docx"),
            "work_dir": str(work), "created_at": time.time(), "finished_at": None, "cancel": threading.Event(), **extra,
        }
        return app.DOCX_SESSIONS["s1"], work

    def start(self, **overrides):
        data = {"presentation": (io.BytesIO(b"x"), "a.pptx"), "profile": "dokumen-2d"}
        data.update(overrides)
        with patch.object(app.threading, "Thread") as thread:
            response = self.client.post("/api/docx/start", data=data, content_type="multipart/form-data")
        return response, thread

    def test_openapi_contract_matches_the_routes_that_really_exist(self):
        spec = self.client.get("/api/docx/openapi.json").json
        self.assertEqual(spec["openapi"], "3.0.3")
        documented = {(path.replace("{session_id}", "<session_id>"), method.upper())
                      for path, item in spec["paths"].items() for method in item}
        real = {(rule.rule, method) for rule in app.app.url_map.iter_rules() if rule.rule.startswith("/api/docx/")
                for method in rule.methods - {"HEAD", "OPTIONS"}}
        self.assertEqual(documented, real)

    def test_swagger_ui_page_points_to_the_openapi_contract(self):
        response = self.client.get("/api/docx/docs")
        self.assertEqual(response.status_code, 200)
        self.assertIn("text/html", response.content_type)
        self.assertIn("/api/docx/openapi.json", response.get_data(as_text=True))

    def test_options_lists_defaults_limits_and_profiles(self):
        with patch.dict("os.environ", {"DOCX_WORKERS": ""}):
            options = self.client.get("/api/docx/options").json
        self.assertEqual([p["name"] for p in options["profil"]], ["dokumen-2d"])
        self.assertEqual(options["mapping_mode"], {"nilai": ["batch", "per_slide"], "bawaan": "batch"})
        self.assertEqual(options["workers"], {"bawaan": app.DOCX_DEFAULT_WORKERS, "min": 1, "maks": app.MAX_WORKERS})
        self.assertEqual(len(options["bab"]["bawaan_aplikasi"]), 17)
        self.assertEqual(options["bab"]["maksimal"], 20)
        self.assertEqual([step["key"] for step in options["langkah"]], [k for k, _ in app.DOCX_STEPS])
        self.assertEqual(options["sesi_kedaluwarsa_detik"], app.MAPPING_SESSION_TTL_SECONDS)

    def test_start_without_chapters_uses_profile_default_then_named_structure_then_app_default(self):
        response, thread = self.start()  # profil belum punya struktur: bawaan aplikasi
        self.assertEqual(response.status_code, 200)
        self.assertEqual(thread.call_args.kwargs["args"][2]["chapters"], app.DOCX_DEFAULT_CHAPTERS.splitlines())

        profile_store.save_chapters("dokumen-2d", {"bawaan": "Mei", "daftar": [
            {"nama": "Mei", "bab": ["1. Satu", "2. Dua"]}, {"nama": "Maret", "bab": ["A", "B", "C"]}]})
        _, thread = self.start()
        self.assertEqual(thread.call_args.kwargs["args"][2]["chapters"], ["1. Satu", "2. Dua"])
        _, thread = self.start(struktur=" maret ")
        self.assertEqual(thread.call_args.kwargs["args"][2]["chapters"], ["A", "B", "C"])
        _, thread = self.start(chapters="Bab Sendiri", struktur="Maret")  # 'chapters' menang atas 'struktur'
        self.assertEqual(thread.call_args.kwargs["args"][2]["chapters"], ["Bab Sendiri"])

    def test_unknown_structure_and_blank_chapters_are_rejected(self):
        profile_store.save_chapters("dokumen-2d", {"bawaan": "Mei", "daftar": [{"nama": "Mei", "bab": ["1. Satu"]}]})
        unknown, _ = self.start(struktur="Tidak Ada")
        self.assertEqual(unknown.status_code, 400)
        self.assertIn("tersedia: Mei", unknown.json["error"])
        blank, _ = self.start(chapters="  \n ")
        self.assertEqual(blank.status_code, 400)
        self.assertEqual(list((self.tmp / "uploads").iterdir()), [])

    def test_report_is_available_only_for_finished_jobs(self):
        self.assertEqual(self.client.get("/api/docx/report/tidak-ada").status_code, 404)
        self.session("running")
        running = self.client.get("/api/docx/report/s1")
        self.assertEqual((running.status_code, running.json["status"]), (409, "running"))
        self.session("cancelled", error="Dibatalkan oleh pengguna.")
        cancelled = self.client.get("/api/docx/report/s1")
        self.assertEqual((cancelled.status_code, cancelled.json["error"]), (409, "Dibatalkan oleh pengguna."))
        self.session("done", result={"filename": "a.docx", "laporan": {"slide": 3}, "token": {"visuals": {"input_tokens": 1, "output_tokens": 2}}})
        done = self.client.get("/api/docx/report/s1")
        self.assertEqual(done.status_code, 200)
        self.assertEqual((done.json["filename"], done.json["laporan"], done.json["download_url"]),
                         ("a.docx", {"slide": 3}, "/api/docx/download/s1"))

    def test_cancel_only_applies_to_running_jobs(self):
        self.assertEqual(self.client.post("/api/docx/cancel/tidak-ada").status_code, 404)
        session, _ = self.session("running")
        response = self.client.post("/api/docx/cancel/s1")
        self.assertEqual((response.status_code, response.json), (202, {"status": "cancelling"}))
        self.assertTrue(session["cancel"].is_set())
        session["status"] = "done"
        again = self.client.post("/api/docx/cancel/s1")
        self.assertEqual((again.status_code, again.json["status"]), (409, "done"))

    def test_delete_removes_finished_sessions_with_their_files_but_not_running_ones(self):
        self.assertEqual(self.client.delete("/api/docx/tidak-ada").status_code, 404)
        _, work = self.session("running")
        blocked = self.client.delete("/api/docx/s1")
        self.assertEqual(blocked.status_code, 409)
        self.assertTrue(work.exists())
        app.DOCX_SESSIONS["s1"]["status"] = "done"
        deleted = self.client.delete("/api/docx/s1")
        self.assertEqual((deleted.status_code, deleted.json), (200, {"dihapus": True}))
        self.assertNotIn("s1", app.DOCX_SESSIONS)
        self.assertFalse(work.exists())
        self.assertEqual(self.client.get("/api/docx/status/s1").status_code, 404)


class DefaultChaptersTests(unittest.TestCase):
    def test_docx_form_is_prefilled_with_the_default_chapter_structure(self):
        html = app.app.test_client().get("/").get_data(as_text=True)
        start = html.index("2. Buat DOCX dari PPT")
        textarea = html[html.index('<textarea name="chapters"', start):]
        textarea = textarea[:textarea.index("</textarea>")]
        lines = textarea[textarea.index(">") + 1:].splitlines()
        self.assertEqual(lines[0], "1. Kesimpulan Diskusi Hari 1 dan Arahan ADG")
        self.assertEqual(lines[-1], "4.7 Pokok-Pokok Komunikasi Kebijakan")
        self.assertEqual(len(lines), 17)
        self.assertEqual(lines, app.DOCX_DEFAULT_CHAPTERS.splitlines())

    def test_default_structure_is_accepted_by_the_chapter_reader(self):
        self.assertEqual(len(app._read_chapters(app.DOCX_DEFAULT_CHAPTERS)), 17)


class ReadWorkersTests(unittest.TestCase):
    def test_form_value_is_clamped_and_defaults_apply(self):
        with patch.dict("os.environ", {"DOCX_WORKERS": ""}):
            self.assertEqual([app._read_workers(v) for v in ("1", "6", "99", "0", "-3", "abc", "")],
                             [1, 6, app.MAX_WORKERS, 1, 1, app.DOCX_DEFAULT_WORKERS, app.DOCX_DEFAULT_WORKERS])
        with patch.dict("os.environ", {"DOCX_WORKERS": "2"}):
            self.assertEqual(app._read_workers(""), 2)


class DocxStartValidationTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        profiles = Path(tmp.name)
        write_profile(profiles)
        patcher = patch.object(profile_store, "PROFILES_DIR", profiles)
        patcher.start()
        self.addCleanup(patcher.stop)
        uploads = patch.object(app, "UPLOAD_DIR", Path(tmp.name) / "uploads")
        (Path(tmp.name) / "uploads").mkdir()
        uploads.start()
        self.addCleanup(uploads.stop)
        self.uploads = Path(tmp.name) / "uploads"
        self.client = app.app.test_client()

    def post(self, **overrides):
        data = {"presentation": (io.BytesIO(b"x"), "a.pptx"), "chapters": "Bab 1", "profile": "dokumen-2d"}
        data.update(overrides)
        return self.client.post("/api/docx/start", data=data, content_type="multipart/form-data")

    def test_rejects_invalid_requests_without_leaving_files(self):
        cases = [
            ({"presentation": (io.BytesIO(b""), "")}, 400, "Pilih file"),
            ({"presentation": (io.BytesIO(b"x"), "a.docx")}, 400, ".pptx"),
            ({"chapters": "  \n "}, 400, "bab"),
            ({"profile": "tidak-ada"}, 404, "tidak ditemukan"),
            ({"profile": "../x"}, 400, "tidak valid"),
            ({"mapping_mode": "aneh"}, 400, "Mode pemetaan"),
        ]
        for overrides, status, message in cases:
            with self.subTest(message=message):
                response = self.post(**overrides)
                self.assertEqual(response.status_code, status)
                self.assertIn(message, response.json["error"])
        self.assertEqual(list(self.uploads.iterdir()), [])

    def test_valid_request_starts_a_session_with_steps(self):
        with patch.object(app.threading, "Thread") as thread:
            response = self.post(model_id=" model-x ", region="")
            self.addCleanup(app.DOCX_SESSIONS.clear)
        self.assertEqual(response.status_code, 200)
        self.assertEqual([s["key"] for s in response.json["steps"]], [k for k, _ in app.DOCX_STEPS])
        thread.return_value.start.assert_called_once()
        config = thread.call_args.kwargs["args"][2]
        self.assertEqual((config["model_id"], config["region"], config["profile"]), ("model-x", None, "dokumen-2d"))
        self.assertEqual(response.json["initial_steps"]["writing"]["total"], 1)

    def test_expired_sessions_are_purged_with_their_files(self):
        work = Path(tempfile.mkdtemp())
        app.DOCX_SESSIONS["old"] = {"created_at": time.time() - app.MAPPING_SESSION_TTL_SECONDS - 5, "work_dir": str(work)}
        app._purge_expired_docx_sessions()
        self.assertNotIn("old", app.DOCX_SESSIONS)
        self.assertFalse(work.exists())


if __name__ == "__main__":
    unittest.main()
