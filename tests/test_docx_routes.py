import copy
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

    def run_job(self, generator=FakeGenerator, chapters=("Ringkasan",), appendix=False, pre_cancel=False,
                sections=None, kolom=None, **config_extra):
        image = self.tmp / "slide-2.png"
        Image.new("RGB", (400, 300), "teal").save(image)
        deck = {"items": [slide("cover", "Judul Dokumen", ["Rapat Dewan Gubernur"]),
                          slide("slide-1", "Pengantar", ["PDB 4.85%", "PDB 4.85%", "Inflasi <3%"]),
                          slide("slide-2", "Grafik Kredit", ["Kredit 4.85%"], reasons=["chart"]),
                          slide("slide-3", "Penutup", ["Tindak lanjut dilakukan segera."]),
                          slide("slide-4", "Data Tambahan", ["Data tambahan disajikan di lampiran."])]}
        if sections is not None:
            deck["sections"] = sections
        if kolom:
            deck["items"][1]["kolom"] = kolom
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
                  "model_id": "model-uji", "region": None, "document_name": "contoh.pptx",
                  "chapters_from_slides": sections is not None, **config_extra}
        with patch.object(app, "ingest_pptx", return_value=deck), \
             patch.object(app, "_attach_slide_images", side_effect=attach), \
             patch.object(app, "BedrockDocumentGenerator", generator):
            app._run_docx_job("sid", ppt, config)
        return ppt, work

    def test_job_following_slides_uses_sections_as_chapters_without_model_mapping(self):
        class NoMapping(FakeGenerator):
            def _map_slides(self, *args, **kwargs):
                raise AssertionError("pemetaan model tidak boleh dipanggil")

        sections = [{"nama": "Bagian A", "refs": ["slide-1", "slide-2"]},
                    {"nama": "Bagian B", "refs": ["slide-3"]},
                    {"nama": "Lampiran X", "refs": ["slide-4"]}]
        self.run_job(generator=NoMapping, sections=sections)
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        self.assertEqual(session["steps"]["mapping"]["status"], "skipped")
        headings = [p.text for p in Document(session["file_path"]).paragraphs if p.style.name.startswith("Heading 1")]
        self.assertEqual(headings, ["Bagian A", "Bagian B"])  # Lampiran X dilewati
        self.assertIn("slide-4 (lampiran)", session["result"]["laporan"]["slide_dikecualikan"])

    DEFAULT_SECTION = {
        "bagian_cocok": "Default Section", "judul_bab": "1. Kesimpulan Diskusi Hari 1 dan Arahan ADG",
        "judul_kolom": "1.{n} Kesimpulan dan Arahan {nama}",
        "kolom": [{"cocok": "MONETER", "nama": "Moneter"}, {"cocok": "SP", "nama": "Sistem Pembayaran"}]}
    COLUMNS = [{"header": "SP PUR", "text": ["KESIMPULAN", "K#1: isi SP"]},
               {"header": "MONETER", "text": ["KESIMPULAN", "K1: isi moneter"]}]
    SECTIONS = [{"nama": "Default Section", "refs": ["slide-1"]}, {"nama": "Bagian B", "refs": ["slide-3"]}]

    def headings(self):
        return [(p.style.name, p.text) for p in Document(app.DOCX_SESSIONS["sid"]["file_path"]).paragraphs
                if p.style.name.startswith("Heading")]

    def test_job_maps_the_default_section_with_the_profile_and_reports_the_split(self):
        (self.tmp / "profiles" / "dokumen-2d" / "default_section.json").write_text(
            json.dumps(self.DEFAULT_SECTION), encoding="utf-8")
        self.run_job(sections=copy.deepcopy(self.SECTIONS), kolom=copy.deepcopy(self.COLUMNS))
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        self.assertEqual(self.headings(), [
            ("Heading 1", "1. Kesimpulan Diskusi Hari 1 dan Arahan ADG"),
            ("Heading 2", "1.1 Kesimpulan dan Arahan Moneter"),
            ("Heading 2", "1.2 Kesimpulan dan Arahan Sistem Pembayaran"),
            ("Heading 1", "Bagian B"), ("Heading 2", "Penutup")])
        self.assertEqual(session["result"]["laporan"]["slide_dipecah_kolom"], ["slide-1 (2 kolom)"])
        texts = [p.text for p in Document(session["file_path"]).paragraphs]
        self.assertIn("K1: isi moneter", " ".join(texts))

    def test_job_without_default_section_config_exports_the_slide_as_is(self):
        self.run_job(sections=copy.deepcopy(self.SECTIONS), kolom=copy.deepcopy(self.COLUMNS))  # profil uji tidak punya default_section.json
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        self.assertEqual([h for h in self.headings() if h[0] == "Heading 1"], [("Heading 1", "Pengantar"), ("Heading 1", "Bagian B")])
        self.assertEqual(session["result"]["laporan"]["slide_dipecah_kolom"], [])
        self.assertNotIn("K1: isi moneter", " ".join(p.text for p in Document(session["file_path"]).paragraphs))

    def test_job_following_slides_fails_clearly_when_the_deck_has_no_sections(self):
        self.run_job(sections=[])
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "error")
        self.assertIn("section PowerPoint", session["error"])

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

    class WrongSummary(FakeGenerator):
        def summarize_images(self, deck, progress, skip_refs=frozenset()):
            for item in deck["items"]:
                if item.get("image_path"):
                    item["summary"] = "Kredit tumbuh 77,7% pada Juli 2027."
            progress("captions", "done", "", 1, 1)

    def comments_of(self):
        return [(c.author, c.text) for c in Document(app.DOCX_SESSIONS["sid"]["file_path"]).comments]

    def test_fact_review_rules_comment_only_image_summaries_without_changing_text(self):
        self.run_job(self.WrongSummary)  # bawaan: aturan
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        self.assertEqual(session["steps"]["review"]["status"], "done")
        review = session["result"]["laporan"]["tinjauan_fakta"]
        self.assertEqual(review["mode"], "aturan")
        self.assertEqual({t["ref"] for t in review["temuan"]}, {"slide-2"})  # hanya slide bergambar; narasi tidak
        comments = self.comments_of()
        self.assertEqual(len(comments), review["jumlah_temuan"])
        self.assertEqual({author for author, _ in comments}, {"Peninjau Fakta"})
        self.assertTrue(all(text.startswith("[Cek fakta:") for _, text in comments))
        document = Document(session["file_path"])
        cell_texts = [c.text for t in document.tables for r in t.rows for c in r.cells]
        self.assertTrue(any("Kredit tumbuh 77,7% pada Juli 2027." in t for t in cell_texts))  # teks tetap

    def test_fact_review_off_skips_the_step_and_adds_no_fact_comments(self):
        self.run_job(self.WrongSummary, fact_review="tidak")
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        self.assertEqual(session["steps"]["review"]["status"], "skipped")
        self.assertEqual(session["result"]["laporan"]["tinjauan_fakta"]["mode"], "tidak")
        self.assertEqual(self.comments_of(), [])

    def test_fact_review_model_mode_adds_model_findings_and_logs_tokens_as_writing(self):
        class Model(self.WrongSummary):
            usage_totals = {"narration": {"input_tokens": 10, "output_tokens": 1},
                            "review": {"input_tokens": 5, "output_tokens": 2}}
            workers = 1

            def review_slide(self, slide, paragraphs):
                return [{"paragraf": 1, "kutipan": "tumbuh", "alasan": "Gambar tidak menunjukkan pertumbuhan."}]

        self.run_job(Model, fact_review="model")
        session = app.DOCX_SESSIONS["sid"]
        self.assertEqual(session["status"], "done", session["error"])
        review = session["result"]["laporan"]["tinjauan_fakta"]
        self.assertEqual(review["slide_ditinjau_model"], ["slide-2"])
        self.assertTrue(any(text.startswith("[Cek fakta: model]") for _, text in self.comments_of()))
        self.assertEqual(self.log.call_args.args[4]["writing"], {"input_tokens": 15, "output_tokens": 3})

    def test_report_always_carries_the_format_review_list(self):
        self.run_job()
        self.assertEqual(app.DOCX_SESSIONS["sid"]["result"]["laporan"]["tinjauan_format"], [])

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
        self.assertEqual(set(timing), {"ingest", "visuals", "captions", "mapping", "writing"})
        self.assertIsNotNone(timing["ingest"]["seconds"])
        self.assertIsNotNone(timing["visuals"]["at"])
        self.assertEqual(timing["captions"]["seconds"], 0)  # dilewati (tanpa gambar), tetap tercatat sendiri
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
        chapters_mode = patch.dict("os.environ", {"DOCX_CHAPTERS_MODE": "daftar"})  # tes struktur bab memakai mode daftar
        chapters_mode.start()
        self.addCleanup(chapters_mode.stop)
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

    STYLE = {
        "statistik": {"angka": {"desimal_koma": 10, "desimal_titik": 1, "persen_rapat": 10, "persen_spasi": 0,
                                "rupiah_rapat": 5, "rupiah_spasi": 0}, "miring": {"jumlah": 3}},
        "panduan": {"panduan": {
            "nada": "Formal-teknokratis", "sudut_pandang": "Orang ketiga", "panjang_kalimat": "16 kata",
            "gaya_poin": "Kalimat utuh", "istilah_dan_singkatan": "SRBI", "format_angka": "Koma",
            "istilah_asing": "Miring", "larangan": ["Jangan pakai saya"], "contoh": []}},
    }

    def test_style_guide_toggle_keeps_format_rules_but_drops_the_model_guide(self):
        on, off = app._style_text(self.STYLE, True), app._style_text(self.STYLE, False)
        self.assertIn("Nada: Formal-teknokratis", on)
        self.assertNotIn("Nada:", off)
        self.assertNotIn("Hindari:", off)
        self.assertTrue(off)  # aturan format dari statistik tetap ada
        self.assertTrue(on.startswith(off))
        self.assertEqual(app._style_text(None, False), "")

    def test_sections_become_chapters_skipping_appendix_empty_and_generic_names(self):
        deck = {"items": [{"ref": f"slide-{i}", "title": f"Judul {i}"} for i in (2, 3, 4, 5, 6, 8)],
                "sections": [
                    {"nama": "Default Section", "refs": ["slide-1", "slide-2", "slide-3"]},  # slide-1 sudah dikecualikan
                    {"nama": "Moneter", "refs": ["slide-4", "slide-5"]},
                    {"nama": "Kosong", "refs": ["slide-1"]},
                    {"nama": "Moneter", "refs": ["slide-6"]},
                    {"nama": "Lampiran Moneter", "refs": ["slide-8"]}]}
        chapters, mapping = app._slide_chapters(deck)
        self.assertEqual(chapters, ["Judul 2", "Moneter", "Moneter (2)"])
        self.assertEqual(mapping, {"Judul 2": ["slide-2", "slide-3"], "Moneter": ["slide-4", "slide-5"],
                                   "Moneter (2)": ["slide-6"]})
        with self.assertRaises(ValueError):
            app._slide_chapters({"items": [{"ref": "slide-1"}], "sections": [{"nama": "Lampiran", "refs": ["slide-1"]}]})

    def test_chapters_follow_slides_by_default_and_ignore_the_chapter_list(self):
        with patch.dict("os.environ", {"DOCX_CHAPTERS_MODE": ""}):
            _, thread = self.start(chapters="Bab A\nBab B")
        config = thread.call_args.kwargs["args"][2]
        self.assertEqual((config["chapters_from_slides"], config["chapters"]), (True, []))

    def test_chapters_mode_is_switched_by_server_env_only_not_by_a_request_parameter(self):
        with patch.dict("os.environ", {"DOCX_CHAPTERS_MODE": "daftar"}):
            _, thread = self.start(chapters="Bab A\nBab B")
            config = thread.call_args.kwargs["args"][2]
            self.assertEqual((config["chapters_from_slides"], config["chapters"]), (False, ["Bab A", "Bab B"]))
        with patch.dict("os.environ", {"DOCX_CHAPTERS_MODE": "slide"}):
            _, thread = self.start(chapters_mode="daftar", chapters="Bab A")  # parameter lama diabaikan
            self.assertTrue(thread.call_args.kwargs["args"][2]["chapters_from_slides"])
        with patch.dict("os.environ", {"DOCX_CHAPTERS_MODE": "aneh"}):
            response, _ = self.start()
            self.assertEqual(response.status_code, 400)

    def test_profile_defaults_to_dokumen_2d_and_can_be_switched_by_env_only(self):
        with patch.dict("os.environ", {"DOCX_PROFILE": ""}):
            response, thread = self.start(profile="")
            self.assertEqual(response.status_code, 200)
            self.assertEqual(thread.call_args.kwargs["args"][2]["profile"], "dokumen-2d")
            html = self.client.get("/").get_data(as_text=True)
            self.assertIn('data-profile="dokumen-2d"', html)
        with patch.dict("os.environ", {"DOCX_PROFILE": "tidak-ada"}):
            response, _ = self.start(profile="")
            self.assertEqual(response.status_code, 404)  # profil bawaan server tidak ditemukan

    def test_page_has_no_profile_picker_and_every_element_the_script_needs(self):
        import re
        html = re.sub(r"<!--.*?-->", "", self.client.get("/").get_data(as_text=True), flags=re.S)  # tanpa blok yang dikomentari
        script = Path("static/docx-convert.js").read_text(encoding="utf-8")
        self.assertNotIn("docx-profile-select", html)
        for element_id in set(re.findall(r'querySelector\("#([\w-]+)"\)', script)):
            self.assertIn(f'id="{element_id}"', html, element_id)  # elemen hilang = skrip berhenti, tombol tetap disabled

    def test_start_reads_style_guide_toggle(self):
        for overrides, expected in (({"style_guide": "tidak"}, False), ({"style_guide": "ya"}, True), ({}, True)):
            _, thread = self.start(**overrides)
            self.assertIs(thread.call_args.kwargs["args"][2]["use_style_guide"], expected, overrides)

    def test_start_reads_and_validates_fact_review_mode(self):
        for overrides, expected in (({}, "aturan"), ({"fact_review": "model"}, "model"), ({"fact_review": "tidak"}, "tidak")):
            _, thread = self.start(**overrides)
            self.assertEqual(thread.call_args.kwargs["args"][2]["fact_review"], expected, overrides)
        response, _ = self.start(fact_review="semua")
        self.assertEqual(response.status_code, 400)

    def test_start_reads_format_review_toggle(self):
        for overrides, expected in (({}, True), ({"format_review": "ya"}, True), ({"format_review": "tidak"}, False)):
            _, thread = self.start(**overrides)
            self.assertIs(thread.call_args.kwargs["args"][2]["format_review"], expected, overrides)

    def test_options_lists_defaults_limits_and_profiles(self):
        with patch.dict("os.environ", {"DOCX_WORKERS": ""}):
            options = self.client.get("/api/docx/options").json
        self.assertEqual([p["name"] for p in options["profil"]], ["dokumen-2d"])
        self.assertEqual(options["mapping_mode"], {"nilai": ["batch", "per_slide"], "bawaan": "batch"})
        self.assertEqual(options["style_guide"], {"nilai": ["ya", "tidak"], "bawaan": "ya"})
        self.assertEqual(options["chapters_mode"]["bawaan"], "slide")
        self.assertEqual(options["chapters_mode"]["nilai"], ["slide", "daftar"])
        self.assertEqual(options["fact_review"], {"nilai": ["tidak", "aturan", "model"], "bawaan": "aturan"})
        self.assertEqual(options["format_review"], {"nilai": ["ya", "tidak"], "bawaan": "ya"})
        self.assertEqual({k: options["workers"][k] for k in ("bawaan", "min", "maks")},
                         {"bawaan": app.DOCX_DEFAULT_WORKERS, "min": 1, "maks": app.MAX_WORKERS})
        self.assertNotIn("split_panels", options)
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
    def test_docx_form_has_no_chapter_structure_or_mapping_fields_because_chapters_follow_slides(self):
        html = app.app.test_client().get("/").get_data(as_text=True)
        form = html[html.index('id="docx-create-form"'):html.index("docx-progress-panel")]
        for removed in ('name="chapters"', 'name="chapters_mode"', 'name="mapping_mode"', "docx-chapter-select",
                        'name="split_panels"', 'name="workers"'):
            self.assertNotIn(removed, form)
        self.assertIn('name="presentation"', form)

    def test_default_structure_is_accepted_by_the_chapter_reader(self):
        self.assertEqual(len(app._read_chapters(app.DOCX_DEFAULT_CHAPTERS)), 17)


class ReadWorkersTests(unittest.TestCase):
    def test_default_is_eight_and_env_is_clamped(self):
        self.assertEqual(app.DOCX_DEFAULT_WORKERS, 8)
        with patch.dict("os.environ", {"DOCX_WORKERS": ""}):
            self.assertEqual(app._read_workers(), 8)
        for raw, expected in (("2", 2), ("99", app.MAX_WORKERS), ("0", 1), ("-3", 1), ("abc", 8)):
            with patch.dict("os.environ", {"DOCX_WORKERS": raw}):
                self.assertEqual(app._read_workers(), expected, raw)

    def test_workers_are_not_a_request_parameter_or_a_form_field(self):
        html = app.app.test_client().get("/").get_data(as_text=True)
        self.assertNotIn('name="workers"', html)


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
        chapters_mode = patch.dict("os.environ", {"DOCX_CHAPTERS_MODE": "daftar"})
        chapters_mode.start()
        self.addCleanup(chapters_mode.stop)
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
