import unittest
from unittest.mock import patch
from unittest.mock import Mock

from botocore.exceptions import ClientError

import threading
import time

from bedrock_service import BedrockDocumentGenerator, BedrockError, run_parallel


class BedrockDocumentGeneratorTests(unittest.TestCase):
    def setUp(self):
        self.generator = object.__new__(BedrockDocumentGenerator)
        self.generator.knowledge_base_client = None
        self.generator.knowledge_base_id = None

    def test_requires_region_when_not_configured(self):
        session = Mock(region_name=None)
        with (
            patch.dict("os.environ", {"AWS_REGION": "", "AWS_DEFAULT_REGION": ""}),
            patch("bedrock_service.boto3.Session", return_value=session),
        ):
            with self.assertRaisesRegex(BedrockError, "Region AWS belum ditentukan"):
                BedrockDocumentGenerator()

        session.client.assert_not_called()

    def test_uses_region_from_aws_profile(self):
        session = Mock(region_name="eu-west-1")
        with (
            patch.dict("os.environ", {"AWS_REGION": "", "AWS_DEFAULT_REGION": ""}),
            patch("bedrock_service.boto3.Session", return_value=session),
        ):
            generator = BedrockDocumentGenerator()

        self.assertEqual(generator.region, "eu-west-1")
        session.client.assert_called_once_with("bedrock-runtime", region_name="eu-west-1")

    def test_sets_session_region_for_login_credential_refresh(self):
        for region, environment, expected in (
            ("ap-southeast-2", {"AWS_REGION": "us-east-1"}, "ap-southeast-2"),
            (None, {"AWS_REGION": "ap-southeast-2"}, "ap-southeast-2"),
            (None, {"AWS_DEFAULT_REGION": "ap-southeast-2"}, "ap-southeast-2"),
        ):
            with self.subTest(region=region, environment=environment):
                session = Mock(region_name=None)
                with (
                    patch.dict("os.environ", environment, clear=True),
                    patch("bedrock_service.boto3.Session", return_value=session) as factory,
                ):
                    generator = BedrockDocumentGenerator(
                        region=region, knowledge_base_id="kb-test"
                    )

                factory.assert_called_once_with(region_name=expected)
                self.assertEqual(generator.region, expected)
                self.assertEqual(session.client.call_args_list[0].kwargs["region_name"], expected)
                self.assertEqual(session.client.call_args_list[1].kwargs["region_name"], expected)

    def test_includes_region_in_bedrock_request_error(self):
        self.generator.model_id = "test-model"
        self.generator.region = "us-east-1"
        self.generator.client = Mock()
        self.generator.client.converse.side_effect = ClientError(
            {"Error": {"Code": "ValidationException", "Message": "request failed"}},
            "Converse",
        )

        with self.assertRaisesRegex(BedrockError, r"Region: us-east-1.*request failed"):
            self.generator._converse([{"text": "hello"}])

    def test_generates_chapter_html_and_removes_unsafe_markup(self):
        self.generator._converse = Mock(
            side_effect=[
                '{"Bab 1 - Pendahuluan": ["slide-1"], "Bab 2 - Penutup": []}',
                '<h2>Bab 1 - Pendahuluan</h2><p>Ringkasan.</p><script>alert(1)</script>',
            ]
        )
        deck = {
            "filename": "presentasi.pptx",
            "items": [
                {
                    "ref": "slide-1",
                    "title": "Tujuan",
                    "text": ["Meningkatkan mutu layanan"],
                    "tables": [],
                    "charts": [],
                    "notes": "",
                }
            ],
        }

        result = self.generator.generate(
            deck, ["Bab 1 - Pendahuluan", "Bab 2 - Penutup"], []
        )

        self.assertEqual(result["filename"], "presentasi.pptx")
        self.assertEqual(len(result["chapters"]), 2)
        self.assertEqual(result["chapters"][0]["slide_refs"], ["slide-1"])
        self.assertEqual(result["chapters"][1]["html"], "")
        self.assertIn("<p>Ringkasan.</p>", result["html"])
        self.assertNotIn("<script>", result["html"])

    def test_rejects_incomplete_or_duplicate_slide_mapping(self):
        self.generator._converse = Mock(
            return_value='{"Bab 1": ["slide-1"], "Bab 2": ["slide-1"]}'
        )
        deck = {
            "filename": "presentasi.pptx",
            "items": [
                {
                    "ref": "slide-1",
                    "title": "",
                    "text": ["Konten"],
                    "tables": [],
                    "charts": [],
                    "notes": "",
                }
            ],
        }

        with self.assertRaisesRegex(BedrockError, "tidak lengkap|duplikat"):
            self.generator.generate(deck, ["Bab 1", "Bab 2"], [])

    def test_retrieves_relevant_knowledge_base_examples(self):
        client = Mock()
        client.retrieve.return_value = {
            "retrievalResults": [
                {"content": {"text": "Contoh struktur laporan yang relevan"}}
            ]
        }
        self.generator.knowledge_base_client = client
        self.generator.knowledge_base_id = "kb-test"
        self.generator._converse = Mock(
            side_effect=[
                '{"Bab 1": ["slide-1"]}',
                "<h2>Bab 1</h2><p>Hasil.</p>",
            ]
        )
        deck = {
            "filename": "presentasi.pptx",
            "items": [
                {
                    "ref": "slide-1",
                    "title": "Tujuan",
                    "text": ["Meningkatkan mutu layanan"],
                    "tables": [],
                    "charts": [],
                    "notes": "",
                }
            ],
        }

        self.generator.generate(deck, ["Bab 1"], [])

        client.retrieve.assert_called_once()
        prompt = self.generator._converse.call_args_list[1].args[0][0]["text"]
        self.assertIn("Contoh struktur laporan yang relevan", prompt)


class ParseJsonTests(unittest.TestCase):
    parse = staticmethod(BedrockDocumentGenerator._parse_json)

    def test_plain_and_fenced_json(self):
        self.assertEqual(self.parse('{"a": ["x"]}'), {"a": ["x"]})
        self.assertEqual(self.parse('```json\n{"a": ["x"]}\n```'), {"a": ["x"]})

    def test_model_that_corrects_itself_uses_the_last_complete_object(self):
        text = '```json\n{"a": ["salah"], "lampiran": []}\n```\n\nSaya perlu memperbaiki:\n```json\n{"a": ["benar"]}\n```'
        self.assertEqual(self.parse(text), {"a": ["benar"]})

    def test_truncated_or_non_object_answers_still_fail(self):
        for text in ('{"a": ["terpotong', "bukan json", "[1, 2]", ""):
            with self.assertRaises(BedrockError, msg=text):
                self.parse(text)


class RunParallelTests(unittest.TestCase):
    def test_results_keep_item_order_regardless_of_finish_order(self):
        delays = {1: 0.06, 2: 0.03, 3: 0.0}

        def worker(n):
            time.sleep(delays[n])
            return n * 10

        done = []
        outcomes = run_parallel([1, 2, 3], worker, workers=3, on_done=lambda i, o: done.append(i))
        self.assertEqual(outcomes, [(True, 10), (True, 20), (True, 30)])
        self.assertEqual(done, [2, 1, 0])  # selesai terbalik, tetap terlapor satu per satu

    def test_items_really_run_at_the_same_time(self):
        barrier = threading.Barrier(3, timeout=2)  # lolos hanya bila tiga thread berjalan bersamaan
        self.assertEqual([o[1] for o in run_parallel([1, 2, 3], lambda n: barrier.wait() >= 0 or n, workers=3)],
                         [True, True, True])

    def test_sequential_when_one_worker(self):
        order = []
        run_parallel([1, 2, 3], order.append, workers=1)
        self.assertEqual(order, [1, 2, 3])

    def test_stop_on_error_raises_first_error_and_cancels_pending(self):
        started = []

        def worker(n):
            started.append(n)
            if n == 1:
                raise BedrockError("gagal")
            return n

        with self.assertRaisesRegex(BedrockError, "gagal"):
            run_parallel(list(range(1, 30)), worker, workers=2)
        self.assertLess(len(started), 29)

    def test_errors_are_returned_when_not_stopping(self):
        def worker(n):
            if n == 2:
                raise ValueError("dua")
            return n

        outcomes = run_parallel([1, 2, 3], worker, workers=2, stop_on_error=False)
        self.assertEqual([o[0] for o in outcomes], [True, False, True])
        self.assertIsInstance(outcomes[1][1], ValueError)

    def test_workers_are_capped_and_empty_input_is_fine(self):
        self.assertEqual(run_parallel([], len, workers=99), [])
        self.assertEqual(run_parallel([1], lambda n: n, workers=0), [(True, 1)])


class ImageSummaryTests(unittest.TestCase):
    SLIDE = {"ref": "slide-5", "title": "Kredit", "text": ["Kredit tumbuh 8,88% (yoy)."], "tables": [], "charts": [],
             "notes": "", "visual_desc": None}

    def make(self, replies):
        generator = object.__new__(BedrockDocumentGenerator)
        generator.usage_totals = {}
        patcher = patch.object(BedrockDocumentGenerator, "_converse", side_effect=replies)
        self.converse = patcher.start()
        self.addCleanup(patcher.stop)
        return generator

    def test_clean_summary_is_returned_after_one_call_with_image_and_context(self):
        generator = self.make(["**Grafik** pertumbuhan kredit 8,88%."])
        summary, flagged = generator.summarize_image(b"png", self.SLIDE, "PERTUMBUHAN KREDIT")
        self.assertEqual((summary, flagged, self.converse.call_count), ("Grafik pertumbuhan kredit 8,88%.", [], 1))
        content = self.converse.call_args.args[0]
        self.assertIn("Judul panel gambar: PERTUMBUHAN KREDIT", content[0]["text"])
        self.assertEqual(content[1]["image"]["source"]["bytes"], b"png")

    def test_number_not_in_slide_text_is_retried_once_then_reported(self):
        generator = self.make(["Kredit tumbuh 9,5%.", "Kredit tumbuh 9,5%."])
        summary, flagged = generator.summarize_image(b"png", self.SLIDE, None)
        self.assertEqual((summary, flagged, self.converse.call_count), ("Kredit tumbuh 9,5%.", ["9.5"], 2))
        self.assertIn("9.5", self.converse.call_args.args[0][0]["text"])

    def test_retry_that_drops_the_number_is_accepted(self):
        generator = self.make(["Kredit tumbuh 9,5%.", "Kredit tumbuh."])
        self.assertEqual(generator.summarize_image(b"png", self.SLIDE, None), ("Kredit tumbuh.", []))

    def make_deck(self, tmp):
        from pathlib import Path

        def png(name):
            path = Path(tmp) / name
            path.write_bytes(b"png")
            return str(path)

        split = {**self.SLIDE, "ref": "s1", "image_path": png("w1.png"),
                 "panel_images": [{"header": "A", "image_path": png("a.png")}, {"header": "B", "image_path": png("b.png")}]}
        whole = {**self.SLIDE, "ref": "s2", "image_path": png("w2.png")}
        text_slide = {**self.SLIDE, "ref": "s3"}
        return {"items": [split, whole, text_slide]}

    def generator(self):
        generator = object.__new__(BedrockDocumentGenerator)
        generator.usage_totals, generator.workers = {}, 1
        return generator

    def test_describe_visuals_can_skip_slides_that_are_split_into_panels(self):
        import tempfile
        for describe_split, calls in ((True, 2), (False, 1)):
            with tempfile.TemporaryDirectory() as tmp:
                generator = self.generator()
                generator.describe_split_slides = describe_split
                deck = self.make_deck(tmp)
                with patch.object(BedrockDocumentGenerator, "_converse", return_value="Deskripsi.") as converse:
                    generator.describe_visuals(deck)
                self.assertEqual(converse.call_count, calls, describe_split)
                self.assertEqual(deck["items"][0].get("visual_desc"), "Deskripsi." if describe_split else None)

    def test_summarize_images_covers_each_panel_or_whole_slide_and_honours_skip_refs(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            generator = self.generator()
            deck = self.make_deck(tmp)
            progress = []
            with patch.object(BedrockDocumentGenerator, "_converse", return_value="Ringkasan tanpa angka.") as converse:
                generator.summarize_images(deck, lambda *a: progress.append(a))
            split, whole, text_slide = deck["items"]
            self.assertEqual(converse.call_count, 3)  # panel A, panel B, slide utuh; slide teks tidak punya gambar
            self.assertEqual([p["summary"] for p in split["panel_images"]], ["Ringkasan tanpa angka."] * 2)
            self.assertEqual((whole["summary"], "summary" in text_slide), ("Ringkasan tanpa angka.", False))
            self.assertEqual(progress[-1][:2], ("captions", "done"))

            deck = self.make_deck(tmp)
            with patch.object(BedrockDocumentGenerator, "_converse", return_value="Ringkasan.") as converse:
                generator.summarize_images(deck, None, skip_refs={"s1"})
            self.assertEqual(converse.call_count, 1)
            self.assertNotIn("summary", deck["items"][0]["panel_images"][0])


class ParallelStepsTests(unittest.TestCase):
    def make(self, workers):
        generator = object.__new__(BedrockDocumentGenerator)
        generator.workers = workers
        generator.usage_totals = {}
        return generator

    def test_describe_visuals_runs_in_parallel_and_counts_tokens_exactly(self):
        import tempfile
        from pathlib import Path
        generator = self.make(4)
        with tempfile.TemporaryDirectory() as tmp:
            items = []
            for i in range(12):
                path = Path(tmp) / f"s{i}.png"
                path.write_bytes(b"png")
                items.append({"ref": f"slide-{i}", "title": "", "text": [], "image_path": str(path)})

            def fake_post(content, max_tokens=6000, step=None):
                time.sleep(0.01)
                generator._track_usage(step, {"inputTokens": 7, "outputTokens": 3})
                return "deskripsi"

            progress = []
            with patch.object(BedrockDocumentGenerator, "_converse", side_effect=fake_post):
                generator.describe_visuals({"items": items}, lambda *a: progress.append(a))
        self.assertTrue(all(item["visual_desc"] == "deskripsi" for item in items))
        self.assertEqual(generator.usage_totals["visuals"], {"input_tokens": 84, "output_tokens": 36})
        self.assertEqual(progress[-1][:2], ("visuals", "done"))
        self.assertEqual([p[3] for p in progress if p[1] == "running"], list(range(0, 13)))

    def test_per_slide_mapping_keeps_deck_order_inside_each_chapter(self):
        generator = self.make(4)
        deck = {"items": [{"ref": f"s{i}"} for i in range(1, 9)]}

        def classify(slide, chapters):
            time.sleep(0.05 - int(slide["ref"][1:]) * 0.005)  # slide akhir selesai lebih dulu
            return chapters[int(slide["ref"][1:]) % 2]

        with patch.object(BedrockDocumentGenerator, "classify_slide", side_effect=classify):
            mapping = generator._map_slides_per_slide(deck, ["A", "B"], None)
        self.assertEqual(mapping, {"A": ["s2", "s4", "s6", "s8"], "B": ["s1", "s3", "s5", "s7"]})


if __name__ == "__main__":
    unittest.main()
