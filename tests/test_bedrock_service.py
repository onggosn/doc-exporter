import unittest
from unittest.mock import patch
from unittest.mock import Mock

from botocore.exceptions import ClientError

import threading
import time

from bedrock_service import BEDROCK_CLIENT_CONFIG, BedrockDocumentGenerator, BedrockError, run_parallel


class BedrockDocumentGeneratorTests(unittest.TestCase):
    def setUp(self):
        self.generator = object.__new__(BedrockDocumentGenerator)

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
        session.client.assert_called_once_with(
            "bedrock-runtime", region_name="eu-west-1", config=BEDROCK_CLIENT_CONFIG
        )

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
                    generator = BedrockDocumentGenerator(region=region)

                factory.assert_called_once_with(region_name=expected)
                self.assertEqual(generator.region, expected)
                self.assertEqual(session.client.call_args_list[0].kwargs["region_name"], expected)

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


class MappingTextTests(unittest.TestCase):
    SLIDE = {
        "ref": "slide-7", "title": "Inflasi", "notes": "Catatan " * 100,  # notes lama tidak boleh ikut terkirim
         "visual_desc": "Grafik garis inflasi.",
        "text": ["Inflasi terkendali " * 60, "baris kedua"],
        "tables": [[["Komoditas", "Mei", "Juni"], ["Beras", "123456", "654321"]]],
        "charts": [{"title": "Laju inflasi", "categories": ["Jan", "Feb"],
                    "series": [{"name": "IHK", "values": [1.1, 2.2]}]}],
    }

    def text(self, **env):
        with patch.dict("os.environ", env):
            return BedrockDocumentGenerator._mapping_text(self.SLIDE)

    def test_keeps_topic_signals_but_drops_numeric_table_and_chart_data(self):
        text = self.text()
        for expected in ("slide-7", "Inflasi", "Chart: Laju inflasi; IHK", "Header tabel: Komoditas Mei Juni",
                         "Deskripsi visual: Grafik garis inflasi."):
            self.assertIn(expected, text)
        for dropped in ("Beras", "654321", "1.1", "Jan", "Catatan", "Speaker"):
            self.assertNotIn(dropped, text)

    def test_text_is_cut_at_the_limit_and_much_shorter_than_before(self):
        text = self.text(MAPPING_TEXT_MAX_CHARS="100")
        body = next(line for line in text.splitlines() if line.startswith("Teks: "))
        self.assertLessEqual(len(body), len("Teks: ") + 101)
        self.assertTrue(body.endswith("…"))
        self.assertLess(len(self.text(MAPPING_TEXT_MAX_CHARS="400")), 900)

    def test_zero_means_no_cut(self):
        body = next(line for line in self.text(MAPPING_TEXT_MAX_CHARS="0").splitlines() if line.startswith("Teks: "))
        self.assertIn("baris kedua", body)

    def test_batch_and_per_slide_prompts_use_the_compact_text(self):
        generator = object.__new__(BedrockDocumentGenerator)
        generator.usage_totals, generator.workers = {}, 1
        deck = {"items": [self.SLIDE]}
        with patch.object(BedrockDocumentGenerator, "_converse", return_value='{"Bab A": ["slide-7"]}') as converse:
            generator._map_slides_batch(deck, ["Bab A"])
        self.assertIn("Header tabel: Komoditas", converse.call_args.args[0][0]["text"])
        self.assertNotIn("654321", converse.call_args.args[0][0]["text"])
        with patch.object(BedrockDocumentGenerator, "_converse", return_value="Bab A") as converse:
            generator.classify_slide(self.SLIDE, ["Bab A"])
        self.assertNotIn("654321", converse.call_args.args[0][0]["text"])


class BatchMappingRepairTests(unittest.TestCase):
    CHAPTERS = ["Moneter", "Makro", "Sistem Pembayaran"]
    DECK = {"items": [{"ref": f"slide-{n}", "title": f"Judul {n}", "text": [f"isi {n}"]} for n in range(1, 7)]}

    def generator(self):
        generator = object.__new__(BedrockDocumentGenerator)
        generator.usage_totals, generator.workers = {}, 1
        return generator

    def run_mapping(self, answers, classify=None):
        """answers: balasan model untuk pemetaan sekaligus; classify: slide atau (slide, kandidat) -> bab."""
        classified = []

        def fake_classify(_self, slide, chapters):
            classified.append((slide["ref"], list(chapters)))
            return classify[slide["ref"]]

        with patch.object(BedrockDocumentGenerator, "_converse", side_effect=answers) as converse, \
                patch.object(BedrockDocumentGenerator, "classify_slide", fake_classify):
            mapping = self.generator()._map_slides_batch(self.DECK, self.CHAPTERS)
        return mapping, converse, classified

    def test_complete_answer_needs_exactly_one_call(self):
        answer = '{"Moneter": ["slide-1", "slide-2"], "Makro": ["slide-3", "slide-4"], "Sistem Pembayaran": ["slide-5", "slide-6"]}'
        mapping, converse, classified = self.run_mapping([answer])
        self.assertEqual((converse.call_count, classified), (1, []))
        self.assertEqual(mapping["Makro"], ["slide-3", "slide-4"])

    def test_missing_slides_are_classified_alone_with_all_chapters_instead_of_resending_the_deck(self):
        answer = '{"Moneter": ["slide-1", "slide-2"], "Makro": ["slide-3"], "Sistem Pembayaran": ["slide-6"]}'
        mapping, converse, classified = self.run_mapping([answer], {"slide-4": "Makro", "slide-5": "Moneter"})
        self.assertEqual(converse.call_count, 1)  # tidak ada retry penuh
        self.assertEqual(classified, [("slide-4", self.CHAPTERS), ("slide-5", self.CHAPTERS)])
        self.assertEqual(mapping, {"Moneter": ["slide-1", "slide-2", "slide-5"], "Makro": ["slide-3", "slide-4"],
                                   "Sistem Pembayaran": ["slide-6"]})  # urutan bab mengikuti urutan deck

    def test_duplicated_slide_is_decided_among_only_its_candidate_chapters(self):
        answer = ('{"Moneter": ["slide-1", "slide-2", "slide-3"], "Makro": ["slide-3", "slide-4"], '
                  '"Sistem Pembayaran": ["slide-5", "slide-6"]}')
        mapping, _converse, classified = self.run_mapping([answer], {"slide-3": "Makro"})
        self.assertEqual(classified, [("slide-3", ["Moneter", "Makro"])])
        self.assertEqual(mapping["Moneter"], ["slide-1", "slide-2"])
        self.assertEqual(mapping["Makro"], ["slide-3", "slide-4"])
        total = sum(len(refs) for refs in mapping.values())
        self.assertEqual(total, 6)  # tiap slide tepat satu kali

    def test_duplicate_within_same_chapter_and_unknown_refs_need_no_extra_call(self):
        answer = ('{"Moneter": ["slide-1", "slide-1", "slide-99"], "Makro": ["slide-2", "slide-3", "slide-4"], '
                  '"Sistem Pembayaran": ["slide-5", "slide-6"]}')
        mapping, converse, classified = self.run_mapping([answer])
        self.assertEqual((converse.call_count, classified), (1, []))
        self.assertEqual(mapping["Moneter"], ["slide-1"])

    def test_invalid_json_is_the_only_case_that_still_retries_the_full_prompt(self):
        good = '{"Moneter": ["slide-1", "slide-2", "slide-3", "slide-4", "slide-5", "slide-6"]}'
        mapping, converse, _ = self.run_mapping(["bukan json", good])
        self.assertEqual(converse.call_count, 2)
        self.assertEqual(len(mapping["Moneter"]), 6)
        with self.assertRaisesRegex(BedrockError, "format yang valid"):
            self.run_mapping(["x", "y", "z"])


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

    def long_slide(self):
        filler = [f"Baris keterangan nomor {i} yang cukup panjang untuk mengisi batas." for i in range(60)]
        return {**self.SLIDE, "text": filler + ["Kredit tumbuh 8,88% (yoy)."]}  # angka ada di ujung, setelah terpotong

    def test_prompt_gets_trimmed_facts_but_numbers_are_still_checked_against_full_facts(self):
        generator = self.make(["Kredit tumbuh 8,88%."])
        with patch.dict("os.environ", {"CAPTION_FACTS_MAX_CHARS": "300"}):
            summary, flagged = generator.summarize_image(b"png", self.long_slide(), None)
        prompt = self.converse.call_args.args[0][0]["text"]
        self.assertNotIn("8,88%", prompt)               # bagian ujung tidak terkirim ke model
        self.assertIn("teks slide dipotong", prompt)
        self.assertLess(len(prompt), 1500)
        self.assertEqual((flagged, self.converse.call_count), ([], 1))  # tapi angka itu tetap dianggap sah

    def test_zero_limit_sends_all_facts(self):
        generator = self.make(["Kredit tumbuh 8,88%."])
        with patch.dict("os.environ", {"CAPTION_FACTS_MAX_CHARS": "0"}):
            generator.summarize_image(b"png", self.long_slide(), None)
        self.assertIn("8,88%", self.converse.call_args.args[0][0]["text"])

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

    PANEL_JSON = '{"1": "Ringkasan panel A.", "2": "Ringkasan panel B."}'

    def test_summarize_images_one_call_per_slide_and_honours_skip_refs(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            generator = self.generator()
            deck = self.make_deck(tmp)
            progress = []
            replies = [self.PANEL_JSON, "Ringkasan slide utuh."]
            with patch.object(BedrockDocumentGenerator, "_converse", side_effect=replies) as converse:
                generator.summarize_images(deck, lambda *a: progress.append(a))
            split, whole, text_slide = deck["items"]
            self.assertEqual(converse.call_count, 2)  # satu untuk kedua panel, satu untuk slide utuh
            self.assertEqual([p["summary"] for p in split["panel_images"]], ["Ringkasan panel A.", "Ringkasan panel B."])
            self.assertEqual(split["panel_images"][0]["summary_flagged"], [])
            self.assertEqual((whole["summary"], "summary" in text_slide), ("Ringkasan slide utuh.", False))
            self.assertEqual(progress[-1], ("captions", "done", "", 3, 3))  # progres dihitung per gambar
            content = converse.call_args_list[0].args[0]
            self.assertEqual(sum(1 for block in content if "image" in block), 2)  # dua gambar, satu teks slide
            self.assertEqual(content[0]["text"].count("Kredit tumbuh 8,88%"), 1)

            deck = self.make_deck(tmp)
            with patch.object(BedrockDocumentGenerator, "_converse", return_value="Ringkasan.") as converse:
                generator.summarize_images(deck, None, skip_refs={"s1"})
            self.assertEqual(converse.call_count, 1)
            self.assertNotIn("summary", deck["items"][0]["panel_images"][0])

    def test_panel_with_unsupported_number_is_redone_alone_while_good_panel_is_kept(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            deck = self.make_deck(tmp)
            split = deck["items"][0]
            replies = ['{"1": "Panel A baik.", "2": "Kredit tumbuh 9,5%."}', "Panel B diperbaiki."]
            with patch.object(BedrockDocumentGenerator, "_converse", side_effect=replies) as converse:
                self.generator().summarize_panels(split, split["panel_images"])
            self.assertEqual(converse.call_count, 2)  # gabungan + ulang hanya panel B
            self.assertEqual([p["summary"] for p in split["panel_images"]], ["Panel A baik.", "Panel B diperbaiki."])
            self.assertEqual(split["panel_images"][1]["summary_flagged"], [])
            single = converse.call_args_list[1].args[0]
            self.assertEqual(sum(1 for block in single if "image" in block), 1)

    def test_invalid_combined_answer_falls_back_to_one_call_per_panel(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            deck = self.make_deck(tmp)
            split = deck["items"][0]
            replies = ["bukan json", "Panel A.", "Panel B."]
            with patch.object(BedrockDocumentGenerator, "_converse", side_effect=replies) as converse:
                self.generator().summarize_panels(split, split["panel_images"])
            self.assertEqual(converse.call_count, 3)
            self.assertEqual([p["summary"] for p in split["panel_images"]], ["Panel A.", "Panel B."])

    TABLE_FACTS = {"text": [], "charts": [], "tables": [
        [["Asumsi"], ["", "Baseline", "PR"], ["Nilai Tukar", "16.880", "17.260"], ["Yield SBN (%)", "6,7%", "7,3%"]]]}

    def test_empty_model_summary_falls_back_to_facts_of_the_panel_table(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            deck = self.make_deck(tmp)
            split = deck["items"][0]
            split["panel_images"][0]["facts"] = self.TABLE_FACTS
            replies = ['{"1": "", "2": "Panel B."}', "", ""]  # gabungan kosong untuk panel A, ulang per panel juga kosong
            with patch.object(BedrockDocumentGenerator, "_converse", side_effect=replies):
                self.generator().summarize_panels(split, split["panel_images"])
            summary = split["panel_images"][0]["summary"]
            self.assertIn("Nilai Tukar", summary)
            self.assertIn("Baseline", summary)
            self.assertEqual(split["panel_images"][0]["summary_flagged"], [])

    def test_panel_facts_are_sent_with_the_panel_in_the_combined_prompt(self):
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            deck = self.make_deck(tmp)
            split = deck["items"][0]
            split["panel_images"][1]["facts"] = self.TABLE_FACTS
            with patch.object(BedrockDocumentGenerator, "_converse", return_value=self.PANEL_JSON) as converse:
                self.generator().summarize_panels(split, split["panel_images"])
            prompt = converse.call_args.args[0][0]["text"]
            self.assertIn("Isi panel menurut file slide: Asumsi", prompt)
            self.assertIn("Nilai Tukar | 16.880 | 17.260", prompt)


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
