import time
import unittest
from unittest.mock import MagicMock, patch

import app
from bedrock_service import BedrockError


class FeatureLoggingTests(unittest.TestCase):
    def setUp(self):
        patcher = patch.object(app, "_log_processing_times")  # jangan menulis ke PostgreSQL sungguhan
        self.log = patcher.start()
        self.addCleanup(patcher.stop)
        self.client = app.app.test_client()

    def fake_generator(self, **attributes):
        generator = MagicMock()
        generator.knowledge_base_client = object()
        generator.knowledge_base_id = "KB1"
        generator.usage_totals = {}
        for key, value in attributes.items():
            setattr(generator, key, value)
        return generator

    def post_kb(self, generator, html="<p>Kredit tumbuh</p>"):
        with patch.object(app, "BedrockDocumentGenerator", return_value=generator):
            return self.client.post("/api/correct-from-kb", data={"html_fragment": html, "knowledge_base_id": "KB1"})

    def test_kb_correction_logs_duration_and_tokens_in_the_kb_columns(self):
        generator = self.fake_generator(usage_totals={"kb": {"input_tokens": 900, "output_tokens": 120}})
        generator.retrieve_kb_materials.return_value = ["materi"]
        generator.correct_fragment_from_kb.return_value = {"html": "<p>x</p>", "changes": []}
        self.assertEqual(self.post_kb(generator).status_code, 200)
        name, status, timing, total_seconds, usage = self.log.call_args.args
        self.assertEqual((name, status), ("[KB Correction]", "done"))
        self.assertIsNotNone(timing["kb"]["seconds"])
        self.assertIsNone(timing["writing"]["seconds"])
        self.assertEqual(usage, {"kb": {"input_tokens": 900, "output_tokens": 120}})
        self.assertGreaterEqual(total_seconds, 0)

    def test_kb_correction_failure_is_logged_as_error_but_validation_errors_are_not(self):
        generator = self.fake_generator(usage_totals={"kb": {"input_tokens": 50, "output_tokens": 0}})
        generator.retrieve_kb_materials.return_value = ["materi"]
        generator.correct_fragment_from_kb.side_effect = BedrockError("gagal")
        self.assertEqual(self.post_kb(generator).status_code, 400)
        self.assertEqual(self.log.call_args.args[:2], ("[KB Correction]", "error"))
        self.assertEqual(self.log.call_args.args[4], {"kb": {"input_tokens": 50, "output_tokens": 0}})

        self.log.reset_mock()
        empty = self.fake_generator()
        empty.retrieve_kb_materials.return_value = []
        self.assertEqual(self.post_kb(empty).status_code, 400)  # tidak ada materi: tidak ada pekerjaan model
        self.assertEqual(self.client.post("/api/correct-from-kb", data={"html_fragment": ""}).status_code, 400)
        self.log.assert_not_called()

    def test_part_document_correction_is_logged_in_the_writing_columns(self):
        app.DOCUMENT_SESSIONS["doc1"] = {
            "created_at": time.time(), "references": [],
            "chapters_result": [{"title": "Bab", "slide_refs": ["slide-1"]}],
            "deck": {"items": [{"ref": "slide-1", "title": "T", "text": []}]},
            "config": {"model_id": None, "region": None, "knowledge_base_id": None},
        }
        self.addCleanup(app.DOCUMENT_SESSIONS.clear)
        generator = self.fake_generator(usage_totals={"writing": {"input_tokens": 700, "output_tokens": 80}})
        generator._slide_text.return_value = "materi"
        generator._retrieve_examples.return_value = []
        generator.correct_fragment.return_value = "<p>x</p>"
        with patch.object(app, "BedrockDocumentGenerator", return_value=generator):
            response = self.client.post("/api/correct-selection", json={
                "document_id": "doc1", "chapter": "Bab", "selected_html": "<p>a</p>"})
        self.assertEqual(response.status_code, 200)
        name, status, timing, _total, usage = self.log.call_args.args
        self.assertEqual((name, status), ("[Part Document: koreksi bagian]", "done"))
        self.assertIsNotNone(timing["writing"]["seconds"])
        self.assertEqual(usage, {"writing": {"input_tokens": 700, "output_tokens": 80}})

    def test_manual_mapping_session_is_logged_once_at_finalize_with_all_steps(self):
        generator = self.fake_generator(
            usage_totals={"visuals": {"input_tokens": 10, "output_tokens": 2}, "mapping": {"input_tokens": 30, "output_tokens": 4},
                          "writing": {"input_tokens": 500, "output_tokens": 90}})
        generator.model_id = generator.region = generator.knowledge_base_id = None
        generator.generate_from_mapping.return_value = {"chapters": []}
        app.MAPPING_SESSIONS["m1"] = {
            "document_name": "materi.pptx", "created_at": time.time() - 5, "pending": set(), "generator": generator,
            "deck": {"items": []}, "chapters": ["A"], "references": [], "mapping": {"A": []}, "render_dir": None,
            "timing": {"ingest": 1.5, "visuals": 2.5, "mapping": 3.5},
        }
        self.addCleanup(app.MAPPING_SESSIONS.clear)
        with patch.object(app, "_store_document_session", return_value="doc"):
            response = self.client.post("/api/finalize", json={"session_id": "m1"})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.log.call_count, 1)
        name, status, timing, total_seconds, usage = self.log.call_args.args
        self.assertEqual((name, status), ("materi.pptx [Mapping manual]", "done"))
        self.assertEqual([timing[k]["seconds"] for k in ("ingest", "visuals", "mapping")], [1.5, 2.5, 3.5])
        self.assertIsNotNone(timing["writing"]["seconds"])
        self.assertGreaterEqual(total_seconds, 5)
        self.assertEqual(set(usage), {"visuals", "mapping", "writing"})

    def test_finalize_with_pending_slides_logs_nothing(self):
        app.MAPPING_SESSIONS["m2"] = {"document_name": "x.pptx", "created_at": time.time(), "pending": {"slide-1"},
                                      "generator": self.fake_generator(), "render_dir": None, "timing": {}}
        self.addCleanup(app.MAPPING_SESSIONS.clear)
        self.assertEqual(self.client.post("/api/finalize", json={"session_id": "m2"}).status_code, 400)
        self.log.assert_not_called()

    def test_usage_for_log_sums_steps_that_share_a_column_and_drops_unknown_ones(self):
        usage = {"narration": {"input_tokens": 5, "output_tokens": 1}, "writing": {"input_tokens": 3, "output_tokens": 2},
                 "style_guide": {"input_tokens": 1, "output_tokens": 1}, "visuals": {"input_tokens": 7, "output_tokens": 0},
                 "lainnya": {"input_tokens": 99, "output_tokens": 99}}
        self.assertEqual(app._usage_for_log(usage), {"writing": {"input_tokens": 9, "output_tokens": 4},
                                                     "visuals": {"input_tokens": 7, "output_tokens": 0}})


if __name__ == "__main__":
    unittest.main()
