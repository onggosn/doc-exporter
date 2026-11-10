import unittest
from unittest.mock import MagicMock, patch

import app


class FeatureLoggingTests(unittest.TestCase):
    def setUp(self):
        patcher = patch.object(app, "_log_processing_times")  # jangan menulis ke PostgreSQL sungguhan
        self.log = patcher.start()
        self.addCleanup(patcher.stop)
        self.client = app.app.test_client()

    def test_usage_for_log_sums_steps_that_share_a_column_and_drops_unknown_ones(self):
        usage = {"narration": {"input_tokens": 5, "output_tokens": 1}, "writing": {"input_tokens": 3, "output_tokens": 2},
                 "style_guide": {"input_tokens": 1, "output_tokens": 1}, "visuals": {"input_tokens": 7, "output_tokens": 0},
                 "lainnya": {"input_tokens": 99, "output_tokens": 99}}
        self.assertEqual(app._usage_for_log(usage), {"writing": {"input_tokens": 9, "output_tokens": 4},
                                                     "visuals": {"input_tokens": 7, "output_tokens": 0}})

    def test_caption_tokens_have_their_own_step(self):
        usage = {"visuals": {"input_tokens": 10, "output_tokens": 2}, "captions": {"input_tokens": 4, "output_tokens": 1}}
        self.assertEqual(app._usage_for_log(usage), usage)


class ProcessingLogInsertTests(unittest.TestCase):
    def run_log(self, enabled=True):
        timing = {step: {"seconds": 1.5, "at": "2026-01-01"} for step in app.LOGGED_STEPS}
        usage = {"captions": {"input_tokens": 4, "output_tokens": 1}, "narration": {"input_tokens": 6, "output_tokens": 2}}
        connection = MagicMock()
        cursor = connection.__enter__.return_value.cursor.return_value.__enter__.return_value
        with patch.object(app, "_LOG_PROCESSING_ENABLED", enabled), \
                patch.object(app.psycopg2, "connect", return_value=connection) as connect, \
                patch.object(app.pg_sql.Composed, "as_string", return_value="SQL", create=True):
            app._log_processing_times("x.pptx", "done", timing, 9.0, app._usage_for_log(usage))
        return connect, cursor

    def test_insert_has_one_value_per_column_including_captions(self):
        connect, cursor = self.run_log()
        query, values = cursor.execute.call_args.args
        # 2 (nama, status) + 4 kolom x 5 langkah + 3 total
        self.assertEqual(len(values), 2 + 4 * len(app.LOGGED_STEPS) + 3)
        self.assertIn("captions", app.LOGGED_STEPS)
        self.assertIn(4, values)  # token input captions ikut tersimpan
        self.assertEqual(values[-3:], [9.0, 10, 3])

    def test_nothing_is_written_when_logging_is_disabled(self):
        connect, _ = self.run_log(enabled=False)
        connect.assert_not_called()


if __name__ == "__main__":
    unittest.main()
