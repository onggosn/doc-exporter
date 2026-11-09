import unittest
from unittest.mock import patch

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


if __name__ == "__main__":
    unittest.main()
