# -*- coding: utf-8 -*-
"""Pure-Python tests for MMI monitor warning thresholds."""

from __future__ import print_function

import importlib.util
import os
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
THRESHOLD_PATH = os.path.join(ROOT, "lib", "mmi", "threshold.py")

spec = importlib.util.spec_from_file_location("mmi_threshold", THRESHOLD_PATH)
threshold = importlib.util.module_from_spec(spec)
spec.loader.exec_module(threshold)


class MonitorThresholdTests(unittest.TestCase):
    def test_type_warning_when_enabled_and_over_limit(self):
        self.assertTrue(threshold.should_warn_on_count(True, 376, 375))

    def test_type_warning_not_at_exact_limit(self):
        self.assertFalse(threshold.should_warn_on_count(True, 375, 375))

    def test_warning_stays_off_when_disabled(self):
        self.assertFalse(threshold.should_warn_on_count(False, 1000, 375))

    def test_instance_param_edit_is_counted(self):
        self.assertTrue(threshold.include_in_instance_param_count(False, False, False))

    def test_type_regeneration_is_not_an_instance_param_edit(self):
        self.assertFalse(threshold.include_in_instance_param_count(False, False, True))

    def test_type_change_and_move_are_not_instance_param_edits(self):
        self.assertFalse(threshold.include_in_instance_param_count(True, False, False))
        self.assertFalse(threshold.include_in_instance_param_count(False, True, False))

    def test_pin_and_move_are_at_or_above(self):
        self.assertTrue(threshold.at_or_above(400, 400))
        self.assertFalse(threshold.at_or_above(399, 400))
        self.assertTrue(threshold.at_or_above(425, 425))
        self.assertFalse(threshold.at_or_above(424, 425))

    def test_parse_limit_accepts_defaults_and_rejects_bad_values(self):
        self.assertEqual(threshold.parse_limit("400"), 400)
        self.assertEqual(threshold.parse_limit("425"), 425)
        self.assertEqual(threshold.parse_limit("375"), 375)
        self.assertIsNone(threshold.parse_limit("0"))
        self.assertIsNone(threshold.parse_limit("-5"))
        self.assertIsNone(threshold.parse_limit("abc"))
        self.assertIsNone(threshold.parse_limit(""))
        self.assertEqual(threshold.normalize_limit("nope", 400), 400)
        self.assertEqual(threshold.normalize_limit(None, 425), 425)


if __name__ == "__main__":
    unittest.main()
