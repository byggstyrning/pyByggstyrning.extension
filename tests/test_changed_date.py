# -*- coding: utf-8 -*-
"""Pure-Python tests for the optional MMI changed date."""

from __future__ import print_function

import datetime
import importlib.util
import os
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODULE_PATH = os.path.join(ROOT, "lib", "mmi", "changed_date.py")

spec = importlib.util.spec_from_file_location("mmi_changed_date", MODULE_PATH)
changed_date = importlib.util.module_from_spec(spec)
spec.loader.exec_module(changed_date)


class ChangedDateTests(unittest.TestCase):
    def test_format_is_yyyy_mm_dd(self):
        self.assertEqual(
            changed_date.format_changed_date(datetime.date(2026, 10, 13)),
            "2026-10-13",
        )
        self.assertEqual(
            changed_date.format_changed_date(datetime.date(2026, 1, 5)),
            "2026-01-05",
        )

    def test_stamps_when_enabled_and_parameter_is_set(self):
        self.assertTrue(
            changed_date.should_stamp_changed_date(True, "MMI Datum", "MMI")
        )

    def test_stays_off_when_switch_is_off(self):
        self.assertFalse(
            changed_date.should_stamp_changed_date(False, "MMI Datum", "MMI")
        )

    def test_stays_off_without_a_parameter(self):
        self.assertFalse(changed_date.should_stamp_changed_date(True, "", "MMI"))
        self.assertFalse(changed_date.should_stamp_changed_date(True, "   ", "MMI"))
        self.assertFalse(changed_date.should_stamp_changed_date(True, None, "MMI"))

    def test_does_not_stamp_the_mmi_parameter(self):
        self.assertFalse(changed_date.should_stamp_changed_date(True, "MMI", "MMI"))
        self.assertFalse(
            changed_date.should_stamp_changed_date(True, " MMI ", "MMI")
        )


if __name__ == "__main__":
    unittest.main()
