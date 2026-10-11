# -*- coding: utf-8 -*-
"""Pure-Python tests for nested-family sharing rollback checks."""

from __future__ import print_function

import importlib.util
import os
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CHECKS_PATH = os.path.join(ROOT, "lib", "filesize", "share_checks.py")

spec = importlib.util.spec_from_file_location("share_checks", CHECKS_PATH)
checks = importlib.util.module_from_spec(spec)
spec.loader.exec_module(checks)


class ShareCheckTests(unittest.TestCase):
    def test_same_types_and_placements_are_kept(self):
        missing, placed_lost = checks.reload_loss(
            ["A", "B"], ["B", "A"], 4, 4)
        self.assertEqual(missing, [])
        self.assertFalse(placed_lost)

    def test_extra_type_and_extra_placements_are_kept(self):
        missing, placed_lost = checks.reload_loss(["A"], ["A", "B"], 2, 5)
        self.assertEqual(missing, [])
        self.assertFalse(placed_lost)

    def test_lost_type_is_flagged(self):
        missing, placed_lost = checks.reload_loss(
            ["A", "B"], ["A"], 3, 3)
        self.assertEqual(missing, ["B"])
        self.assertFalse(placed_lost)

    def test_lost_placement_is_flagged(self):
        missing, placed_lost = checks.reload_loss(["A"], ["A"], 3, 2)
        self.assertEqual(missing, [])
        self.assertTrue(placed_lost)

    def test_category_delta_includes_categories_that_disappear(self):
        delta = checks.category_delta(
            {"Electrical Fixtures": 10, "Generic Models": 2},
            {"Electrical Fixtures": 14})
        self.assertEqual(delta["Electrical Fixtures"], 4)
        self.assertEqual(delta["Generic Models"], -2)
        self.assertNotIn("Doors", delta)


if __name__ == "__main__":
    unittest.main()
