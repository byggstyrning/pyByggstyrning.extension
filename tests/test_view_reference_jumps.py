# -*- coding: utf-8 -*-
"""Pure-Python tests for the 3D View Reference sheet chip."""

import importlib.util
import os
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODULE_PATH = os.path.join(ROOT, "lib", "revit", "view_reference_jumps.py")

spec = importlib.util.spec_from_file_location("view_reference_jumps", MODULE_PATH)
jumps = importlib.util.module_from_spec(spec)
spec.loader.exec_module(jumps)


class SheetChipLabelTests(unittest.TestCase):
    def test_one_sheet_number_is_a_chip(self):
        self.assertEqual(jumps.chip_label_for_sheet_numbers(["A-101"]), "Sheet A-101")

    def test_repeated_number_is_still_one_sheet(self):
        self.assertEqual(
            jumps.chip_label_for_sheet_numbers(["A-101", "A-101"]),
            "Sheet A-101",
        )

    def test_mixed_sheets_hide_the_chip(self):
        self.assertIsNone(jumps.chip_label_for_sheet_numbers(["A-101", "B-200"]))

    def test_missing_sheet_hides_the_chip(self):
        self.assertIsNone(jumps.chip_label_for_sheet_numbers(["-", "", None]))
        self.assertIsNone(jumps.chip_label_for_sheet_numbers([]))


class ViewportZoomRectTests(unittest.TestCase):
    def test_one_viewport_is_padded(self):
        self.assertEqual(
            jumps.zoom_rect_from_outlines([(0.0, 0.0, 10.0, 5.0)]),
            (-0.8, -0.4, 10.8, 5.4),
        )

    def test_several_viewports_on_one_sheet_share_one_frame(self):
        rect = jumps.zoom_rect_from_outlines([
            (0.0, 0.0, 2.0, 1.0),
            (4.0, 1.0, 6.0, 3.0),
        ])
        self.assertAlmostEqual(rect[0], -0.48)
        self.assertAlmostEqual(rect[1], -0.24)
        self.assertAlmostEqual(rect[2], 6.48)
        self.assertAlmostEqual(rect[3], 3.24)

    def test_empty_outlines_do_not_zoom(self):
        self.assertIsNone(jumps.zoom_rect_from_outlines([]))
        self.assertIsNone(jumps.padded_rectangle(1.0, 1.0, 1.0, 2.0))
