# -*- coding: utf-8 -*-
"""Pure-Python tests for 3D View Reference name formulas."""

from __future__ import print_function

import importlib.util
import os
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FORMULA_PATH = os.path.join(ROOT, "lib", "revit", "view_reference_formula.py")

spec = importlib.util.spec_from_file_location("view_reference_formula", FORMULA_PATH)
formula = importlib.util.module_from_spec(spec)
spec.loader.exec_module(formula)


class ViewReferenceFormulaTests(unittest.TestCase):
    def test_default_view_formula_is_the_view_name(self):
        normalized = formula.normalize_view_name_formula(None)
        self.assertEqual(
            normalized["parts"],
            [{"source": formula.SOURCE_VIEW_NAME}],
        )

    def test_empty_formula_becomes_one_default_part(self):
        normalized = formula.normalize_formula({}, formula.SOURCE_VIEW_NAME)
        self.assertEqual(len(normalized["parts"]), 1)
        self.assertEqual(normalized["parts"][0]["source"], formula.SOURCE_VIEW_NAME)

    def test_shared_separator_is_copied_onto_each_gap(self):
        normalized = formula.normalize_view_name_formula({
            "separator": " - ",
            "parts": [
                {"source": formula.SOURCE_VIEW, "name": "Phase"},
                {"source": formula.SOURCE_VIEW_NAME},
            ],
        })
        self.assertEqual(normalized["parts"][0]["separator"], " - ")
        self.assertNotIn("separator", normalized["parts"][-1])

    def test_compose_joins_separators_and_skips_empty_parts(self):
        text = formula.compose_parts(
            {
                "parts": [
                    {"source": formula.SOURCE_VIEW, "name": "Missing", "separator": " / "},
                    {"source": formula.SOURCE_VIEW_NAME, "separator": "-"},
                    {"source": formula.SOURCE_SHEET_NUMBER},
                ],
            },
            formula.SOURCE_VIEW_NAME,
            lambda part: {
                formula.SOURCE_VIEW_NAME: "Section A",
                formula.SOURCE_SHEET_NUMBER: "A101",
            }.get(part.get("source"), ""),
        )
        self.assertEqual(text, "Section A-A101")

    def test_view_name_label(self):
        self.assertEqual(
            formula.part_label({"source": formula.SOURCE_VIEW_NAME}),
            "View name",
        )
        self.assertEqual(
            formula.part_label({"source": formula.SOURCE_VIEW, "name": "Phase"}),
            "View: Phase",
        )


if __name__ == "__main__":
    unittest.main()
