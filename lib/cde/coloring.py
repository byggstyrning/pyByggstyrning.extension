# -*- coding: utf-8 -*-
"""Temporary view graphics coloring of elements by a CDE value.

Reuses the project's existing override helpers (``apply_color_to_elements``,
``solid_fill_pattern_id``) so coloring behaves like the ColorElements tool, but
keyed off any CDE/Revit value the schedule exposes. Visualizes data in the
active view without needing Revit parameters or tags.
"""
import colorsys

import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import Color, OverrideGraphicSettings

from pyrevit import script, revit

import sys
import os.path as op
_lib_dir = op.dirname(op.dirname(__file__))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

from revit.revit_utils import apply_color_to_elements

logger = script.get_logger()


def generate_color_map(values):
    """Return {value: DB.Color} spreading distinct hues across the values."""
    distinct = []
    for value in values:
        key = _norm(value)
        if key not in distinct:
            distinct.append(key)
    count = max(len(distinct), 1)
    color_map = {}
    for i, key in enumerate(distinct):
        hue = float(i) / count
        r, g, b = colorsys.hsv_to_rgb(hue, 0.65, 0.95)
        color_map[key] = Color(int(r * 255), int(g * 255), int(b * 255))
    return color_map


def _norm(value):
    return "" if value is None else unicode(value)


def build_legend(color_map):
    """Return [(value, (r, g, b))] for displaying a legend in the UI."""
    legend = []
    for value, color in color_map.items():
        legend.append((value, (color.Red, color.Green, color.Blue)))
    return sorted(legend, key=lambda item: item[0])


def apply_coloring(doc, view, rows, param_key, color_map=None):
    """Color ``rows`` in ``view`` by each row's ``param_key`` value.

    ``rows`` are :class:`ElementRow`; only rows matched in Revit are colored.
    Returns the color_map actually used (so the UI can render a legend).
    """
    matched = [r for r in rows if r.element_id is not None]
    if color_map is None:
        color_map = generate_color_map(r.get_cell(param_key) for r in matched)

    # Group element ids by color value.
    groups = {}
    for row in matched:
        key = _norm(row.get_cell(param_key))
        groups.setdefault(key, []).append(row.element_id)

    with revit.Transaction("CDE: color by {}".format(param_key), doc):
        for key, element_ids in groups.items():
            color = color_map.get(key)
            if color is None:
                continue
            apply_color_to_elements(doc, view, element_ids, color)
    return color_map


def reset_coloring(doc, view, rows):
    """Clear temporary overrides for the given rows in ``view``."""
    matched = [r.element_id for r in rows if r.element_id is not None]
    blank = OverrideGraphicSettings()
    with revit.Transaction("CDE: reset colors", doc):
        for element_id in matched:
            view.SetElementOverrides(element_id, blank)
