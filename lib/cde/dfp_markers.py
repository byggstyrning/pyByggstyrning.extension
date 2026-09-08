# -*- coding: utf-8 -*-
"""Build TemporaryGraphics marker sessions from CDE schedule rows + DFP values."""

from __future__ import division

from Autodesk.Revit.DB import XYZ

from cde.dfp_icons import (
    code_from_value_key,
    get_dfp_summary_composite,
    composite_layout_dict,
    build_marker_tooltip,
    has_authored_dfp_value,
    is_dfp_value_key,
    dfp_code_sort_key,
    _MAX_OVERLAY_PX,
)
from revit.compat import get_element_id_value

DFP_SESSION_ID = "dfp"
DEFAULT_MAX_DOORS = 400


def _door_marker_xyz(element):
    """Anchor above door bbox (Hub overlay places markers above the door)."""
    try:
        bbox = element.get_BoundingBox(None)
        if bbox is not None:
            cx = (bbox.Min.X + bbox.Max.X) / 2.0
            cy = (bbox.Min.Y + bbox.Max.Y) / 2.0
            dz = bbox.Max.Z - bbox.Min.Z
            lift = max(dz * 0.12, 0.8)
            return XYZ(cx, cy, bbox.Max.Z + lift)
    except Exception:
        pass
    try:
        loc = element.Location
        if loc is not None and hasattr(loc, "Point"):
            pt = loc.Point
            return XYZ(pt.X, pt.Y, pt.Z + 1.0)
    except Exception:
        pass
    return None


def _sorted_param_keys(param_keys):
    keys = [k for k in (param_keys or []) if is_dfp_value_key(k)]
    return sorted(keys, key=lambda k: dfp_code_sort_key(code_from_value_key(k) or k))


def _active_flags_for_row(row, sorted_keys):
    return [has_authored_dfp_value(row.get_cell(k)) for k in sorted_keys]


def build_dfp_view_points(doc, rows, param_keys, view=None, max_doors=None):
    """Return ``{view_id: [point_dict, ...]}`` — one summary composite per door.

    Collapsed: bitmap of active/true functions only (not clickable).
    Hover (driver): overlay cell buttons aligned to full column grid.
    """
    if view is None:
        try:
            view = doc.ActiveView
        except Exception:
            return {}
    if view is None:
        return {}
    sorted_keys = _sorted_param_keys(param_keys)
    if not sorted_keys:
        return {}

    door_budget = max_doors if max_doors is not None else DEFAULT_MAX_DOORS
    view_id = get_element_id_value(view.Id)
    overlay_layout = composite_layout_dict(len(sorted_keys), max_px=_MAX_OVERLAY_PX)
    points = []
    doors_used = 0

    for row in rows:
        if doors_used >= door_budget:
            break
        if row.element_id is None:
            continue
        element = doc.GetElement(row.element_id)
        if element is None:
            continue
        anchor = _door_marker_xyz(element)
        if anchor is None:
            continue

        active_flags = _active_flags_for_row(row, sorted_keys)
        active_codes = []
        for key, active in zip(sorted_keys, active_flags):
            if active:
                code = code_from_value_key(key)
                if code:
                    active_codes.append(code)

        try:
            img_path, summary_layout, _ = get_dfp_summary_composite(active_codes)
        except Exception:
            continue

        mark = getattr(row, "Mark", None) or getattr(row, "GlobalId", "")
        gid = getattr(row, "GlobalId", "") or ""
        tooltip = (
            build_marker_tooltip(active_codes)
            if active_codes
            else u"(no functions — hover to assign)"
        )
        if mark:
            tooltip = u"{}\n{}".format(mark, tooltip)
        tooltip = u"{}\nHover to edit functions.".format(tooltip)

        points.append({
            "x": anchor.X,
            "y": anchor.Y,
            "z": anchor.Z,
            "image_path": img_path,
            "tooltip": tooltip,
            "clickable": False,
            "dfp_door": True,
            "session_id": DFP_SESSION_ID,
            "global_id": gid,
            "mark": mark,
            "sorted_keys": list(sorted_keys),
            "active_flags": list(active_flags),
            "overlay_layout": overlay_layout,
            "summary_layout": summary_layout,
        })
        doors_used += 1

    if not points:
        return {}
    return {view_id: points}


def default_dfp_param_keys(param_defs, visible_keys=None):
    """All catalogued DFP keys, optionally restricted to visible grid columns."""
    keys = []
    for pdef in param_defs or []:
        key = getattr(pdef, "key", None)
        if key and is_dfp_value_key(key):
            keys.append(key)
    if visible_keys:
        visible = set(visible_keys)
        filtered = [k for k in keys if k in visible]
        if filtered:
            return filtered
    return keys
