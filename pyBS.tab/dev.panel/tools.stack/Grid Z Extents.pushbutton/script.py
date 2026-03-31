# -*- coding: utf-8 -*-
"""Set grid 3D vertical extents (Z bottom / Z top) for selected grids."""
# pylint: disable=import-error,invalid-name,broad-except

__title__ = "Grid Z Extents"
__author__ = "Byggstyrning AB"
__doc__ = """Set Z bottom and Z top of selected grids to 0 m and 100 m.

Select one or more model grids (including segments of a multi-segment grid), then run.
Non-grid selection is skipped. Uses Revit Grid.SetVerticalExtents (internal length units).
"""

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import Grid, UnitUtils

try:
    from Autodesk.Revit.DB import UnitTypeId
    _USE_FORGE_UNITS = True
except ImportError:
    from Autodesk.Revit.DB import DisplayUnitType
    _USE_FORGE_UNITS = False

try:
    from Autodesk.Revit.DB import MultiSegmentGrid
except ImportError:
    MultiSegmentGrid = None

from Autodesk.Revit.Exceptions import (
    ArgumentException,
    ArgumentOutOfRangeException,
    InvalidOperationException,
)

from pyrevit import revit, forms, script

logger = script.get_logger()

# Target vertical range (meters); edit here if needed
EXTENT_BOTTOM_M = 0.0
EXTENT_TOP_M = 100.0


def _meters_to_internal(meters):
    """Convert meters to Revit internal length units (feet)."""
    if _USE_FORGE_UNITS:
        return UnitUtils.ConvertToInternalUnits(meters, UnitTypeId.Meters)
    return UnitUtils.ConvertToInternalUnits(meters, DisplayUnitType.DUT_METERS)


def _grid_display_name(grid):
    """Best-effort label for messages."""
    try:
        p = grid.LookupParameter("Name")
        if p and p.HasValue:
            return p.AsString() or ""
    except Exception:
        pass
    return ""


def _collect_grids_from_selection(doc, sel_ids):
    """Return (list of Grid, skipped_non_grid count). Dedupes by element id."""
    grids = []
    seen = set()
    skipped_non_grid = 0

    def _add_grid(g):
        eid = g.Id.IntegerValue
        if eid not in seen:
            seen.add(eid)
            grids.append(g)

    for eid in sel_ids:
        elem = doc.GetElement(eid)
        if elem is None:
            continue
        if isinstance(elem, Grid):
            _add_grid(elem)
        elif MultiSegmentGrid is not None and isinstance(elem, MultiSegmentGrid):
            for gid in elem.GetGridIds():
                child = doc.GetElement(gid)
                if isinstance(child, Grid):
                    _add_grid(child)
        else:
            skipped_non_grid += 1

    return grids, skipped_non_grid


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    sel_ids = list(uidoc.Selection.GetElementIds())
    if not sel_ids:
        forms.alert("Select one or more grids, then run again.", title="Grid Z Extents")
        return

    grids, skipped_non_grid = _collect_grids_from_selection(doc, sel_ids)
    if not grids:
        forms.alert(
            "No grids in the selection. Select model grid lines or a multi-segment grid.",
            title="Grid Z Extents",
        )
        return

    bottom_i = _meters_to_internal(EXTENT_BOTTOM_M)
    top_i = _meters_to_internal(EXTENT_TOP_M)

    ok = 0
    errors = []

    with revit.Transaction("Set grid vertical extents", doc):
        for grid in grids:
            label = _grid_display_name(grid) or str(grid.Id.IntegerValue)
            try:
                grid.SetVerticalExtents(bottom_i, top_i)
                ok += 1
            except ArgumentException as ex:
                errors.append("Grid {0}: {1}".format(label, str(ex)))
            except ArgumentOutOfRangeException as ex:
                errors.append("Grid {0}: {1}".format(label, str(ex)))
            except InvalidOperationException as ex:
                errors.append("Grid {0}: {1}".format(label, str(ex)))
            except Exception as ex:
                errors.append("Grid {0}: {1}".format(label, str(ex)))

    lines = [
        "Updated {0} grid(s): Z bottom = {1} m, Z top = {2} m.".format(
            ok, EXTENT_BOTTOM_M, EXTENT_TOP_M
        ),
    ]
    if skipped_non_grid:
        lines.append("Skipped {0} non-grid element(s).".format(skipped_non_grid))
    if errors:
        lines.append("Failures ({0}):".format(len(errors)))
        for err in errors[:10]:
            lines.append("  - {0}".format(err))
        if len(errors) > 10:
            lines.append("  ... and {0} more.".format(len(errors) - 10))

    logger.info("Grid Z Extents: " + " ".join(lines[:3]))
    forms.alert("\n".join(lines), title="Grid Z Extents")


if __name__ == "__main__":
    main()
