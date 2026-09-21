# -*- coding: utf-8 -*-
__title__ = "Add to View"
__author__ = "PyRevit Extensions"
__doc__ = """Places a 3D View Reference family at the location and extent of the active view.
If the view already has a reference it is updated instead.
"""

# Import .NET libraries
import clr
clr.AddReference("System")
from System.Collections.Generic import List

# Import Revit API
from Autodesk.Revit.DB import ElementId

# Import pyRevit libraries
import sys
import os.path as op
from pyrevit import revit, script, forms

# Add the extension directory to the path
pushbutton_dir = op.dirname(__file__)
stack_dir = op.dirname(pushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from revit import view_references

logger = script.get_logger()

# Get Revit document and UIDocument
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def create_3d_view_reference():
    """Create or update the 3D view reference for the active view."""
    active_view = uidoc.ActiveView

    if view_references.get_view_kind(active_view) is None:
        forms.alert(
            "'{}' is a {} view. References can be placed for sections, elevations, "
            "detail views and plan callouts.".format(active_view.Name, active_view.ViewType),
            title="View Not Supported", exitscript=True)

    family_symbol = view_references.find_family_symbol(doc)
    if not family_symbol:
        forms.alert(
            "This tool requires the '{}' family.\n\n"
            "Load it with the Load Family button and try again.".format(
                view_references.FAMILY_NAME),
            title="Required Family Not Found", exitscript=True)

    with revit.Transaction("Create 3D View Reference"):
        result = view_references.sync_view_references(doc, [active_view], family_symbol)

    if result.skipped:
        forms.alert("No reference placed for '{}': {}".format(*result.skipped[0]),
                    title="3D View Reference")
        return

    # Select the element instead of isolating it
    uidoc.Selection.SetElementIds(List[ElementId](result.element_ids))


if __name__ == '__main__':
    create_3d_view_reference()
