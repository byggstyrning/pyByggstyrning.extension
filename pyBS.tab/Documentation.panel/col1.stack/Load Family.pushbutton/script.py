# -*- coding: utf-8 -*-
__title__ = "Load Family"
__author__ = "Jonatan Jacobsson"
__doc__ = """Loads the 3D View Reference family into the current project.
If the project already has an older version of the family it is upgraded;
the references already placed keep their values.
"""

# Import libraries
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

# Get the current Revit document
doc = revit.doc

logger = script.get_logger()

try:
    with revit.Transaction("Load 3D View Reference Family"):
        changed = view_references.load_reference_family(doc, pushbutton_dir)
except Exception as e:
    logger.error("Error loading family: {}".format(str(e)))
    script.exit()

if view_references.find_family_symbol(doc) is None:
    forms.alert("The '{}' family could not be loaded.".format(view_references.FAMILY_NAME),
                title="Load Family")
elif changed:
    forms.show_balloon("Load Family", "'{}' loaded.".format(view_references.FAMILY_NAME))
else:
    forms.show_balloon("Load Family", "'{}' is already loaded and up to date.".format(
        view_references.FAMILY_NAME))
