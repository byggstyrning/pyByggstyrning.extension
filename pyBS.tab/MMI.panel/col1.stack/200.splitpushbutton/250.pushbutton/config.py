# -*- coding: utf-8 -*-
"""Shift-click handler for MMI button 250.

When shift-clicking, this selects all elements with MMI value 250.
"""

__title__ = "Select MMI 250"
__author__ = ""
__doc__ = "Shift-click: Select all elements with MMI value 250"

# Add the lib directory to sys.path for importing
import sys
import os.path as op
script_path = __file__
pushbutton_dir = op.dirname(script_path)
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

from pyrevit import revit, forms
from mmi.utils import select_elements_by_mmi

# Select elements with MMI value 250
count = select_elements_by_mmi(revit.doc, revit.uidoc, "250")

if count == 0:
    forms.show_balloon(
        header="No Elements Found",
        text="No elements found with MMI value 250",
        is_new=True
    )

