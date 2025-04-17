# -*- coding: utf-8 -*-
"""Template for creating new MMI pushbuttons.

Copy this file to your new pushbutton folder and customize as needed.
"""

__title__ = "Template"  # Replace with the actual MMI value
__author__ = ""
__context__ = 'Selection'
__doc__ = """MMI Value Description Here
Add a detailed description of what this MMI value means.

Sets the MMI parameter value to X on selected elements.
Based on MMI veilederen: https://mmi-veilederen.no/?page_id=85"""

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

from pyrevit import revit
from mmi.core import set_selection_mmi_value

# Set the MMI value on selected elements
# Replace "XXX" with the actual MMI value for this button
set_selection_mmi_value(revit.doc, "XXX")

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS. 