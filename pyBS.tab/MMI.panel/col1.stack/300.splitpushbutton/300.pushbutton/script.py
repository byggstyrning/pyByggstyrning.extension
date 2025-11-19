# -*- coding: utf-8 -*-
__title__ = "300"
__author__ = ""
__doc__ = """Klar för utvärdering
Objektet är vidareutvecklat för produktion och anbud. Mindre justeringar kan förekomma. Objektet är placerat i plan, profil och höjd. Objektet är nu i 3D. Tekniska data som geometri, form och toleranser kan vara definierade.

If elements are selected: Sets the MMI parameter value to 300 on selected elements.
If no selection: Selects all elements with MMI value 300.
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
from mmi.utils import select_elements_by_mmi

# Check if there's a selection
selection = revit.get_selection()

if not selection or not selection.element_ids:
    # No selection: select elements by MMI value
    select_elements_by_mmi(revit.doc, revit.uidoc, "300")
else:
    # Selection exists: set MMI value on selected elements
    set_selection_mmi_value(revit.doc, "300")

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS. 