# -*- coding: utf-8 -*-
__title__ = "350"
__author__ = ""
__context__ = 'Selection'
__doc__ = """Samordnat för granskning
Objektet är samordnat med relevanta discipliner för att undvika kollisioner. Kan omfatta anbud. Alla relevanta kollisioner är hanterade. Huvuddesignen är låst. Kan omfatta förenklade 3D-objekt. Placering i plan, profil och höjd är definierad.

Sets the MMI parameter value to 350 on selected elements.
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
set_selection_mmi_value(revit.doc, "350")

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS. 