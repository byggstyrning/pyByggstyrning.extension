# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

__title__ = "Reset colors"
__author__ = ""
__doc__ = """This is Reset colors Button.
Click on it to reset colors in the active view."""

if __name__ == '__main__':
    doc = __revit__.ActiveUIDocument.Document
    active_view = doc.ActiveView
    with Transaction(doc, 'Reset Colors') as t:
        t.Start()
        # Create an empty OverrideGraphicSettings to reset colors
        ogs = OverrideGraphicSettings()
        # Get all elements in the view
        collector = FilteredElementCollector(doc, active_view.Id).ToElements()
        for element in collector:
            active_view.SetElementOverrides(element.Id, ogs)
        t.Commit()

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.