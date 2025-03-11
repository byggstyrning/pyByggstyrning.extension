# -*- coding: utf-8 -*-
__title__ = "300"
__author__ = ""
__context__ = 'Selection'
__doc__ = """Underlag för detaljering
Huvudlösningar har tagits fram och beslutats, utarbetade som underlag för ytterligare detaljer.

Sets the MMI parameter value to 300 on selected elements.
Based on MMI veilederen: https://mmi-veilederen.no/?page_id=85"""

from Autodesk.Revit.DB import Transaction, ElementId, FilteredElementCollector, BuiltInCategory
from pyrevit import revit, DB, forms

# Get the current document and selection
doc = revit.doc
selection = revit.get_selection()

if not selection:
    forms.alert('Please select at least one element.', exitscript=True)

# Start a transaction
t = Transaction(doc, 'Set MMI Parameter to 300')
t.Start()

try:
    # Set MMI parameter for each selected element
    for element_id in selection.element_ids:
        element = doc.GetElement(element_id)
        if element:
            # Try to set the parameter value
            param = element.LookupParameter('MMI')
            if param and not param.IsReadOnly:
                param.Set('300')
            else:
                print("Could not set MMI parameter on element {}".format(element_id))
    
    # No success alert as requested
except Exception as e:
    print("Error: {}".format(e))
    forms.alert('Error: {}'.format(e), title='Error')
finally:
    # Commit the transaction
    t.Commit()

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS. 