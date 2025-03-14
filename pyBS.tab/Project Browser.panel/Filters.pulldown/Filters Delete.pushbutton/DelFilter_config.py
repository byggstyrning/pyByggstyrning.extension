# Author:	Giuseppe Dotto


from pyrevit import revit, DB, script
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

doc = revit.doc
collector = DB.FilteredElementCollector

allFilters = collector(doc).OfClass(DB.ParameterFilterElement)
allFilters = dict([[f.Name, f] for f in allFilters])
# MAKE FORMS
selected_filters = forms.SelectFromList.show(sorted(allFilters.keys()),
					multiselect=True,
					button_name='Select Filters to Delete')

if not selected_filters: script.exit()
selected_filters = [allFilters[k] for k in selected_filters]

#APPLY
msg = "Delete {} filters?".format(len(selected_filters))
if forms.alert(msg, yes = True, no = True):
	with revit.Transaction('Delete filters.'):
		for flt in selected_filters:
			try:	doc.Delete(flt.Id)
			except:	pass