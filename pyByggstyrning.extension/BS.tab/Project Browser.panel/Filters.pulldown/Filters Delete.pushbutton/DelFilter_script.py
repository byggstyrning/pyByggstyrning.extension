# Author:	Giuseppe Dotto

__doc__ = "Delete selected filter from one or more selected Views/ViewTemplates.\n\n"\
		"SHIFT-CLICK: Delete the filter itself."

from pyrevit import revit, DB, script
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms

doc = revit.doc
collector = DB.FilteredElementCollector

allViewClass = collector(doc).OfClass(DB.View)
_views = []
_templates = []
for v in allViewClass:
	if not v.IsTemplate and v.ViewTemplateId == DB.ElementId.InvalidElementId:
		_views.append(v)
	elif v.IsTemplate:
		_templates.append(v)

# MAKE FORMS
# ALL FILTERS
allFilters = collector(doc).OfClass(DB.ParameterFilterElement)
allFilters = dict([[f.Name, f] for f in allFilters])
selected_filters = forms.SelectFromList.show(sorted(allFilters.keys()),
					multiselect=True,
					button_name='Select Filters to Delete from view')

if not selected_filters: script.exit()
selected_filters = [allFilters[k] for k in selected_filters]

# Views / Viewtemplates
ops = {"Views":_views, "ViewTemplates":_templates}
selected_view = forms.SelectFromList.show(ops,
					multiselect=True,
					name_attr='Name',
					button_name='Select Views/ViewTemplate')

if not selected_view: script.exit()

#APPLY
msg = "Edit {} views/templates and remove {} filters each?".format(len(selected_view), len(selected_filters))
if forms.alert(msg, yes = True, no = True):
	with revit.Transaction('Delete filters for views'):
		for flt in selected_filters:
			for vw in selected_view:
				try:	vw.RemoveFilter(flt.Id)
				except:	pass