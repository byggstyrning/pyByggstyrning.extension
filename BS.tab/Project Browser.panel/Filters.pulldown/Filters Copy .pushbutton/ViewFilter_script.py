# Author:	Giuseppe Dotto

__doc__ = "Copy one or more View Filters from a view/template to multiple ones."

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
# Views / Viewtemplates
ops = {"Views":_views, "ViewTemplates":_templates}
selected_view = forms.SelectFromList.show(ops,
					multiselect=False,
					name_attr='Name',
					button_name='Select Views/ViewTemplate')

if not selected_view: script.exit()

# CHOOSE BETWEEN FILTERS
filters = [doc.GetElement(i) for i in selected_view.GetFilters()]
selected_filters = forms.SelectFromList.show(filters,
					multiselect=True,
					name_attr='Name',
					button_name='Select Filters')

if not selected_filters: script.exit()

# DESTINATION VIEWS
target_view = forms.SelectFromList.show(ops,
					multiselect=True,
					name_attr='Name',
					button_name='Select Views/ViewTemplate')

if not target_view: script.exit()


#APPLY
msg = "Some views haven't been edit:\n"
descr = ''
nrOfWarning = 0
if forms.alert("Edit {} views/templates?".format(len(target_view)),
				 yes = True, no = True):
	with revit.Transaction('Copy/Paste view filters'):
		for flt in selected_filters:
			ogs = selected_view.GetFilterOverrides(flt.Id)
			for vw in target_view:
				try:	vw.SetFilterOverrides(flt.Id, ogs)
				except Exception as ex:
					nrOfWarning += 1
					msg += '{}. {} ({})\n'.format(nrOfWarning, vw.Name, flt.Name)
					descr += "{}. {}\n".format(nrOfWarning, ex)

if nrOfWarning > 0:
	forms.alert(msg, expanded=descr)