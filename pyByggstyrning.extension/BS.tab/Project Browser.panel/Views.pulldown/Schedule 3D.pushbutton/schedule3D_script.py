# Author:	Giuseppe Dotto

__doc__ =	"Create a 3D View showing all the element listed in the given schedule."

from sched import scheduler
from pyrevit import revit, DB, UI
from pyrevit import PyRevitException, PyRevitIOError
from pyrevit import forms, script

from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,
							Separator, Button, CheckBox)


# import DB
doc = revit.doc
uidoc = __revit__.ActiveUIDocument

collector = DB.FilteredElementCollector(doc)
BIC = DB.BuiltInCategory
BIP = DB.BuiltInParameter


def select_schedules():
	all_schedule = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)
	all_schedule = sorted(list(all_schedule), key=lambda s:s.Name)
	return	forms.SelectFromList.show(all_schedule, multiselect=True, name_attr='Name')


def give_viewName(nm):
	existing_3DView = DB.FilteredElementCollector(doc).OfClass(DB.View3D)
	existing_3DView = [v.Name for v in existing_3DView]
	for i in range(1,50):
		if nm not in existing_3DView:
			return	nm
		else:
			nm += ' ({})'.format(i)


def match_phase(vw1, vw2):
	vw2.Parameter[BIP.VIEW_PHASE].Set(vw1.Parameter[BIP.VIEW_PHASE].AsElementId())
	vw2.Parameter[BIP.VIEW_PHASE_FILTER].Set(vw1.Parameter[BIP.VIEW_PHASE_FILTER].AsElementId())
	

# SELECT SCHEDULE
active_view = doc.ActiveView
if active_view.GetType() == DB.ViewSchedule:
	if forms.alert("Do you want to use current schedule?", yes=True, no=True, warn_icon=False):
		select_schedules = [active_view]
	else:
		select_schedules = select_schedules()
else:
	select_schedules = select_schedules()


vft = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
vft = [t for t in vft if t.ViewFamily == DB.ViewFamily.ThreeDimensional][0]
with revit.TransactionGroup('Create 3dViews from Schedules'):
	for sc in select_schedules:
		# GET ELEMENTS
		all_elem_ids = DB.FilteredElementCollector(doc, sc.Id).ToElementIds()

		with revit.Transaction('create View'):
			# CREATE 3D VIEW
			new_view = DB.View3D.CreateIsometric(doc, vft.Id)
			new_view.IsolateElementsTemporary(all_elem_ids)

			nm = 'Schedule 3D - ' + sc.Name
			new_view.Name = give_viewName(nm)

			# CORRECT SETTINGS
			match_phase(sc, new_view)
			
		uidoc.RequestViewChange(new_view)