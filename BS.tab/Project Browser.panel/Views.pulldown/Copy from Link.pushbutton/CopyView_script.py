
# 2023, Macro4BIM (www.macro4bim.com)
# Author: Giuseppe Dotto

__title__ = 'Copy views\nfrom Link'
__doc__ = "Copy selected views from the selected linked model into the current project.\n"\
		"The views comes with the same graphic style but without the annotations. It is not possible to copy Dependent views."

from System.Collections.Generic import *
from pyrevit import DB, revit, forms, script
# import DB

doc = revit.doc

with forms.WarningBar(title='SELECT REVIT LINK'):
	link = revit.pick_element_by_category(DB.BuiltInCategory.OST_RvtLinks)
if not link:	script.exit()
link_doc = link.GetLinkDocument()

linked_views = DB.FilteredElementCollector(link_doc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType()
if not linked_views:	script.exit()

linked_views = [v for v in linked_views if not v.IsTemplate]

linked_views.sort(key=lambda v: v.Name)

selected = forms.SelectFromList.show(linked_views, button_name='Select Item', name_attr='Name', multiselect=True)
selected = [v.Id for v in selected]
selected = List[DB.ElementId](selected)

with revit.Transaction('Copy views for Link'):
	DB.ElementTransformUtils.CopyElements(link_doc, selected, doc, None, None)

	nr = len(selected)
	forms.alert('{} views copied.'.format(nr), warn_icon=False)