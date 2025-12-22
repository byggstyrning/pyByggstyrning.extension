# -*- coding: utf-8 -*-
"""Isolate all 3D Zone Generic Model instances in the active view.

Finds all Generic Model family instances whose family name begins with "3DZone"
and isolates them in the active view. Works with 3D views and plan views.
"""

__title__ = "Isolate 3D Zones"
__author__ = "Byggstyrning AB"
__doc__ = "Isolate all 3D Zone Generic Model instances in the active view"

# Import standard libraries
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from System.Collections.Generic import List

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Initialize logger
logger = script.get_logger()

if __name__ == '__main__':
    doc = revit.doc
    uidoc = revit.uidoc
    
    # Get active view
    active_view = doc.ActiveView
    view_type = active_view.ViewType
    
    # Check if view is 3D or plan view
    if view_type not in [ViewType.ThreeD, ViewType.FloorPlan, ViewType.CeilingPlan]:
        forms.alert(
            "This tool only works in 3D views or plan views.\n\n"
            "Current view type: {}".format(view_type),
            title="Invalid View Type",
            exitscript=True
        )
    
    logger.debug("Active view: {} (Type: {})".format(active_view.Name, view_type))
    
    # Find all Generic Model family instances
    collector = FilteredElementCollector(doc)\
        .OfClass(FamilyInstance)\
        .OfCategory(BuiltInCategory.OST_GenericModel)\
        .WhereElementIsNotElementType()
    
    all_generic_models = collector.ToElements()
    
    # Filter for instances whose family name begins with "3DZone"
    zone_instances = []
    for instance in all_generic_models:
        try:
            family_name = instance.Symbol.FamilyName
            if family_name and family_name.startswith("3DZone"):
                zone_instances.append(instance)
        except Exception as ex:
            logger.debug("Error checking family name for element {}: {}".format(
                instance.Id, ex))
            continue
    
    if not zone_instances:
        forms.alert(
            "No 3D Zone instances found in the model.\n\n"
            "3D Zones are Generic Model family instances whose family name begins with '3DZone'.",
            title="No 3D Zones Found",
            exitscript=True
        )
    
    logger.debug("Found {} 3D Zone instances".format(len(zone_instances)))
    
    # Isolate the elements
    try:
        element_ids = List[ElementId]([inst.Id for inst in zone_instances])
        
        with revit.Transaction("Isolate 3D Zones"):
            active_view.IsolateElementsTemporary(element_ids)
        
        logger.debug("Isolated {} 3D Zone instances in view: {}".format(
            len(zone_instances), active_view.Name))
        
    except Exception as ex:
        logger.error("Error isolating elements: {}".format(ex))
        forms.alert(
            "Error isolating 3D Zones:\n\n{}".format(str(ex)),
            title="Isolation Error",
            exitscript=True
        )

