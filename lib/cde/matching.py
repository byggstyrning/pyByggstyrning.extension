# -*- coding: utf-8 -*-
"""Join CDE elements (by IFC GlobalId) to Revit elements.

Collects Revit elements of a mapped category (doors first), reads their IFC
GlobalId, and exposes the Revit-derived fields the schedule groups/filters on
(mark, level, from/to room). Category is IfcClass-driven so other classes
follow the same path.
"""
import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter)

from pyrevit import script

import sys
import os.path as op
_lib_dir = op.dirname(op.dirname(__file__))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

from revit.compat import get_element_id_value

logger = script.get_logger()

# IfcClass -> Revit BuiltInCategory. Extend as more categories are supported.
IFC_CLASS_TO_BIC = {
    "IfcDoor": BuiltInCategory.OST_Doors,
    "IfcWindow": BuiltInCategory.OST_Windows,
    "IfcWall": BuiltInCategory.OST_Walls,
}


def get_builtin_category(ifc_class):
    return IFC_CLASS_TO_BIC.get(ifc_class)


def read_ifc_guid(element):
    """Return the element's IFC GlobalId, or None.

    Prefers the built-in ``IFC_GUID`` parameter and falls back to the common
    custom parameter spellings used by some exporters.
    """
    try:
        param = element.get_Parameter(BuiltInParameter.IFC_GUID)
        if param and param.HasValue:
            value = param.AsString()
            if value:
                return value
    except Exception:
        pass
    for name in ("IfcGUID", "IFCGuid", "IFC GUID", "IfcGuid"):
        try:
            param = element.LookupParameter(name)
            if param and param.HasValue:
                value = param.AsString()
                if value:
                    return value
        except Exception:
            continue
    return None


def collect_revit_elements(doc, ifc_class, view=None):
    """Return Revit instance elements of ``ifc_class``'s category.

    When ``view`` is given, only elements owned by/visible in that view are
    returned (used by the "active view only" toggle).
    """
    bic = get_builtin_category(ifc_class)
    if bic is None:
        logger.warn("CDE: no Revit category mapped for {}".format(ifc_class))
        return []
    if view is not None:
        collector = FilteredElementCollector(doc, view.Id)
    else:
        collector = FilteredElementCollector(doc)
    return list(collector.OfCategory(bic).WhereElementIsNotElementType().ToElements())


def build_guid_index(doc, ifc_class, view=None):
    """Return {global_id: ElementId} for matched Revit elements."""
    index = {}
    for element in collect_revit_elements(doc, ifc_class, view):
        guid = read_ifc_guid(element)
        if guid:
            index[guid] = element.Id
    return index


def get_active_view_guids(doc, view, ifc_class):
    """Return the set of IFC GlobalIds present in ``view``."""
    return set(build_guid_index(doc, ifc_class, view).keys())


# --- Revit-derived fields -------------------------------------------------

def _param_string(element, built_in):
    try:
        param = element.get_Parameter(built_in)
        if param and param.HasValue:
            return param.AsString() or param.AsValueString() or ""
    except Exception:
        pass
    return ""


def get_level_name(doc, element):
    try:
        if getattr(element, "LevelId", None) and get_element_id_value(element.LevelId) > 0:
            level = doc.GetElement(element.LevelId)
            if level is not None:
                return level.Name
    except Exception:
        pass
    return _param_string(element, BuiltInParameter.FAMILY_LEVEL_PARAM)


def _room_name(room):
    if room is None:
        return ""
    try:
        return room.Name or ""
    except Exception:
        return ""


def get_door_rooms(element):
    """Return (from_room_name, to_room_name) for a door instance, best-effort."""
    from_name = to_name = ""
    try:
        from_name = _room_name(getattr(element, "FromRoom", None))
    except Exception:
        pass
    try:
        to_name = _room_name(getattr(element, "ToRoom", None))
    except Exception:
        pass
    return from_name, to_name


def get_revit_info(doc, element, ifc_class="IfcDoor"):
    """Return a dict of Revit-derived fields for the schedule row."""
    info = {
        "element_id": element.Id,
        "element_id_value": get_element_id_value(element.Id),
        "mark": _param_string(element, BuiltInParameter.ALL_MODEL_MARK),
        "level": get_level_name(doc, element),
        "from_room": "",
        "to_room": "",
    }
    if ifc_class == "IfcDoor":
        info["from_room"], info["to_room"] = get_door_rooms(element)
    return info
