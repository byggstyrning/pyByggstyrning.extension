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


def collect_revit_infos(doc, ifc_class, view=None, include_params=False, param_reader=None):
    """Return {global_id: revit_info} in a single collector pass (no re-GetElement)."""
    infos = {}
    for element in collect_revit_elements(doc, ifc_class, view):
        guid = read_ifc_guid(element)
        if not guid:
            continue
        info = get_revit_info(doc, element, ifc_class)
        if include_params and param_reader is not None:
            info["revit_params"] = param_reader(element)
        else:
            info["revit_params"] = {}
        infos[guid] = info
    return infos


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


def _room_info(room):
    """Return (number, name) for a room using BuiltInParameter (IronPython-safe).

    room.Name raises AttributeError in IronPython; use parameter accessors.
    """
    if room is None:
        return "", ""
    try:
        from Autodesk.Revit.DB import BuiltInParameter as BIP
        number = ""
        name = ""
        try:
            p = room.get_Parameter(BIP.ROOM_NUMBER)
            if p:
                number = p.AsString() or ""
        except Exception:
            pass
        try:
            p = room.get_Parameter(BIP.ROOM_NAME)
            if p:
                name = p.AsString() or ""
        except Exception:
            pass
        return number, name
    except Exception:
        return "", ""


def get_door_rooms(doc, element):
    """Return (from_name, from_number, to_name, to_number) for a door instance.

    FromRoom/ToRoom are phase-indexed properties in IronPython — accessing
    them without a Phase returns an indexer object, not a Room.  Use the last
    project phase (matches Revit's own door schedule behaviour).
    """
    from_name = from_number = to_name = to_number = ""
    try:
        phases = doc.Phases
        phase = phases[phases.Size - 1]
        try:
            from_number, from_name = _room_info(element.FromRoom[phase])
        except Exception:
            pass
        try:
            to_number, to_name = _room_info(element.ToRoom[phase])
        except Exception:
            pass
    except Exception:
        pass
    return from_name, from_number, to_name, to_number


def get_revit_info(doc, element, ifc_class="IfcDoor"):
    """Return a dict of Revit-derived fields for the schedule row."""
    info = {
        "element_id": element.Id,
        "element_id_value": get_element_id_value(element.Id),
        "mark": _param_string(element, BuiltInParameter.ALL_MODEL_MARK),
        "level": get_level_name(doc, element),
        "from_room": "",
        "from_room_number": "",
        "to_room": "",
        "to_room_number": "",
    }
    if ifc_class == "IfcDoor":
        (info["from_room"], info["from_room_number"],
         info["to_room"], info["to_room_number"]) = get_door_rooms(doc, element)
    return info


# --- Revit parameter scan (schedule column source) -------------------------

REVIT_KEY_PREFIX = "Revit:"


def _revit_param_key(name):
    return "{}{}".format(REVIT_KEY_PREFIX, name)


def _revit_value_type(param):
    """Map a Revit parameter to schedule value_type (best-effort)."""
    try:
        from Autodesk.Revit.DB import StorageType
        st = param.StorageType
        if st == StorageType.Integer:
            try:
                from Autodesk.Revit.DB import ParameterType
                if param.Definition.ParameterType == ParameterType.YesNo:
                    return "bool"
            except Exception:
                pass
            return "number"
        if st == StorageType.Double:
            return "number"
    except Exception:
        pass
    return "string"


def _read_param_value(param):
    """Read a parameter value into a Python type suitable for the grid."""
    try:
        from Autodesk.Revit.DB import StorageType
        if not param.HasValue:
            return ""
        st = param.StorageType
        if st == StorageType.Integer:
            try:
                from Autodesk.Revit.DB import ParameterType
                if param.Definition.ParameterType == ParameterType.YesNo:
                    return bool(param.AsInteger())
            except Exception:
                pass
            return param.AsInteger()
        if st == StorageType.Double:
            return param.AsDouble()
        if st == StorageType.String:
            return param.AsString() or ""
        if st == StorageType.ElementId:
            return param.AsValueString() or ""
    except Exception:
        pass
    try:
        return param.AsValueString() or ""
    except Exception:
        return ""


def read_revit_parameters(element):
    """Return {param_key: value} for instance/type params on one element."""
    values = {}
    try:
        for param in element.Parameters:
            if param is None or param.IsReadOnly:
                continue
            try:
                name = param.Definition.Name
            except Exception:
                continue
            if not name:
                continue
            key = _revit_param_key(name)
            values[key] = _read_param_value(param)
    except Exception:
        pass
    return values


def collect_revit_parameter_defs(doc, revit_infos, param_def_fn, max_elements=25):
    """Scan matched Revit elements and return ParameterDef rows (source=revit)."""
    defs = []
    seen = set()
    count = 0
    for info in (revit_infos or {}).values():
        if count >= max_elements:
            break
        eid = info.get("element_id")
        if eid is None:
            continue
        try:
            element = doc.GetElement(eid)
        except Exception:
            continue
        if element is None:
            continue
        count += 1
        try:
            for param in element.Parameters:
                if param is None or param.IsReadOnly:
                    continue
                try:
                    name = param.Definition.Name
                except Exception:
                    continue
                if not name:
                    continue
                key = _revit_param_key(name)
                if key in seen:
                    continue
                seen.add(key)
                defs.append(param_def_fn(
                    key, name, _revit_value_type(param), group="Revit", source="revit"))
        except Exception:
            pass
    defs.sort(key=lambda d: d.label.upper())
    return defs
