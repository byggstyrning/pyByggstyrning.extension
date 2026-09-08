# -*- coding: utf-8 -*-
"""Zone Audit: explain, element by element, why a 3D Zone mapping is or is not written.

The tracer walks the same gates as ``core.write_parameters_to_elements`` in the same
order, calling the same ``core``/``containment`` functions, and records a reason code
per element instead of a silent counter. It is READ-ONLY: it never writes a parameter
and never starts a transaction.

Two entry points:

- ``audit_zone(ctx, zone)``   -> every model element whose bounding box touches the
  zone, classified into buckets (MAPPED_HERE / MISSED / MAPPED_OTHER / OUTSIDE ...).
  "Geometry intersects" is decided independently of the sample-point method, with
  Revit's ``ElementIntersectsSolidFilter`` against the zone solid(s), so an element
  that is geometrically inside the zone but skipped by Write shows up as MISSED with
  the gate that dropped it.
- ``trace_element(ctx, element)`` -> the full decision trail for one element: sample
  points, index cells, per-zone bbox/solid hits, vote tally, the result of the real
  containment call, and a dry-run of the parameter copy.

Reconstruction vs reality: for the "element" (3D Zone) strategy the tracer rebuilds
``containment.get_containing_element_indexed`` step by step to get per-point detail,
then ALSO calls the real function and flags ``tracer_mismatch`` when the two disagree.
"""

import math
import os
import os.path as op
import sys
import time
import datetime
import codecs
import traceback
from collections import defaultdict

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter, ElementId, StorageType,
    Outline, BoundingBoxIntersectsFilter, ElementIntersectsSolidFilter, Options,
    CategoryType, FamilyInstance, ElementType, XYZ,
    HostObject, HostObjectUtils, ShellLayerType, PlanarFace,
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.DB.Mechanical import Space
from System.Collections.Generic import List
from pyrevit import script

try:
    from zone3d import containment
    from zone3d import core
except ImportError:
    import containment
    import core

# lib/ must be importable for revit.compat / revit.revit_utils (same trick as core.py)
_lib_dir = op.dirname(op.dirname(op.abspath(__file__)))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)
from revit.compat import get_element_id_value, make_element_id
from revit.revit_utils import is_element_editable

logger = script.get_logger()

THREE_D_ZONE_MARKER = "3DZONE_FILTER"

# Must match core.write_parameters_to_elements (element_index_cell_size)
CELL_SIZE_FEET = 50.0
FT_TO_MM = 304.8

# --------------------------------------------------------------------------------------
# Reason codes (in gate order) and buckets
# --------------------------------------------------------------------------------------

REASONS = [
    ("MAPPED", "Contained by the zone; at least one parameter would be written"),
    ("ALREADY_CORRECT", "Contained by the zone; every target value already equals the zone value"),
    ("NOT_TARGET_CATEGORY", "Category is not in the configuration's target categories"),
    ("NOT_IN_VIEW", "'Active view only' is on and the element is not visible in the active view"),
    ("EXCLUDED_ROOM_OR_ZONE", "Rooms and 3DZone families are never targets in the room strategy"),
    ("NO_TARGET_PARAM", "None of the target parameters exists (writable) on the instance or its type"),
    ("NOT_EDITABLE", "Workshared model: element is owned by another user"),
    ("TARGET_NOT_EMPTY", "Only-empty configuration skips it, and its existing value differs from the zone's value (or no zone found)"),
    ("TARGET_NOT_EMPTY_MATCHES", "Only-empty configuration skips it, and its existing value already equals the zone's value"),
    ("NO_TEST_POINTS", "No sample points could be derived (no Location and no bounding box)"),
    ("INDEX_MISS", "No zone in the 3x3 index cells (50 ft) around the first sample point"),
    ("ALL_POINTS_MISS", "Zones were candidates, but no sample point is inside any zone solid and the element's geometry intersects no zone solid"),
    ("VOTE_BELOW_THRESHOLD", "Floor/roof vote: best zone got fewer than half of the sample points"),
    ("PHASE_MISMATCH", "Room strategy: a room contains the points, but not in a phase where the element exists"),
    ("NOT_CONTAINED", "Strategy reports no containing element (no per-point detail for this strategy)"),
    ("ZONE_NAME_FILTER", "Containing element is not a 3DZone family (family name filter)"),
    ("CONTAINED_OTHER_ZONE", "Contained, but by a different zone than the one audited"),
    ("SOURCE_PARAM_EMPTY", "Zone found, but the zone's source parameter has no value"),
    ("SOURCE_PARAM_MISSING", "Zone found, but the zone has no parameter with the source name"),
    ("TARGET_PARAM_MISSING", "Zone found, but the target parameter is missing at copy time"),
    ("TARGET_PARAM_READONLY", "Zone found, but the target parameter is read-only"),
    ("STORAGE_TYPE_MISMATCH", "Zone found, but source and target storage types differ (e.g. Text vs Integer)"),
    ("UNSUPPORTED_STORAGE", "Zone found, but the parameter storage type is not supported by Write"),
    ("TRACE_ERROR", "The tracer failed on this element (see detail)"),
]
REASON_DESCRIPTIONS = dict(REASONS)

MAPPED_REASONS = ("MAPPED", "ALREADY_CORRECT", "TARGET_NOT_EMPTY_MATCHES")

BUCKETS = [
    ("MISSED", "Solid intersects the audited zone, but Write does not map it to this zone"),
    ("MAPPED_HERE", "Mapped to the audited zone"),
    ("MAPPED_OTHER", "Solid intersects the audited zone, but Write maps it to another zone"),
    ("MAPPED_HERE_NO_GEOM", "Mapped to this zone although Revit's solid filter finds no intersection"),
    ("OUTSIDE", "Only the bounding box touches the zone; the solid does not. Not mapped here"),
    ("UNVERIFIED", "Revit's solid intersection test failed for this element; mapping status only"),
]
BUCKET_DESCRIPTIONS = dict(BUCKETS)

# Priority when several parameters fail for different reasons: first match wins
_PARAM_FAILURE_PRIORITY = [
    "TARGET_PARAM_MISSING", "TARGET_PARAM_READONLY", "STORAGE_TYPE_MISMATCH",
    "SOURCE_PARAM_EMPTY", "SOURCE_PARAM_MISSING", "UNSUPPORTED_STORAGE",
]


# --------------------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------------------

def _id(element_or_id):
    try:
        if isinstance(element_or_id, ElementId):
            return get_element_id_value(element_or_id)
        return get_element_id_value(element_or_id.Id)
    except Exception:
        return None


def _safe(fn, default=None):
    try:
        return fn()
    except Exception:
        return default


def _u(value):
    """Best-effort unicode conversion for log output."""
    if value is None:
        return u""
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value).decode("utf-8", "replace")
        except Exception:
            return u"?"


def ids_str(ids):
    """'[1, 2, 3]' without IronPython's long-integer L suffix."""
    try:
        return u"[" + u", ".join(u"{}".format(i) for i in sorted(ids)) + u"]"
    except Exception:
        return _u(ids)


def fmt_pt(pt):
    """Format an XYZ in Revit internal feet and in mm."""
    if pt is None:
        return u"(none)"
    return u"({:.3f}, {:.3f}, {:.3f}) ft = ({:.0f}, {:.0f}, {:.0f}) mm".format(
        pt.X, pt.Y, pt.Z, pt.X * FT_TO_MM, pt.Y * FT_TO_MM, pt.Z * FT_TO_MM)


def fmt_pt_short(pt):
    if pt is None:
        return u"(none)"
    return u"({:.0f},{:.0f},{:.0f})mm".format(pt.X * FT_TO_MM, pt.Y * FT_TO_MM, pt.Z * FT_TO_MM)


def fmt_bbox(bbox):
    if bbox is None:
        return u"(no bbox)"
    return u"min {} max {}".format(fmt_pt_short(bbox.Min), fmt_pt_short(bbox.Max))


def param_value_str(param):
    """Human readable parameter value (None when the parameter has no value)."""
    try:
        if param is None or not param.HasValue:
            return None
        st = param.StorageType
        if st == StorageType.String:
            return param.AsString()
        if st == StorageType.Integer:
            return u"{}".format(param.AsInteger())
        if st == StorageType.Double:
            return u"{:.6g}".format(param.AsDouble())
        if st == StorageType.ElementId:
            eid = param.AsElementId()
            return u"id:{}".format(get_element_id_value(eid)) if eid else None
        return _u(param.AsValueString())
    except Exception:
        return None


def storage_type_name(param):
    try:
        return _u(param.StorageType).replace("StorageType.", "")
    except Exception:
        return u"?"


def element_type_of(element):
    try:
        type_id = element.GetTypeId()
        if type_id and get_element_id_value(type_id) > 0:
            return element.Document.GetElement(type_id)
    except Exception:
        pass
    return None


def describe_element(element):
    """Return a dict with id, category, family, type, name, level, workset for an element."""
    info = {
        "id": _id(element),
        "category": u"",
        "family": u"",
        "type": u"",
        "name": u"",
        "level": u"",
        "workset": u"",
        "class": u"",
    }
    try:
        info["class"] = _u(element.GetType().Name)
    except Exception:
        pass
    try:
        if element.Category:
            info["category"] = _u(element.Category.Name)
    except Exception:
        pass
    try:
        info["name"] = _u(element.Name)
    except Exception:
        pass
    try:
        if isinstance(element, FamilyInstance) and element.Symbol:
            info["family"] = _u(element.Symbol.FamilyName)
            info["type"] = _u(element.Symbol.Name)
        else:
            et = element_type_of(element)
            if et is not None:
                info["type"] = _u(et.Name)
                fam_param = et.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                if fam_param and fam_param.HasValue:
                    info["family"] = _u(fam_param.AsString())
                elif hasattr(et, "FamilyName"):
                    info["family"] = _u(et.FamilyName)
    except Exception:
        pass
    info["level"] = describe_level(element)
    try:
        doc = element.Document
        if doc.IsWorkshared:
            ws = doc.GetWorksetTable().GetWorkset(element.WorksetId)
            if ws:
                info["workset"] = _u(ws.Name)
    except Exception:
        pass
    return info


_LEVEL_PARAMS = [
    "FAMILY_LEVEL_PARAM", "INSTANCE_REFERENCE_LEVEL_PARAM", "SCHEDULE_LEVEL_PARAM",
    "INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM", "WALL_BASE_CONSTRAINT", "LEVEL_PARAM",
    "RBS_START_LEVEL_PARAM", "ROOM_LEVEL_ID",
]


def describe_level(element):
    """Level name for an element (LevelId first, then the usual level parameters)."""
    try:
        doc = element.Document
        level_id = None
        if hasattr(element, "LevelId") and element.LevelId and get_element_id_value(element.LevelId) > 0:
            level_id = element.LevelId
        if level_id is None:
            for bip_name in _LEVEL_PARAMS:
                bip = getattr(BuiltInParameter, bip_name, None)
                if bip is None:
                    continue
                try:
                    p = element.get_Parameter(bip)
                except Exception:
                    p = None
                if p and p.HasValue and p.StorageType == StorageType.ElementId:
                    eid = p.AsElementId()
                    if eid and get_element_id_value(eid) > 0:
                        level_id = eid
                        break
        if level_id is not None:
            lvl = doc.GetElement(level_id)
            if lvl is not None:
                return _u(lvl.Name)
    except Exception:
        pass
    return u""


def describe_element_short(element):
    info = describe_element(element)
    label = u"{} | {}".format(info["id"], info["category"])
    if info["family"] or info["type"]:
        label += u" | {} : {}".format(info["family"], info["type"])
    elif info["name"]:
        label += u" | {}".format(info["name"])
    if info["level"]:
        label += u" | Level {}".format(info["level"])
    return label


def zone_label(ctx, zone):
    """'family : type #id [param=value, ...]' using the configuration's source parameters."""
    info = describe_element(zone)
    values = []
    for name in ctx.source_param_names:
        try:
            p = zone.LookupParameter(name)
            if p is None:
                et = element_type_of(zone)
                p = et.LookupParameter(name) if et is not None else None
            v = param_value_str(p)
            values.append(u"{}={}".format(name, v if v is not None else u"<empty>"))
        except Exception:
            values.append(u"{}=?".format(name))
    label = u"{} : {} #{}".format(info["family"] or info["category"], info["type"], info["id"])
    if values:
        label += u" [{}]".format(u", ".join(values))
    if info["level"]:
        label += u" (Level {})".format(info["level"])
    return label


def _is_identity_transform(t):
    # Same test as containment.get_containing_element_indexed
    return (t.Origin.X == 0 and t.Origin.Y == 0 and t.Origin.Z == 0 and
            t.BasisX.X == 1 and t.BasisX.Y == 0 and t.BasisX.Z == 0 and
            t.BasisY.X == 0 and t.BasisY.Y == 1 and t.BasisY.Z == 0 and
            t.BasisZ.X == 0 and t.BasisZ.Y == 0 and t.BasisZ.Z == 1)


def _points_to_source_coords(ctx, points):
    """Host -> link coordinates when zones live in a linked document (mirror of Write)."""
    if ctx.link_instance is None or not points:
        return list(points)
    try:
        t = ctx.link_instance.GetTotalTransform()
        if _is_identity_transform(t):
            return list(points)
        inv = t.Inverse
        return [inv.OfPoint(p) for p in points]
    except Exception as ex:
        logger.debug("Zone Audit: link transform failed: {}".format(ex))
        return list(points)


def _cell_of(pt):
    return (int(pt.X / CELL_SIZE_FEET), int(pt.Y / CELL_SIZE_FEET))


def _cells_around(pt):
    ix, iy = _cell_of(pt)
    cells = []
    for di in [-1, 0, 1]:
        for dj in [-1, 0, 1]:
            cells.append((ix + di, iy + dj))
    return cells


def _cells_of_bbox(bbox):
    if bbox is None:
        return []
    cells = []
    for ix in range(int(bbox.Min.X / CELL_SIZE_FEET), int(bbox.Max.X / CELL_SIZE_FEET) + 1):
        for iy in range(int(bbox.Min.Y / CELL_SIZE_FEET), int(bbox.Max.Y / CELL_SIZE_FEET) + 1):
            cells.append((ix, iy))
    return cells


# --------------------------------------------------------------------------------------
# Context: everything Write prepares before its per-element loop
# --------------------------------------------------------------------------------------

class AuditContext(object):
    """Holds the same prepared state as core.write_parameters_to_elements."""

    def __init__(self):
        self.doc = None
        self.config = None
        self.config_name = u""
        self.view_id = None
        self.view_name = u""
        self.strategy = None
        self.source_categories = []
        self.categories_for_containment = []
        self.source_param_names = []
        self.target_param_names = []
        self.target_filter_categories = []
        self.target_cat_ints = None          # None = all categories
        self.source_doc = None
        self.link_instance = None
        self.link_name = u""
        self.use_view_filter = False         # source side (Write skips it for links)
        self.visible_ids = None              # target side (set of int ids) when view_id given
        self.source_elements_all = []        # before the source-parameter filter
        self.source_without_params = []
        self.source_elements = []            # filtered + sorted, what Write uses
        self.source_ids = set()
        self.source_all_ids = set()
        self.sort_property = "ElementId"
        self.sort_descending = False
        self.ifc_export_only_empty = False
        self.is_workshared = False
        self.element_index = None
        self.cell_size = CELL_SIZE_FEET
        # room/space/area strategy state
        self.rooms_by_level = None
        self.spaces_by_level = None
        self.areas_by_level = None
        self.ordered_phases = None
        self.main_doc_ordered_phases = None
        self.rooms_by_phase_by_level = None
        self.phase_map = None
        self.source_coplanar_cache = None
        # bookkeeping
        self.notes = []
        self.timings = {}
        self.geometry_detail_level = u""
        self._zone_geometry = {}
        self.error = None

    def note(self, text):
        self.notes.append(_u(text))
        logger.debug(u"[ZoneAudit] {}".format(_u(text)))


def _collect_source_elements(ctx):
    """Mirror of the source-element collection in core.write_parameters_to_elements."""
    source_doc = ctx.source_doc
    view_id = ctx.view_id if ctx.use_view_filter else None
    source_elements = []
    seen = set()
    if ctx.source_categories:
        for category in ctx.source_categories:
            if category == THREE_D_ZONE_MARKER or str(category) == THREE_D_ZONE_MARKER:
                if view_id is not None:
                    coll = FilteredElementCollector(source_doc, view_id)
                else:
                    coll = FilteredElementCollector(source_doc)
                elements = coll.WhereElementIsNotElementType()\
                    .OfCategory(BuiltInCategory.OST_GenericModel).ToElements()
                for el in elements:
                    try:
                        if hasattr(el, "Symbol"):
                            symbol = el.Symbol
                            if symbol and hasattr(symbol, "FamilyName"):
                                family_name = symbol.FamilyName
                                if family_name and "3DZone" in family_name:
                                    v = _id(el)
                                    if v not in seen:
                                        seen.add(v)
                                        source_elements.append(el)
                    except Exception:
                        continue
            else:
                if view_id is not None:
                    coll = FilteredElementCollector(source_doc, view_id)
                else:
                    coll = FilteredElementCollector(source_doc)
                elements = coll.WhereElementIsNotElementType().OfCategory(category).ToElements()
                for el in elements:
                    v = _id(el)
                    if v not in seen:
                        seen.add(v)
                        source_elements.append(el)
    else:
        if view_id is not None:
            source_elements = list(FilteredElementCollector(source_doc, view_id)
                                   .WhereElementIsNotElementType().ToElements())
        else:
            source_elements = list(FilteredElementCollector(source_doc)
                                   .WhereElementIsNotElementType().ToElements())
    return source_elements


def build_context(doc, zone_config, view_id=None, progress=None):
    """Prepare the audit context exactly the way Write prepares a configuration.

    Args:
        doc: host document (where the targets live)
        zone_config: configuration dict from zone3d.config
        view_id: ElementId of the active view when 'active view only' is on, else None
        progress: optional callable(step_text)

    Returns:
        AuditContext
    """
    ctx = AuditContext()
    ctx.doc = doc
    ctx.config = zone_config
    ctx.config_name = _u(zone_config.get("name", "Unknown"))
    ctx.view_id = view_id
    if view_id is not None:
        ctx.view_name = _u(_safe(lambda: doc.GetElement(view_id).Name, u"?"))

    def step(text):
        if progress:
            try:
                progress(text)
            except Exception:
                pass

    t0 = time.time()
    try:
        # Same reload + cache reset as core.execute_configuration
        core._reload_containment_module()
        core._element_type_cache = {}
        containment.clear_geometry_cache()

        ctx.source_categories = list(zone_config.get("source_categories", []))
        categories_for_strategy = []
        for cat in ctx.source_categories:
            if cat == THREE_D_ZONE_MARKER or str(cat) == THREE_D_ZONE_MARKER:
                categories_for_strategy.append(BuiltInCategory.OST_GenericModel)
            else:
                categories_for_strategy.append(cat)
        ctx.categories_for_containment = categories_for_strategy
        strategy = containment.detect_containment_strategy(categories_for_strategy)
        if not strategy and categories_for_strategy:
            strategy = "overlap"
        ctx.strategy = strategy
        if not strategy:
            ctx.error = u"Could not detect containment strategy for {}".format(ctx.source_categories)
            return ctx

        ctx.source_param_names = list(zone_config.get("source_params", []))
        ctx.target_param_names = list(zone_config.get("target_params", []))
        ctx.sort_property = zone_config.get("source_sort_property", "ElementId")
        ctx.sort_descending = bool(zone_config.get("source_sort_descending", False))
        ctx.ifc_export_only_empty = bool(zone_config.get("ifc_export_only_empty", False))
        ctx.is_workshared = bool(doc.IsWorkshared)
        if not ctx.source_param_names:
            ctx.error = u"Configuration has no source parameters; Write returns immediately."
            return ctx
        if len(ctx.source_param_names) != len(ctx.target_param_names):
            ctx.error = u"Source and target parameter lists differ in length; Write returns immediately."
            return ctx

        step("Resolving source document")
        ctx.source_doc, ctx.link_instance = core.get_source_document(doc, zone_config)
        if ctx.link_instance is not None:
            ctx.link_name = _u(_safe(lambda: ctx.link_instance.Name, u"?"))
        ctx.use_view_filter = view_id is not None and ctx.link_instance is None
        try:
            ctx.geometry_detail_level = _u(containment._get_geometry_options(ctx.source_doc).DetailLevel)
        except Exception:
            ctx.geometry_detail_level = u"?"

        step("Collecting source elements (zones)")
        t = time.time()
        ctx.source_elements_all = _collect_source_elements(ctx)
        ctx.source_all_ids = set(_id(el) for el in ctx.source_elements_all)
        filtered = []
        for el in ctx.source_elements_all:
            if core.has_source_parameter(el, ctx.source_param_names):
                filtered.append(el)
            else:
                ctx.source_without_params.append(el)
        ctx.source_elements = core.sort_source_elements(filtered, ctx.sort_property, descending=ctx.sort_descending)
        ctx.source_ids = set(_id(el) for el in ctx.source_elements)
        ctx.timings["source_collect"] = time.time() - t
        ctx.note(u"Source elements found: {} (with a source parameter value: {}, without: {})".format(
            len(ctx.source_elements_all), len(ctx.source_elements), len(ctx.source_without_params)))

        if strategy in ["element", "area"]:
            step("Precomputing zone geometry")
            t = time.time()
            containment.precompute_geometries(ctx.source_elements, ctx.source_doc)
            ctx.timings["precompute"] = time.time() - t
        if strategy == "element":
            step("Building spatial index")
            t = time.time()
            ctx.element_index = containment.build_source_element_spatial_index(
                ctx.source_elements, ctx.source_doc, ctx.cell_size,
                sort_property=ctx.sort_property, sort_descending=ctx.sort_descending)
            ctx.timings["index"] = time.time() - t
            ctx.note(u"Spatial index: {} cells of {} ft".format(len(ctx.element_index), ctx.cell_size))

        # Target-side filters
        ctx.target_filter_categories = list(zone_config.get("target_filter_categories", []))
        if ctx.target_filter_categories:
            ints = set()
            for cat in ctx.target_filter_categories:
                v = containment._category_to_int(cat)
                if v is not None:
                    ints.add(v)
            ctx.target_cat_ints = ints
        if view_id is not None:
            step("Collecting elements visible in the active view")
            t = time.time()
            ctx.visible_ids = set(get_element_id_value(i) for i in
                                  FilteredElementCollector(doc, view_id).WhereElementIsNotElementType().ToElementIds())
            ctx.timings["visible"] = time.time() - t
            ctx.note(u"Active view only: {} elements visible in view '{}'".format(len(ctx.visible_ids), ctx.view_name))

        # Room / space / area state (mirror of core)
        if strategy == "room":
            source_ordered_phases = containment.get_ordered_phases(ctx.source_doc)
            main_ordered_phases = containment.get_ordered_phases(doc) if ctx.link_instance else source_ordered_phases
            rooms_list = [el for el in ctx.source_elements if isinstance(el, Room)]
            ctx.rooms_by_phase_by_level = containment.build_rooms_by_phase_and_level(rooms_list)
            ctx.rooms_by_level = defaultdict(list)
            for el in ctx.source_elements:
                if isinstance(el, Room) and el.LevelId:
                    ctx.rooms_by_level[el.LevelId].append(el)
            ctx.ordered_phases = source_ordered_phases
            ctx.main_doc_ordered_phases = main_ordered_phases
            ctx.phase_map = containment.get_phase_map_for_link(doc, ctx.link_instance) if ctx.link_instance else None
        elif strategy == "space":
            ctx.spaces_by_level = defaultdict(list)
            for el in ctx.source_elements:
                if isinstance(el, Space) and el.LevelId:
                    ctx.spaces_by_level[el.LevelId].append(el)
        elif strategy == "area":
            from Autodesk.Revit.DB import Area
            ctx.areas_by_level = defaultdict(list)
            for el in ctx.source_elements:
                if isinstance(el, Area):
                    level_id = None
                    if hasattr(el, "LevelId") and el.LevelId:
                        level_id = el.LevelId
                    elif hasattr(el, "get_Parameter"):
                        level_param = el.get_Parameter("Level")
                        if level_param:
                            level_id = level_param.AsElementId()
                    ctx.areas_by_level[level_id if level_id else None].append(el)

        if strategy == "overlap" and containment._source_uses_coplanar_overlap(ctx.categories_for_containment):
            ctx.source_coplanar_cache = containment.build_source_coplanar_descriptor_cache(
                ctx.source_elements, ctx.source_doc, ctx.link_instance)
    except Exception as ex:
        ctx.error = u"build_context failed: {}\n{}".format(ex, traceback.format_exc())
        logger.error(ctx.error)
    ctx.timings["total_context"] = time.time() - t0
    return ctx


def describe_config(ctx):
    """Multi-line text describing the configuration and the prepared state."""
    cfg = ctx.config or {}
    lines = []
    lines.append(u"Configuration: {} (order {}, enabled {})".format(
        ctx.config_name, cfg.get("order", "?"), cfg.get("enabled", "?")))
    lines.append(u"Strategy: {}".format(ctx.strategy))
    lines.append(u"Source categories: {}".format(u", ".join(_u(c) for c in ctx.source_categories) or u"(all)"))
    lines.append(u"Source parameters: {}".format(u", ".join(_u(p) for p in ctx.source_param_names)))
    lines.append(u"Target parameters: {}".format(u", ".join(_u(p) for p in ctx.target_param_names)))
    lines.append(u"Target categories: {}".format(
        u", ".join(_u(c) for c in ctx.target_filter_categories) if ctx.target_filter_categories else u"(all model elements with a target parameter)"))
    if ctx.link_instance is not None:
        lines.append(u"Zones come from linked document: {}".format(ctx.link_name))
    else:
        lines.append(u"Zones come from the host document ({})".format(
            _u(cfg.get("linked_document_name")) if cfg.get("use_linked_document") else u"no link configured"))
        if cfg.get("use_linked_document"):
            lines.append(u"  NOTE: config asks for link '{}' but it was not resolved; Write falls back to the host document".format(
                _u(cfg.get("linked_document_name"))))
    lines.append(u"Sort: {} {}".format(ctx.sort_property, u"descending" if ctx.sort_descending else u"ascending"))
    lines.append(u"Only empty targets: {}".format(ctx.ifc_export_only_empty))
    lines.append(u"Active view only: {}{}".format(
        ctx.view_id is not None, u" (view '{}')".format(ctx.view_name) if ctx.view_id is not None else u""))
    lines.append(u"Workshared: {} (ownership check {})".format(ctx.is_workshared, u"on" if ctx.is_workshared else u"skipped"))
    lines.append(u"Geometry detail level used for zone solids: {}".format(ctx.geometry_detail_level))
    for n in ctx.notes:
        lines.append(u"  - {}".format(n))
    if ctx.timings:
        lines.append(u"Context timings (s): {}".format(
            u", ".join(u"{} {:.2f}".format(k, v) for k, v in sorted(ctx.timings.items()))))
    if ctx.error:
        lines.append(u"ERROR: {}".format(ctx.error))
    return u"\n".join(lines)


# --------------------------------------------------------------------------------------
# Zone geometry (independent of the sample-point method)
# --------------------------------------------------------------------------------------

def get_zone_geometry(ctx, zone):
    """All solids of a zone (host coordinates) plus what Write actually cached for it."""
    zid = _id(zone)
    if zid in ctx._zone_geometry:
        return ctx._zone_geometry[zid]
    geom = {
        "id": zid,
        "solids_source": [],
        "solids_host": [],
        "solid_count": 0,
        "cached_solid_count": 0,
        "cached": False,
        "bbox_source": None,
        "bbox_host": None,
        "in_source_set": zid in ctx.source_ids,
        "in_source_set_before_param_filter": zid in ctx.source_all_ids,
        "cells": [],
        "error": None,
    }
    try:
        opts = Options()
        opts.ComputeReferences = False
        try:
            opts.DetailLevel = containment._get_geometry_options(ctx.source_doc).DetailLevel
        except Exception:
            pass
        ge = zone.get_Geometry(opts)
        solids = containment._collect_all_solids_from_geometry(ge) if ge else []
        geom["solids_source"] = solids
        geom["solid_count"] = len(solids)
        cached = containment._geometry_cache.get(zid)
        if cached is not None:
            geom["cached"] = True
            geom["cached_solid_count"] = len(cached.get("solids", []))
        bbox = zone.get_BoundingBox(None)
        geom["bbox_source"] = bbox
        if ctx.link_instance is not None:
            t = ctx.link_instance.GetTotalTransform()
            geom["solids_host"] = containment._transform_solids_with_transform(solids, t)
            geom["bbox_host"] = containment._transform_axis_aligned_bbox(bbox, t) if bbox else None
        else:
            geom["solids_host"] = solids
            geom["bbox_host"] = bbox
        geom["cells"] = _cells_of_bbox(bbox)
    except Exception as ex:
        geom["error"] = u"{}".format(ex)
    ctx._zone_geometry[zid] = geom
    return geom


def zone_status(ctx, zone):
    """(code, text) describing whether Write can use this zone as a source at all."""
    geom = get_zone_geometry(ctx, zone)
    zid = _id(zone)
    if not geom["in_source_set_before_param_filter"]:
        why = u"not collected as a source element"
        if ctx.use_view_filter:
            why += u" (active view only also filters zones; is it visible in '{}'?)".format(ctx.view_name)
        try:
            fam = zone.Symbol.FamilyName if hasattr(zone, "Symbol") and zone.Symbol else None
            if fam is not None and "3DZone" not in fam:
                why += u"; family name '{}' does not contain '3DZone'".format(fam)
        except Exception:
            pass
        if zone.Document != ctx.source_doc:
            why += u"; zone is in a different document than the configured source"
        return ("ZONE_NOT_IN_SOURCE_SET", why)
    if not geom["in_source_set"]:
        return ("ZONE_SOURCE_PARAM_EMPTY",
                u"none of the source parameters {} has a value on this zone, so Write drops it".format(ctx.source_param_names))
    if geom["solid_count"] == 0:
        return ("ZONE_NO_SOLID", u"no solid with volume found in the zone geometry")
    if geom["cached_solid_count"] < geom["solid_count"]:
        return ("ZONE_OK_PARTIAL_CACHE",
                u"zone has {} solids but Write's precompute cached only {} (only the first solid is tested)".format(
                    geom["solid_count"], geom["cached_solid_count"]))
    return ("ZONE_OK", u"zone is a valid source ({} solid(s), {} cells)".format(geom["solid_count"], len(geom["cells"])))


def collect_zone_candidates(ctx, zone):
    """Model elements whose bounding box touches the zone, with an independent solid test.

    Returns:
        (candidates, intersecting_ids, info) where info has counts and notes.
    """
    geom = get_zone_geometry(ctx, zone)
    info = {"bbox_hits": 0, "skipped_view_specific": 0, "skipped_non_model": 0,
            "skipped_zones": 0, "skipped_types": 0, "solid_test_ok": True, "notes": []}
    candidates = []
    if geom["bbox_host"] is None:
        info["notes"].append(u"zone has no bounding box; no candidates")
        return candidates, set(), info
    bbox = geom["bbox_host"]
    outline = Outline(bbox.Min, bbox.Max)
    coll = FilteredElementCollector(ctx.doc).WhereElementIsNotElementType()\
        .WherePasses(BoundingBoxIntersectsFilter(outline))
    zone_id = _id(zone)
    same_doc = (zone.Document == ctx.doc)
    for el in coll:
        try:
            info["bbox_hits"] += 1
            if same_doc and _id(el) == zone_id:
                continue
            if isinstance(el, ElementType):
                info["skipped_types"] += 1
                continue
            if _safe(lambda: el.ViewSpecific, False):
                info["skipped_view_specific"] += 1
                continue
            cat = el.Category
            if cat is None or cat.CategoryType != CategoryType.Model:
                info["skipped_non_model"] += 1
                continue
            if core.is_3dzone_family(el):
                info["skipped_zones"] += 1
                continue
            candidates.append(el)
        except Exception:
            continue

    intersecting = set()
    if candidates and geom["solids_host"]:
        try:
            ids = List[ElementId]([el.Id for el in candidates])
            for solid in geom["solids_host"]:
                try:
                    if solid is None or solid.Volume <= 1e-9:
                        continue
                    passed = FilteredElementCollector(ctx.doc, ids)\
                        .WherePasses(ElementIntersectsSolidFilter(solid)).ToElementIds()
                    for eid in passed:
                        intersecting.add(get_element_id_value(eid))
                except Exception as ex:
                    info["solid_test_ok"] = False
                    info["notes"].append(u"ElementIntersectsSolidFilter failed: {}".format(ex))
        except Exception as ex:
            info["solid_test_ok"] = False
            info["notes"].append(u"solid test setup failed: {}".format(ex))
    elif not geom["solids_host"]:
        info["solid_test_ok"] = False
        info["notes"].append(u"zone has no solids; geometry test skipped")
    return candidates, intersecting, info


# --------------------------------------------------------------------------------------
# Parameter dry run (mirror of core.copy_parameters / copy_parameter_value, no Set)
# --------------------------------------------------------------------------------------

def describe_target_params(element, target_param_names):
    """Per target parameter: where it lives (instance/type/missing), read-only, storage, value."""
    out = []
    et = element_type_of(element)
    for name in target_param_names:
        entry = {"name": name, "where": "missing", "readonly": None, "storage": u"", "value": None}
        try:
            p = element.LookupParameter(name)
            if p is not None:
                entry["where"] = "instance"
            elif et is not None:
                p = et.LookupParameter(name)
                if p is not None:
                    entry["where"] = "type"
            if p is not None:
                entry["readonly"] = bool(p.IsReadOnly)
                entry["storage"] = storage_type_name(p)
                entry["value"] = param_value_str(p)
        except Exception as ex:
            entry["where"] = u"error: {}".format(ex)
        out.append(entry)
    return out


def format_target_params(entries):
    parts = []
    for e in entries:
        if e["where"] == "missing":
            parts.append(u"{}: missing".format(e["name"]))
        else:
            parts.append(u"{}: {} {}{} = {}".format(
                e["name"], e["where"], e["storage"],
                u" READ-ONLY" if e["readonly"] else u"",
                u"<empty>" if e["value"] is None else u"'{}'".format(e["value"])))
    return u"; ".join(parts)


def simulate_copy_parameters(source_el, target_el, source_names, target_names):
    """Dry run of core.copy_parameters. Returns (summary_code, per_param_list).

    summary_code is MAPPED / ALREADY_CORRECT or the first failure code by priority.
    """
    results = []
    if len(source_names) != len(target_names):
        return "TRACE_ERROR", [{"outcome": "TRACE_ERROR", "detail": u"parameter lists differ in length"}]
    source_type = element_type_of(source_el)
    target_type = element_type_of(target_el)
    for sname, tname in zip(source_names, target_names):
        r = {"source": sname, "target": tname, "outcome": None, "source_where": None,
             "target_where": None, "source_value": None, "target_value": None,
             "source_storage": u"", "target_storage": u"", "detail": u""}
        try:
            sp = source_el.LookupParameter(sname)
            r["source_where"] = "instance" if sp is not None else None
            if sp is None and source_type is not None:
                sp = source_type.LookupParameter(sname)
                r["source_where"] = "type" if sp is not None else None
            if sp is None:
                r["outcome"] = "SOURCE_PARAM_MISSING"
                results.append(r)
                continue
            r["source_storage"] = storage_type_name(sp)
            r["source_value"] = param_value_str(sp)
            if not sp.HasValue:
                r["outcome"] = "SOURCE_PARAM_EMPTY"
                results.append(r)
                continue

            tp = target_el.LookupParameter(tname)
            r["target_where"] = "instance" if tp is not None else None
            if tp is None and target_type is not None:
                tp = target_type.LookupParameter(tname)
                r["target_where"] = "type" if tp is not None else None
            if tp is None:
                r["outcome"] = "TARGET_PARAM_MISSING"
                results.append(r)
                continue
            r["target_storage"] = storage_type_name(tp)
            r["target_value"] = param_value_str(tp)
            if tp.IsReadOnly:
                r["outcome"] = "TARGET_PARAM_READONLY"
                results.append(r)
                continue
            if sp.StorageType != tp.StorageType:
                r["outcome"] = "STORAGE_TYPE_MISMATCH"
                results.append(r)
                continue

            st = sp.StorageType
            if st == StorageType.String:
                value = sp.AsString()
                if not value:
                    # Write counts this as "copied" but never calls Set; nothing changes.
                    r["outcome"] = "SOURCE_PARAM_EMPTY"
                    r["detail"] = u"empty string on zone (Write counts it as copied but writes nothing)"
                elif tp.HasValue and tp.AsString() == value:
                    r["outcome"] = "ALREADY_CORRECT"
                else:
                    r["outcome"] = "WOULD_COPY"
            elif st == StorageType.Integer:
                value = sp.AsInteger()
                if tp.HasValue and tp.AsInteger() == value:
                    r["outcome"] = "ALREADY_CORRECT"
                else:
                    r["outcome"] = "WOULD_COPY"
            elif st == StorageType.Double:
                value = sp.AsDouble()
                if tp.HasValue and abs(tp.AsDouble() - value) < 1e-9:
                    r["outcome"] = "ALREADY_CORRECT"
                else:
                    r["outcome"] = "WOULD_COPY"
            elif st == StorageType.ElementId:
                value = sp.AsElementId()
                if not value or value == ElementId.InvalidElementId:
                    r["outcome"] = "SOURCE_PARAM_EMPTY"
                elif tp.HasValue and tp.AsElementId() == value:
                    r["outcome"] = "ALREADY_CORRECT"
                else:
                    r["outcome"] = "WOULD_COPY"
            else:
                r["outcome"] = "UNSUPPORTED_STORAGE"
            if r["outcome"] == "WOULD_COPY" and r["target_where"] == "type":
                r["detail"] = u"writes the TYPE parameter (all instances of the type change)"
        except Exception as ex:
            r["outcome"] = "TRACE_ERROR"
            r["detail"] = u"{}".format(ex)
        results.append(r)

    outcomes = [r["outcome"] for r in results]
    if "WOULD_COPY" in outcomes:
        return "MAPPED", results
    if "ALREADY_CORRECT" in outcomes:
        return "ALREADY_CORRECT", results
    for code in _PARAM_FAILURE_PRIORITY:
        if code in outcomes:
            return code, results
    if "TRACE_ERROR" in outcomes:
        return "TRACE_ERROR", results
    return "TRACE_ERROR", results


def format_param_results(results):
    parts = []
    for r in results:
        if r.get("outcome") == "TRACE_ERROR" and "source" not in r:
            parts.append(u"error: {}".format(r.get("detail")))
            continue
        sv = u"<empty>" if r.get("source_value") is None else u"'{}'".format(r.get("source_value"))
        tv = u"<empty>" if r.get("target_value") is None else u"'{}'".format(r.get("target_value"))
        txt = u"{} {} ({}) -> {} {} ({}) = {}".format(
            r.get("source"), sv, r.get("source_storage") or (r.get("source_where") or u"missing"),
            r.get("target"), tv, r.get("target_storage") or (r.get("target_where") or u"missing"),
            r.get("outcome"))
        if r.get("detail"):
            txt += u" [{}]".format(r.get("detail"))
        parts.append(txt)
    return u"; ".join(parts)


# --------------------------------------------------------------------------------------
# Containment tracing
# --------------------------------------------------------------------------------------

def _real_containment(ctx, element):
    """Call the containment function exactly as core.write_parameters_to_elements does."""
    if ctx.strategy == "room" and ctx.ordered_phases is not None and ctx.rooms_by_phase_by_level is not None:
        element_phases = ctx.main_doc_ordered_phases if ctx.link_instance else ctx.ordered_phases
        return containment.get_containing_room_phase_aware(
            element, ctx.source_doc, ctx.rooms_by_phase_by_level, ctx.ordered_phases,
            element_phases, ctx.link_instance, host_doc=ctx.doc, phase_map=ctx.phase_map)
    overlap_exclude_id = None
    if ctx.strategy == "overlap" and ctx.link_instance is None:
        overlap_exclude_id = _id(element)
    return containment.get_containing_element_by_strategy(
        element, ctx.source_doc, ctx.strategy, ctx.categories_for_containment,
        ctx.rooms_by_level, ctx.spaces_by_level, ctx.areas_by_level,
        ctx.element_index, ctx.cell_size,
        sort_property=ctx.sort_property, sort_descending=ctx.sort_descending,
        link_instance=ctx.link_instance, exclude_element_id=overlap_exclude_id,
        source_coplanar_cache=ctx.source_coplanar_cache)


def _diagnose_host_faces(element, doc):
    """Say where a HostObject's sample points came from and why face sampling may have failed."""
    parts = []
    try:
        face_pts = containment._get_host_object_test_points(element, doc)
        if face_pts:
            parts.append(u"HostObjectUtils face grid: {} points".format(len(face_pts)))
        else:
            geom_fn = getattr(containment, "_get_geometry_face_test_points", None)
            geom_pts = geom_fn(element, doc) if geom_fn else []
            if geom_pts:
                parts.append(u"HostObjectUtils gave no faces; geometry-face fallback: {} points".format(len(geom_pts)))
            else:
                parts.append(u"HostObjectUtils gave no faces and no geometry faces; LocationCurve fallback used")
        if hasattr(element, "WallType"):
            for side in (ShellLayerType.Exterior, ShellLayerType.Interior):
                try:
                    refs = HostObjectUtils.GetSideFaces(element, side)
                    refs = list(refs) if refs else []
                    n_planar = 0
                    for ref in refs:
                        try:
                            face = element.GetGeometryObjectFromReference(ref)
                        except Exception as fex:
                            parts.append(u"{}: GetGeometryObjectFromReference failed: {}".format(side, fex))
                            continue
                        if isinstance(face, PlanarFace):
                            n_planar += 1
                    parts.append(u"{} side: {} face refs, {} planar".format(side, len(refs), n_planar))
                except Exception as ex:
                    parts.append(u"{} side: GetSideFaces failed: {}".format(side, ex))
            try:
                parts.append(u"wall kind {}".format(element.WallType.Kind))
            except Exception:
                pass
    except Exception as ex:
        parts.append(u"diagnostic failed: {}".format(ex))
    return u"host faces: " + u"; ".join(parts)


def _trace_element_strategy(ctx, element, row):
    """Rebuild containment.get_containing_element_indexed with full detail.

    Fills row["points_host"], ["points_source"], ["use_vote"], ["cells"], ["candidates"],
    ["matrix"], ["votes"], ["vote_threshold"], ["zone_id"] and returns a reason code or
    None when a zone was found.
    """
    target_doc = element.Document
    use_vote = containment._is_3d_zone_vote_target(element)
    row["use_vote"] = use_vote
    n_body = 0
    n_tail = 0
    if use_vote:
        pts = containment._merge_3d_zone_vote_test_points(element, target_doc)
    else:
        pts, n_body, n_tail = containment.get_element_test_points(element, target_doc, return_body_count=True)
    pts = list(pts) if pts else []
    row["points_host"] = pts
    row["n_body_points"] = n_body
    row["n_tail_points"] = n_tail
    if isinstance(element, HostObject):
        row["gates"].append(_diagnose_host_faces(element, target_doc))
    if not pts:
        return "NO_TEST_POINTS"
    pts_src = _points_to_source_coords(ctx, pts)
    row["points_source"] = pts_src

    index = ctx.element_index or {}
    cells = []
    if use_vote:
        for pt in pts_src:
            cells.extend(_cells_around(pt))
    else:
        cells.extend(_cells_around(pts_src[0]))
    row["cells"] = cells
    all_candidates = []
    for cell in cells:
        if cell in index:
            all_candidates.extend(index[cell])
    if not all_candidates:
        return "INDEX_MISS"
    seen = set()
    candidates = []
    for el in all_candidates:
        v = _id(el)
        if v not in seen:
            seen.add(v)
            candidates.append(el)
    if ctx.sort_property == "ElementId":
        candidates.sort(key=lambda el: _id(el), reverse=ctx.sort_descending)
    else:
        candidates = core.sort_source_elements(candidates, ctx.sort_property, descending=ctx.sort_descending)
    row["candidates"] = candidates

    # inside matrix: candidate id -> list of (in_bbox, inside) per point
    matrix = {}
    for cand in candidates:
        cid = _id(cand)
        cached = containment._geometry_cache.get(cid)
        bbox = cached.get("bbox") if cached else None
        per_point = []
        for pt in pts_src:
            in_bbox = True
            if bbox is not None and not containment.is_point_in_bbox(pt, bbox):
                in_bbox = False
            inside = False
            if in_bbox:
                inside = bool(containment.is_point_in_element(cand, pt, ctx.source_doc))
            per_point.append((in_bbox, inside))
        matrix[cid] = per_point
    row["matrix"] = matrix

    if use_vote:
        n = len(pts_src)
        threshold = max(1, int(math.ceil(float(n) * float(containment.ROOF_CONTAINMENT_VOTE_MIN_FRACTION))))
        row["vote_threshold"] = threshold
        votes = defaultdict(int)
        first_hits = []
        for p_idx in range(n):
            hit = None
            for cand in candidates:
                if matrix[_id(cand)][p_idx][1]:
                    hit = cand
                    break
            first_hits.append(_id(hit) if hit is not None else None)
            if hit is not None:
                votes[_id(hit)] += 1
        row["votes"] = dict(votes)
        row["first_hits"] = first_hits
        if not votes:
            return "ALL_POINTS_MISS"
        max_votes = max(votes.values())
        if max_votes < threshold:
            return "VOTE_BELOW_THRESHOLD"
        tied = sorted([cid for cid in votes if votes[cid] == max_votes])
        row["zone_id"] = tied[0]
        return None

    # Pass 1 (point-based families): body points decide, point by point, zones in sort order
    for p_idx in range(min(n_body, len(pts_src))):
        for cand in candidates:
            if matrix[_id(cand)][p_idx][1]:
                row["zone_id"] = _id(cand)
                row["gates"].append(u"decided by body point #{} of {} (body points: solid centroid(s) first, bbox centre last)".format(
                    p_idx + 1, n_body))
                return None
    regular_end = len(pts_src) - n_tail
    # Pass 2: first zone in sort order that contains any regular point
    for cand in candidates:
        hits = matrix[_id(cand)]
        if any(h[1] for h in hits[n_body:regular_end]):
            row["zone_id"] = _id(cand)
            row["gates"].append(u"decided by sort order: first zone containing any regular point")
            return None
    # Pass 3: last resort, touching/facing points just outside the element
    for cand in candidates:
        hits = matrix[_id(cand)]
        if any(h[1] for h in hits[regular_end:]):
            row["zone_id"] = _id(cand)
            row["gates"].append(u"decided by a LAST-RESORT point (touching/facing): the body itself is in no zone, this zone is adjacent")
            return None
    # Pass 4: exact solid intersection (same Revit test the audit uses for "geometry intersects")
    cand_fn = getattr(containment, "_candidates_for_bbox", None)
    hit_fn = getattr(containment, "element_intersects_zone_solid", None)
    if cand_fn is not None and hit_fn is not None:
        pass4 = []
        try:
            for cand in cand_fn(element, ctx.element_index or {}, ctx.cell_size, ctx.link_instance,
                                ctx.sort_property, ctx.sort_descending):
                if hit_fn(element, cand, ctx.source_doc, ctx.link_instance):
                    pass4.append(_id(cand))
        except Exception as ex:
            row["gates"].append(u"pass 4 failed: {}".format(ex))
        row["pass4_hits"] = pass4
        if pass4:
            row["zone_id"] = pass4[0]
            row["gates"].append(u"decided by EXACT SOLID INTERSECTION (pass 4): no sample point inside any zone; "
                                u"zone solids intersecting the element, in sort order: {}".format(
                                    u", ".join(u"{}".format(z) for z in pass4)))
            return None
    return "ALL_POINTS_MISS"


def _trace_spatial_strategy(ctx, element, row):
    """Per-point detail for room/space/area strategies (result comes from the real call)."""
    doc_for_points = element.Document
    pts = list(containment.get_element_test_points(element, doc_for_points) or [])
    if containment._is_roof_element(element):
        fp = containment._get_roof_footprint_test_points(element, doc_for_points)
        if fp:
            pts.extend(fp)
    row["points_host"] = pts
    if not pts:
        return "NO_TEST_POINTS"
    pts_src = _points_to_source_coords(ctx, pts)
    row["points_source"] = pts_src

    if ctx.strategy == "room":
        candidates = [el for el in ctx.source_elements if isinstance(el, Room)]
        inside_fn = containment.is_point_in_room
    elif ctx.strategy == "space":
        candidates = [el for el in ctx.source_elements if isinstance(el, Space)]
        inside_fn = containment.is_point_in_space
    else:
        candidates = list(ctx.source_elements)
        inside_fn = lambda area, pt: containment.is_point_in_area(area, pt, ctx.source_doc)

    # bbox pre-filter (0.5 ft, as the containment module does)
    ebox = element.get_BoundingBox(None)
    if ebox is not None and ctx.link_instance is not None:
        try:
            ebox = containment._transform_axis_aligned_bbox(ebox, ctx.link_instance.GetTotalTransform().Inverse)
        except Exception:
            pass
    filtered = []
    for cand in candidates:
        cb = cand.get_BoundingBox(None)
        if ebox is None or cb is None:
            filtered.append(cand)
            continue
        if (ebox.Min.X - 0.5 <= cb.Max.X and ebox.Max.X + 0.5 >= cb.Min.X and
                ebox.Min.Y - 0.5 <= cb.Max.Y and ebox.Max.Y + 0.5 >= cb.Min.Y and
                ebox.Min.Z - 0.5 <= cb.Max.Z and ebox.Max.Z + 0.5 >= cb.Min.Z):
            filtered.append(cand)
    filtered.sort(key=lambda el: _id(el))
    row["candidates"] = filtered
    matrix = {}
    any_inside = False
    for cand in filtered:
        per_point = []
        for pt in pts_src:
            inside = bool(inside_fn(cand, pt))
            any_inside = any_inside or inside
            per_point.append((True, inside))
        matrix[_id(cand)] = per_point
    row["matrix"] = matrix
    if not any_inside:
        return "ALL_POINTS_MISS"
    return None


def trace_element(ctx, element, audited_zone=None):
    """Full decision trail for one element. Returns a row dict (see keys below)."""
    row = {
        "element": element,
        "info": describe_element(element),
        "reason": None,
        "detail": u"",
        "gates": [],
        "zone_id": None,            # zone chosen by the tracer's reconstruction
        "write_zone_id": None,      # zone returned by the real containment call
        "tracer_mismatch": False,
        "points_host": [],
        "points_source": [],
        "use_vote": False,
        "n_body_points": 0,
        "n_tail_points": 0,
        "cells": [],
        "candidates": [],
        "matrix": {},
        "votes": {},
        "first_hits": [],
        "vote_threshold": None,
        "target_params": [],
        "param_results": [],
        "bbox": _safe(lambda: element.get_BoundingBox(None)),
        "audited_zone_id": _id(audited_zone) if audited_zone is not None else None,
        "geometry_intersects": None,
        "bucket": None,
    }
    try:
        eid = _id(element)
        # Gate 1: target set (categories, active view)
        cat = element.Category
        cat_int = get_element_id_value(cat.Id) if cat is not None and cat.Id else None
        if isinstance(element, ElementType):
            row["reason"] = "NOT_TARGET_CATEGORY"
            row["detail"] = u"element types are never targets"
            return row
        if ctx.target_cat_ints is not None and cat_int not in ctx.target_cat_ints:
            row["reason"] = "NOT_TARGET_CATEGORY"
            row["detail"] = u"category '{}' not in target categories".format(row["info"]["category"])
            return row
        row["gates"].append(u"category OK")
        if ctx.visible_ids is not None and eid not in ctx.visible_ids:
            row["reason"] = "NOT_IN_VIEW"
            row["detail"] = u"not visible in view '{}' (active view only)".format(ctx.view_name)
            return row
        if ctx.visible_ids is not None:
            row["gates"].append(u"visible in view OK")

        # Gate 2: room strategy exclusions
        if ctx.strategy == "room" and (isinstance(element, Room) or core.is_3dzone_family(element)):
            row["reason"] = "EXCLUDED_ROOM_OR_ZONE"
            return row

        # Gate 3: target parameters
        row["target_params"] = describe_target_params(element, ctx.target_param_names)
        if not core.has_target_parameter(element, ctx.target_param_names):
            row["reason"] = "NO_TARGET_PARAM"
            row["detail"] = format_target_params(row["target_params"])
            return row
        row["gates"].append(u"target params: {}".format(format_target_params(row["target_params"])))

        # Gate 4: ownership
        if ctx.is_workshared:
            editable, why = is_element_editable(ctx.doc, element)
            if not editable:
                row["reason"] = "NOT_EDITABLE"
                row["detail"] = _u(why)
                return row
            row["gates"].append(u"editable OK")

        # Gate 5: only-empty. Write stops here; the tracer continues so the log can say
        # whether the value already on the element matches the zone it sits in.
        skipped_only_empty = False
        if ctx.ifc_export_only_empty and not core.are_target_parameters_empty(element, ctx.target_param_names):
            skipped_only_empty = True
            row["gates"].append(u"only-empty: target already has a value, Write skips this element")

        # Gate 6: containment
        if ctx.strategy == "element":
            reason = _trace_element_strategy(ctx, element, row)
        elif ctx.strategy in ("room", "space", "area"):
            reason = _trace_spatial_strategy(ctx, element, row)
        else:
            reason = None
            row["detail"] = u"overlap strategy: solid-intersection based, no per-point detail"

        real = None
        try:
            real = _real_containment(ctx, element)
        except Exception as ex:
            row["gates"].append(u"real containment call raised: {}".format(ex))
        row["write_zone_id"] = _id(real) if real is not None else None

        if ctx.strategy == "element":
            if row["zone_id"] != row["write_zone_id"]:
                row["tracer_mismatch"] = True
                row["gates"].append(u"TRACER MISMATCH: reconstruction says {}, Write says {}".format(
                    row["zone_id"], row["write_zone_id"]))
            # Trust the real result for the verdict
            zone_id = row["write_zone_id"]
        else:
            zone_id = row["write_zone_id"]
            row["zone_id"] = zone_id
            if zone_id is None and reason is None:
                # points hit a spatial element but the real function said no
                reason = "PHASE_MISMATCH" if ctx.strategy == "room" else "NOT_CONTAINED"

        if zone_id is None:
            if skipped_only_empty:
                row["reason"] = "TARGET_NOT_EMPTY"
                row["detail"] = u"{}; and no zone found either ({}: {})".format(
                    format_target_params(row["target_params"]), reason or "NOT_CONTAINED", _containment_detail(ctx, row))
            else:
                row["reason"] = reason or "NOT_CONTAINED"
                row["detail"] = _containment_detail(ctx, row)
            return row

        containing = real
        # Gate 7: 3DZone family name filter (mirror of core)
        if THREE_D_ZONE_MARKER in ctx.source_categories:
            fam = _safe(lambda: containing.Symbol.FamilyName, None) if hasattr(containing, "Symbol") else None
            if not fam or "3DZone" not in fam:
                row["reason"] = "ZONE_NAME_FILTER"
                row["detail"] = u"containing element {} family '{}'".format(zone_id, fam)
                return row
        row["gates"].append(u"contained by zone {}".format(zone_id))

        # Gate 8: is it the audited zone?
        if audited_zone is not None and zone_id != _id(audited_zone):
            summary, results = simulate_copy_parameters(containing, element, ctx.source_param_names, ctx.target_param_names)
            row["param_results"] = results
            row["reason"] = "CONTAINED_OTHER_ZONE"
            row["detail"] = u"mapped to zone {} ({}); {}".format(zone_id, zone_label(ctx, containing), _containment_detail(ctx, row))
            if skipped_only_empty:
                row["detail"] = u"only-empty skip, existing value {} that zone; ".format(
                    u"matches" if summary == "ALREADY_CORRECT" else u"differs from") + row["detail"]
            return row

        # Gate 9: parameter dry run
        summary, results = simulate_copy_parameters(containing, element, ctx.source_param_names, ctx.target_param_names)
        row["param_results"] = results
        if skipped_only_empty:
            row["reason"] = "TARGET_NOT_EMPTY_MATCHES" if summary == "ALREADY_CORRECT" else "TARGET_NOT_EMPTY"
            row["detail"] = u"Write skips it (only-empty); {}".format(format_param_results(results))
        else:
            row["reason"] = summary
            row["detail"] = format_param_results(results)
        return row
    except Exception as ex:
        row["reason"] = "TRACE_ERROR"
        row["detail"] = u"{}\n{}".format(ex, traceback.format_exc())
        return row


def _containment_detail(ctx, row):
    """One-line summary of the containment step for the CSV 'detail' column."""
    parts = []
    n_pts = len(row.get("points_host") or [])
    parts.append(u"{} sample points".format(n_pts))
    if row.get("use_vote"):
        parts.append(u"vote threshold {}".format(row.get("vote_threshold")))
        if row.get("votes"):
            parts.append(u"votes " + u", ".join(u"{}:{}".format(k, v) for k, v in sorted(row["votes"].items())))
    cands = row.get("candidates") or []
    parts.append(u"{} candidate zones".format(len(cands)))
    matrix = row.get("matrix") or {}
    audited = row.get("audited_zone_id")
    if audited is not None and audited in matrix:
        hits = matrix[audited]
        parts.append(u"audited zone: {}/{} points in bbox, {}/{} inside solid".format(
            sum(1 for h in hits if h[0]), len(hits), sum(1 for h in hits if h[1]), len(hits)))
    elif audited is not None and cands:
        parts.append(u"audited zone not among candidates")
    hit_zones = [cid for cid, hits in matrix.items() if any(h[1] for h in hits)]
    if hit_zones:
        parts.append(u"zones hit by any point: {}".format(u", ".join(u"{}".format(z) for z in sorted(hit_zones))))
    return u"; ".join(parts)


# --------------------------------------------------------------------------------------
# Zone audit and buckets
# --------------------------------------------------------------------------------------

def classify_bucket(row):
    reason = row.get("reason")
    audited = row.get("audited_zone_id")
    mapped_here = reason in MAPPED_REASONS and (audited is None or row.get("write_zone_id") == audited)
    gi = row.get("geometry_intersects")
    if mapped_here:
        if gi is False:
            return "MAPPED_HERE_NO_GEOM"
        return "MAPPED_HERE"
    if reason == "CONTAINED_OTHER_ZONE":
        if gi is True:
            return "MAPPED_OTHER"
        return "OUTSIDE" if gi is False else "UNVERIFIED"
    if gi is True:
        return "MISSED"
    if gi is False:
        return "OUTSIDE"
    return "UNVERIFIED"


def audit_zone(ctx, zone, progress=None):
    """Audit one zone. Returns (zone_report, rows)."""
    t0 = time.time()
    report = {
        "zone": zone,
        "zone_id": _id(zone),
        "label": zone_label(ctx, zone),
        "status": zone_status(ctx, zone),
        "geometry": get_zone_geometry(ctx, zone),
        "candidate_info": None,
        "bucket_counts": {},
        "reason_counts": {},
        "n_candidates": 0,
        "n_intersecting": 0,
        "seconds": 0.0,
    }
    candidates, intersecting, info = collect_zone_candidates(ctx, zone)
    report["candidate_info"] = info
    report["n_candidates"] = len(candidates)
    report["n_intersecting"] = len(intersecting)
    rows = []
    total = len(candidates)
    for idx, el in enumerate(candidates):
        if progress:
            try:
                if progress(idx + 1, total):
                    break
            except Exception:
                pass
        row = trace_element(ctx, el, audited_zone=zone)
        if info["solid_test_ok"]:
            row["geometry_intersects"] = _id(el) in intersecting
        else:
            row["geometry_intersects"] = None
        row["bucket"] = classify_bucket(row)
        rows.append(row)
    counts = defaultdict(int)
    rcounts = defaultdict(int)
    for row in rows:
        counts[row["bucket"]] += 1
        rcounts[row["reason"]] += 1
    report["bucket_counts"] = dict(counts)
    report["reason_counts"] = dict(rcounts)
    report["seconds"] = time.time() - t0
    return report, rows


def find_zones(ctx):
    """All 3DZone-like source elements Write would consider (before the parameter filter)."""
    return list(ctx.source_elements_all)


# --------------------------------------------------------------------------------------
# Output: CSV, text trace, log folder
# --------------------------------------------------------------------------------------

CSV_COLUMNS = [
    "zone_id", "zone_label", "element_id", "category", "family", "type", "name", "level", "workset",
    "bucket", "reason", "reason_description", "detail", "write_zone_id", "tracer_zone_id",
    "tracer_mismatch", "geometry_intersects", "n_points", "n_points_inside_audited_zone",
    "n_candidates", "votes", "vote_threshold", "target_params", "param_results",
    "bbox_min_mm", "bbox_max_mm", "config", "strategy",
]


def default_log_dir(doc):
    """<model folder>\\zone-audit-logs when the model is a local file, else %APPDATA%."""
    candidates = []
    try:
        path = doc.PathName
        if path:
            folder = op.dirname(path)
            if folder and op.isdir(folder):
                candidates.append(op.join(folder, "zone-audit-logs"))
    except Exception:
        pass
    appdata = os.environ.get("APPDATA") or op.expanduser("~")
    candidates.append(op.join(appdata, "pyByggstyrning", "logs", "zone-audit"))
    for folder in candidates:
        try:
            if not op.isdir(folder):
                os.makedirs(folder)
            probe = op.join(folder, ".write-test")
            with open(probe, "w") as f:
                f.write("ok")
            os.remove(probe)
            return folder
        except Exception:
            continue
    return appdata


def log_basename(doc, prefix="zone-audit"):
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    model = u"model"
    try:
        if doc.Title:
            model = _u(doc.Title)
    except Exception:
        pass
    safe = u"".join(ch if (ch.isalnum() or ch in u"-_.") else u"_" for ch in model)
    return u"{}_{}_{}".format(prefix, safe, stamp)


def _csv_cell(value):
    text = _u(value)
    text = text.replace(u"\r", u" ").replace(u"\n", u" | ").replace(u";", u",")
    return text


def row_to_csv_record(ctx, row, zone_report=None):
    info = row["info"]
    bbox = row.get("bbox")
    matrix = row.get("matrix") or {}
    audited = row.get("audited_zone_id")
    n_inside = u""
    if audited is not None and audited in matrix:
        n_inside = sum(1 for h in matrix[audited] if h[1])
    votes = row.get("votes") or {}
    return [
        zone_report["zone_id"] if zone_report else (row.get("write_zone_id") or u""),
        zone_report["label"] if zone_report else u"",
        info["id"], info["category"], info["family"], info["type"], info["name"], info["level"], info["workset"],
        row.get("bucket") or u"", row.get("reason") or u"", REASON_DESCRIPTIONS.get(row.get("reason"), u""),
        row.get("detail") or u"",
        row.get("write_zone_id") if row.get("write_zone_id") is not None else u"",
        row.get("zone_id") if row.get("zone_id") is not None else u"",
        row.get("tracer_mismatch"),
        u"" if row.get("geometry_intersects") is None else row.get("geometry_intersects"),
        len(row.get("points_host") or []), n_inside, len(row.get("candidates") or []),
        u", ".join(u"{}:{}".format(k, v) for k, v in sorted(votes.items())),
        row.get("vote_threshold") if row.get("vote_threshold") is not None else u"",
        format_target_params(row.get("target_params") or []),
        format_param_results(row.get("param_results") or []),
        fmt_pt_short(bbox.Min) if bbox else u"", fmt_pt_short(bbox.Max) if bbox else u"",
        ctx.config_name, ctx.strategy,
    ]


def write_csv(path, records):
    """Semicolon-separated UTF-8 CSV with BOM (opens directly in Swedish Excel)."""
    with codecs.open(path, "w", encoding="utf-8") as f:
        f.write(u"\ufeff")
        f.write(u";".join(CSV_COLUMNS) + u"\r\n")
        for rec in records:
            f.write(u";".join(_csv_cell(v) for v in rec) + u"\r\n")
    return path


def format_trace_block(ctx, row, verbose=True):
    """Human readable trace for one element (used for the text log and the output window)."""
    info = row["info"]
    lines = []
    lines.append(u"=== Element {} | {} | {} : {} | {} | Level {} | Workset {} ===".format(
        info["id"], info["category"], info["family"], info["type"], info["name"], info["level"], info["workset"]))
    if row.get("bucket"):
        lines.append(u"Bucket: {} - {}".format(row["bucket"], BUCKET_DESCRIPTIONS.get(row["bucket"], u"")))
    lines.append(u"Reason: {} - {}".format(row["reason"], REASON_DESCRIPTIONS.get(row["reason"], u"")))
    if row.get("detail"):
        lines.append(u"Detail: {}".format(row["detail"]))
    if row.get("geometry_intersects") is not None:
        lines.append(u"Solid intersects audited zone (ElementIntersectsSolidFilter): {}".format(
            u"YES" if row["geometry_intersects"] else u"no"))
    if row.get("bbox") is not None:
        lines.append(u"Bounding box: {}".format(fmt_bbox(row["bbox"])))
    if row.get("gates"):
        lines.append(u"Gates passed: " + u" | ".join(row["gates"]))
    if row.get("target_params") and row["reason"] in ("NO_TARGET_PARAM", "TARGET_NOT_EMPTY", "TARGET_NOT_EMPTY_MATCHES"):
        lines.append(u"Target parameters: {}".format(format_target_params(row["target_params"])))
    pts = row.get("points_host") or []
    if pts:
        if row.get("use_vote"):
            rule = u"floor/roof vote, threshold {}".format(row.get("vote_threshold"))
        elif row.get("n_body_points") or row.get("n_tail_points"):
            rule = u"{} body points [B] decide first; then first zone in sort order with any regular [R] hit; {} touching/facing [T] points only when nothing else hits".format(
                row.get("n_body_points", 0), row.get("n_tail_points", 0))
        else:
            rule = u"first zone in sort order with any hit"
        lines.append(u"Sample points: {} ({})".format(len(pts), rule))
        if verbose:
            matrix = row.get("matrix") or {}
            pts_src = row.get("points_source") or pts
            n_b = row.get("n_body_points", 0) or 0
            n_t = row.get("n_tail_points", 0) or 0
            for i, pt in enumerate(pts):
                in_bbox = [cid for cid, hits in matrix.items() if i < len(hits) and hits[i][0]]
                inside = [cid for cid, hits in matrix.items() if i < len(hits) and hits[i][1]]
                extra = u""
                if ctx.link_instance is not None and i < len(pts_src):
                    extra = u" -> link coords {}".format(fmt_pt_short(pts_src[i]))
                tier = u"B" if i < n_b else (u"T" if i >= len(pts) - n_t else u"R")
                lines.append(u"  #{:<3}{} {}{}  bbox-hit: {}  inside: {}".format(
                    i + 1, tier, fmt_pt(pt), extra,
                    ids_str(in_bbox) if in_bbox else u"-", ids_str(inside) if inside else u"-"))
    if row.get("cells"):
        uniq = sorted(set(row["cells"]))
        lines.append(u"Index cells looked up ({} ft): {}".format(ctx.cell_size, uniq if len(uniq) <= 12 else u"{} cells".format(len(uniq))))
    cands = row.get("candidates") or []
    if cands:
        parts = []
        matrix = row.get("matrix") or {}
        for cand in cands:
            hits = matrix.get(_id(cand), [])
            parts.append(u"{} ({}/{} in bbox, {}/{} inside)".format(
                _id(cand), sum(1 for h in hits if h[0]), len(hits), sum(1 for h in hits if h[1]), len(hits)))
        lines.append(u"Candidate zones in order: " + u", ".join(parts))
    if row.get("votes"):
        lines.append(u"Votes: " + u", ".join(u"zone {}: {}".format(k, v) for k, v in sorted(row["votes"].items())))
    lines.append(u"Write result (real containment call): {}{}".format(
        row.get("write_zone_id") if row.get("write_zone_id") is not None else u"none",
        u"  ** TRACER MISMATCH (reconstruction said {}) **".format(row.get("zone_id")) if row.get("tracer_mismatch") else u""))
    if row.get("param_results"):
        lines.append(u"Parameter dry run: {}".format(format_param_results(row["param_results"])))
    return u"\n".join(lines)


def write_trace_txt(path, ctx, header_lines, blocks):
    with codecs.open(path, "w", encoding="utf-8") as f:
        f.write(u"\ufeff")
        f.write(u"Zone Audit trace - {}\n".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        f.write(describe_config(ctx) + u"\n\n")
        for line in header_lines:
            f.write(_u(line) + u"\n")
        f.write(u"\n")
        for block in blocks:
            f.write(_u(block) + u"\n\n")
    return path


def _esc(text):
    text = _u(text)
    return text.replace(u"&", u"&amp;").replace(u"<", u"&lt;").replace(u">", u"&gt;")


def run_element_trace(doc, zone_config, elements, view_id=None, output=None, verbose=True):
    """Trace the given elements against one configuration and write CSV + TXT logs.

    This is the "why is this element (not) mapped?" mode used by the Write Mappings
    dialog. Read-only. Prints to the pyRevit output window when ``output`` is given.

    Returns:
        dict: {"csv": path, "txt": path, "totals": {reason: count}, "ctx": AuditContext, "rows": [...]}
    """
    ctx = build_context(doc, zone_config, view_id=view_id)
    if output is not None:
        output.print_md(u"## Configuration: {}".format(_esc(ctx.config_name)))
        output.print_html(u"<pre style='white-space:pre-wrap;font-size:11px'>{}</pre>".format(_esc(describe_config(ctx))))
    if ctx.error:
        return {"csv": None, "txt": None, "totals": {}, "ctx": ctx, "rows": [], "error": ctx.error}

    log_dir = default_log_dir(doc)
    base = log_basename(doc, prefix="element-trace")
    csv_path = op.join(log_dir, base + u".csv")
    txt_path = op.join(log_dir, base + u".txt")
    records = []
    blocks = []
    header_lines = []
    totals = {}
    rows = []
    for el in elements:
        row = trace_element(ctx, el)
        hits = zones_intersecting_element(ctx, el)
        zone_lines = []
        intersecting_ids = []
        for zone, status, intersects, zbox in hits:
            zone_lines.append(u"  zone {} | {} | status {} | bbox {} | solid intersects element: {}".format(
                _id(zone), zone_label(ctx, zone), status, fmt_bbox(zbox),
                u"YES" if intersects else (u"no" if intersects is False else u"test failed")))
            if intersects:
                intersecting_ids.append(_id(zone))
        if row.get("write_zone_id") is not None and intersecting_ids and row["write_zone_id"] not in intersecting_ids:
            zone_lines.append(u"  NOTE: Write maps this element to zone {} whose solid does not intersect it".format(row["write_zone_id"]))
        if row.get("write_zone_id") is None and intersecting_ids:
            zone_lines.append(u"  => geometrically inside zone(s) {} but NOT mapped: reason {}".format(
                ids_str(intersecting_ids), row["reason"]))
        row["geometry_intersects"] = True if intersecting_ids else (False if hits else None)
        row["bucket"] = classify_bucket(row)
        block = format_trace_block(ctx, row, verbose=verbose)
        block += u"\nZones whose bounding box overlaps this element ({}):\n{}".format(
            len(hits), u"\n".join(zone_lines) if zone_lines else u"  (none)")
        if output is not None:
            output.print_md(u"### Element {}".format(_esc(describe_element_short(el))))
            output.print_md(u"Select: {}".format(output.linkify(el.Id)))
            output.print_html(u"<pre style='white-space:pre-wrap;font-size:11px'>{}</pre>".format(_esc(block)))
        blocks.append(block)
        rows.append(row)
        records.append(row_to_csv_record(ctx, row, None))
        totals[row["reason"]] = totals.get(row["reason"], 0) + 1
        header_lines.append(u"Element {} | reason {} | write zone {} | intersecting zones {}".format(
            _id(el), row["reason"], row.get("write_zone_id"), ids_str(intersecting_ids)))
    write_csv(csv_path, records)
    write_trace_txt(txt_path, ctx, header_lines, blocks)
    if output is not None:
        output.print_md(u"CSV: `{}`".format(csv_path))
        output.print_md(u"TXT trace: `{}`".format(txt_path))
    return {"csv": csv_path, "txt": txt_path, "totals": totals, "ctx": ctx, "rows": rows}


def reasons_help_text():
    lines = [u"Reason codes (in gate order):"]
    for code, desc in REASONS:
        lines.append(u"  {:<22} {}".format(code, desc))
    lines.append(u"Buckets:")
    for code, desc in BUCKETS:
        lines.append(u"  {:<22} {}".format(code, desc))
    return u"\n".join(lines)


def zones_intersecting_element(ctx, element):
    """Independent check for element-trace mode: which zones' solids intersect this element.

    Uses ElementIntersectsSolidFilter with each zone solid (host coordinates) on a
    collector limited to the element itself. Zones are pre-filtered by bounding box.

    Returns:
        list of (zone, status_code, intersects_bool_or_None)
    """
    results = []
    ebox = _safe(lambda: element.get_BoundingBox(None))
    if ebox is None:
        return results
    try:
        ids = List[ElementId]([element.Id])
    except Exception:
        return results
    for zone in ctx.source_elements_all:
        geom = get_zone_geometry(ctx, zone)
        zb = geom.get("bbox_host")
        if zb is None:
            continue
        if not (ebox.Min.X <= zb.Max.X and ebox.Max.X >= zb.Min.X and
                ebox.Min.Y <= zb.Max.Y and ebox.Max.Y >= zb.Min.Y and
                ebox.Min.Z <= zb.Max.Z and ebox.Max.Z >= zb.Min.Z):
            continue
        status = zone_status(ctx, zone)[0]
        intersects = None
        for solid in geom.get("solids_host") or []:
            try:
                if solid is None or solid.Volume <= 1e-9:
                    continue
                passed = FilteredElementCollector(ctx.doc, ids)\
                    .WherePasses(ElementIntersectsSolidFilter(solid)).ToElementIds()
                if intersects is None:
                    intersects = False
                if len(list(passed)) > 0:
                    intersects = True
                    break
            except Exception as ex:
                logger.debug("Zone Audit: solid test failed for zone {}: {}".format(_id(zone), ex))
        results.append((zone, status, intersects, zb))
    return results
