# -*- coding: utf-8 -*-
"""Detect clashes between Revit categories and place isolated 3D views on a sheet."""

__title__ = "Clash\nViews"
__author__ = "pyByggstyrning"
__doc__ = """Pick two or more model categories. All unique pairs (C(N,2)) are clash-detected
using a bounding-box pre-filter plus geometric confirmation in the host document.
Each pair (sub-grouped by level) becomes an isolated 3D view; every view is placed
end-to-end on a single sheet in a grid layout.

Optionally clash against a linked model: enable "Against link model", pick a link,
select link categories. Pairs are then host-category x link-category (cross-product).
"Combined" merges those pairs into one view per host category.
Cross-doc clash uses bbox prefilter + solid-solid intersection (link solids transformed
to host coordinate space for geometric confirmation)."""

import clr
clr.AddReference("System")
clr.AddReference("System.Collections")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")

from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection

from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementId, BuiltInCategory, View,
    ViewFamilyType, ViewFamily, View3D, ViewSheet, Viewport,
    XYZ, BoundingBoxXYZ, Outline,
    BoundingBoxIntersectsFilter, ElementIntersectsElementFilter,
    UnitUtils, Transaction, TransactionGroup, BuiltInParameter,
    OverrideGraphicSettings, FillPatternElement,
    Color as RevitColor,
    RevitLinkInstance, RevitLinkType,
    Category, CategoryType,
    Reference,
    BooleanOperationsUtils, BooleanOperationsType, Solid, SolidUtils,
    Options, GeometryInstance, ViewDetailLevel,
    WorksetVisibility, FilteredWorksetCollector, WorksetKind,
)
from Autodesk.Revit.UI import RevitCommandId, PostableCommand

import sys
import os.path as op
from collections import OrderedDict, namedtuple
from itertools import combinations

from pyrevit import revit, DB
from pyrevit import forms, script

script_path = __file__
pushbutton_dir = op.dirname(script_path)
stack_dir = op.dirname(pushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from revit.compat import get_element_id_value, make_element_id

try:
    from revit.clash_markers import (
        is_temporary_graphics_available,
        find_marker_driver,
        start_or_get_driver,
        clean_marker_session,
        _is_view_sheet,
        register_marker_session,
        set_marker_toggle_active,
        is_marker_toggle_active,
    )
    _CLASH_MARKERS_AVAILABLE = is_temporary_graphics_available()
except Exception:
    _CLASH_MARKERS_AVAILABLE = False
    find_marker_driver = None
    start_or_get_driver = None
    clean_marker_session = None
    _is_view_sheet = None
    register_marker_session = None
    set_marker_toggle_active = None
    is_marker_toggle_active = None

logger = script.get_logger()

DEBUG_MODE = False

# Color clash sides via SetCategoryOverrides (link B uses host VG inheritance).
USE_CATEGORY_LINK_OVERRIDES = True

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# =============================================================================
# Category discovery
# =============================================================================

# Documentation / datum / annotation categories that must never appear in clash lists.
# CategoryType.Model is not enough: Sheets, Project Information, Sheet Collections,
# RVT Links, etc. can still report as model and have collector "instances".
_NON_GEOMETRIC_CATEGORIES = {
    # Sheets / views / documentation
    "OST_Sheets", "OST_SheetCollections", "OST_Views", "OST_Viewports",
    "OST_TitleBlocks", "OST_Schedules", "OST_ScheduleGraphics",
    "OST_RasterImages", "OST_DrawingList", "OST_LegendComponents",
    "OST_ColorFillLegends", "OST_RevisionClouds", "OST_Revisions",
    "OST_RevisionCloudTags", "OST_ColorFillSchema",
    # Project / internal
    "OST_ProjectInformation", "OST_Materials", "OST_Phases",
    "OST_DesignOptions", "OST_DesignOptionSets",
    "OST_RvtLinks", "OST_ImportObject", "OST_Coordination_Model",
    "OST_PointClouds", "OST_DecalElement",
    # Datum / cameras / views
    "OST_Grids", "OST_Levels", "OST_SectionBox", "OST_Viewers",
    "OST_CLines", "OST_Cameras", "OST_Sections", "OST_Elev",
    "OST_Matchline", "OST_ReferenceLines", "OST_ReferencePlanes",
    "OST_SketchLines", "OST_CenterLines", "OST_WorkPlaneGrid",
    "OST_ProjectBasePoint", "OST_SharedBasePoint", "OST_IOSSitePoint",
    "OST_VolumeOfInterest", "OST_ScopeBoxes",
    # Site property lines (not clash solids)
    "OST_SiteProperty", "OST_SitePropertyLineSegment",
    # Spatial placeholders (no solid geometry for intersection)
    "OST_Rooms", "OST_MEPSpaces", "OST_Areas", "OST_AreaSchemes",
    "OST_HVAC_Zones",
    # Annotation (belt-and-suspenders if CategoryType check fails)
    "OST_Dimensions", "OST_TextNotes", "OST_Tags", "OST_GenericAnnotation",
    "OST_SpotElevations", "OST_SpotCoordinates", "OST_SpotSlopes",
    "OST_AnnotationCrop", "OST_AnnotationCutlines", "OST_AnnotationObjects",
    "OST_ReferenceViewer", "OST_ReferenceViewerSymbol",
    "OST_GridHeads", "OST_LevelHeads", "OST_SectionHeads",
    "OST_ElevationMarks", "OST_CalloutHeads",
    "OST_CropBoundary", "OST_CropRegions",
    "OST_Annotation_SketchLines", "OST_Annotation_Lines",
    "OST_HiddenLines", "OST_DemolishedLines", "OST_OverheadLines",
    "OST_Lines", "OST_Curves", "OST_CurveGroups",
    "OST_Constraints", "OST_WeakDims",
    # Logical MEP systems (not the modeled elements)
    "OST_PipingSystem", "OST_DuctSystem", "OST_ConduitSystem",
}


def _bic_key(bic):
    """Stable BuiltInCategory name for denylist checks (IronPython-safe)."""
    if bic is None:
        return ""
    try:
        key = bic.ToString()
        if key:
            return key
    except Exception:
        pass
    return str(bic)


def _category_type_is_model(cat):
    """True only for CategoryType.Model. Uses ToString first (IronPython enums)."""
    try:
        ct = cat.CategoryType
    except Exception:
        return False
    try:
        name = ct.ToString()
        if name:
            return name == "Model"
    except Exception:
        pass
    try:
        return ct == CategoryType.Model
    except Exception:
        return False


def _is_clashable_category(cat):
    """Keep model categories that can have 3D clash geometry.

    Do not require HasMaterialQuantities: MEP curves (Ducts, Pipes, Conduits,
    Cable Trays) often report False even though they have clash solids.
    """
    if not cat:
        return False
    if not _category_type_is_model(cat):
        return False

    try:
        bic = cat.BuiltInCategory
        if bic is None:
            return False
        if _bic_key(bic) in _NON_GEOMETRIC_CATEGORIES:
            return False
    except Exception:
        return False

    return True


def _get_category_bic(cat):
    """Return the BuiltInCategory value for a Category object, or None."""
    try:
        bic = cat.BuiltInCategory
        if bic is None:
            return None
        try:
            if bic == BuiltInCategory.INVALID:
                return None
        except Exception:
            pass
        return bic
    except Exception:
        return None


_DISCOVER_CACHE = {}


def _document_cache_key(document):
    """Stable cache key — id(document) changes when GetLinkDocument() is recalled."""
    try:
        path = document.PathName or ""
    except Exception:
        path = ""
    try:
        title = document.Title or ""
    except Exception:
        title = ""
    try:
        linked = bool(document.IsLinked)
    except Exception:
        linked = False
    return (path, title, linked)


def _discover_clashable_categories(document):
    """Return sorted list of (name, BuiltInCategory) for model categories with instances.

    One pass over instances (not one collector per category). Cached per Document
    for the rest of this command run.
    """
    cache_key = _document_cache_key(document)
    cached = _DISCOVER_CACHE.get(cache_key)
    if cached is not None:
        return list(cached)

    by_name = {}
    try:
        for el in FilteredElementCollector(document).WhereElementIsNotElementType():
            try:
                cat = el.Category
            except Exception:
                continue
            if not cat:
                continue
            try:
                name = cat.Name
            except Exception:
                continue
            if not name or name in by_name:
                continue
            if not _is_clashable_category(cat):
                continue
            bic = _get_category_bic(cat)
            if bic is None:
                continue
            by_name[name] = bic
    except Exception as ex:
        logger.warning("Error discovering categories: {}".format(ex))

    results = sorted(by_name.items(), key=lambda pair: pair[0].lower())
    _DISCOVER_CACHE[cache_key] = results
    return list(results)


# =============================================================================
# Link discovery
# =============================================================================

def _discover_links(document):
    """Return sorted list of (display_name, link_instance_id) for loaded RevitLinkInstances."""
    results = []
    try:
        for inst in FilteredElementCollector(document).OfClass(RevitLinkInstance).ToElements():
            try:
                link_type_id = inst.GetTypeId()
                link_type = document.GetElement(link_type_id)
                if link_type is None:
                    continue
                if not RevitLinkType.IsLoaded(document, link_type_id):
                    continue
                name = inst.Name or "Unnamed Link"
                results.append((name, inst.Id))
            except Exception as ex:
                logger.debug("Skipping link instance: {}".format(ex))
    except Exception as ex:
        logger.warning("Error discovering links: {}".format(ex))

    results.sort(key=lambda pair: pair[0].lower())
    return results


def _get_link_doc(document, link_instance_id):
    """Return the linked Document for a RevitLinkInstance ElementId, or None."""
    try:
        inst = document.GetElement(link_instance_id)
        if inst is None:
            return None
        return inst.GetLinkDocument()
    except Exception as ex:
        logger.debug("Could not get link doc: {}".format(ex))
        return None


# =============================================================================
# Geometry helpers
# =============================================================================

def _outline_from_bbox(bbox, tol=1e-3):
    """Build an Outline from a BoundingBoxXYZ, with a small tolerance."""
    if not bbox:
        return None
    try:
        mn = XYZ(bbox.Min.X - tol, bbox.Min.Y - tol, bbox.Min.Z - tol)
        mx = XYZ(bbox.Max.X + tol, bbox.Max.Y + tol, bbox.Max.Z + tol)
        return Outline(mn, mx)
    except Exception:
        return None


def _union_bbox(elements, padding=0.0):
    """Return BoundingBoxXYZ covering all elements, expanded uniformly by padding (ft)."""
    mn_x = mn_y = mn_z = None
    mx_x = mx_y = mx_z = None
    for el in elements:
        try:
            eb = el.get_BoundingBox(None)
        except Exception:
            eb = None
        if not eb:
            continue
        if mn_x is None:
            mn_x, mn_y, mn_z = eb.Min.X, eb.Min.Y, eb.Min.Z
            mx_x, mx_y, mx_z = eb.Max.X, eb.Max.Y, eb.Max.Z
        else:
            if eb.Min.X < mn_x: mn_x = eb.Min.X
            if eb.Min.Y < mn_y: mn_y = eb.Min.Y
            if eb.Min.Z < mn_z: mn_z = eb.Min.Z
            if eb.Max.X > mx_x: mx_x = eb.Max.X
            if eb.Max.Y > mx_y: mx_y = eb.Max.Y
            if eb.Max.Z > mx_z: mx_z = eb.Max.Z

    if mn_x is None:
        return None

    bb = BoundingBoxXYZ()
    bb.Min = XYZ(mn_x - padding, mn_y - padding, mn_z - padding)
    bb.Max = XYZ(mx_x + padding, mx_y + padding, mx_z + padding)
    return bb


def _union_bbox_mixed(host_elements, link_elements_with_transform, padding=0.0):
    """BoundingBoxXYZ covering host elements + link elements (bbox transformed to host space)."""
    mn_x = mn_y = mn_z = None
    mx_x = mx_y = mx_z = None

    points = []
    for el in host_elements:
        try:
            eb = el.get_BoundingBox(None)
        except Exception:
            eb = None
        if not eb:
            continue
        points.append(eb.Min)
        points.append(eb.Max)

    for el, transform in link_elements_with_transform:
        try:
            eb = el.get_BoundingBox(None)
        except Exception:
            eb = None
        if not eb:
            continue
        for corner in _bbox_corners(eb):
            points.append(transform.OfPoint(corner))

    for pt in points:
        if mn_x is None:
            mn_x, mn_y, mn_z = pt.X, pt.Y, pt.Z
            mx_x, mx_y, mx_z = pt.X, pt.Y, pt.Z
        else:
            if pt.X < mn_x: mn_x = pt.X
            if pt.Y < mn_y: mn_y = pt.Y
            if pt.Z < mn_z: mn_z = pt.Z
            if pt.X > mx_x: mx_x = pt.X
            if pt.Y > mx_y: mx_y = pt.Y
            if pt.Z > mx_z: mx_z = pt.Z

    if mn_x is None:
        return None

    bb = BoundingBoxXYZ()
    bb.Min = XYZ(mn_x - padding, mn_y - padding, mn_z - padding)
    bb.Max = XYZ(mx_x + padding, mx_y + padding, mx_z + padding)
    return bb


def _clash_pair_bbox_center(el_a, el_b, link_transform=None, b_is_link=False):
    """Fallback: union bbox midpoint when solid clash point is unavailable."""
    host_elements = [el_a]
    link_pairs = []
    if b_is_link and link_transform is not None:
        link_pairs = [(el_b, link_transform)]
    else:
        host_elements.append(el_b)
    if link_pairs:
        bbox = _union_bbox_mixed(host_elements, link_pairs, padding=0.0)
    else:
        bbox = _union_bbox(host_elements, padding=0.0)
    if bbox is None:
        return None
    return XYZ(
        (bbox.Min.X + bbox.Max.X) / 2.0,
        (bbox.Min.Y + bbox.Max.Y) / 2.0,
        (bbox.Min.Z + bbox.Max.Z) / 2.0,
    )


def _solid_reference_point(solid):
    """Centroid of a solid, or bbox midpoint as fallback."""
    try:
        pt = solid.ComputeCentroid()
        if pt is not None:
            return pt
    except Exception:
        pass
    try:
        bb = solid.GetBoundingBox()
        if bb is not None:
            return XYZ(
                (bb.Min.X + bb.Max.X) / 2.0,
                (bb.Min.Y + bb.Max.Y) / 2.0,
                (bb.Min.Z + bb.Max.Z) / 2.0,
            )
    except Exception:
        pass
    return None


def _intersection_point_from_solids(solids_a, solids_b, tol=1e-9):
    """Clash point: centroid of the largest solid-solid intersection volume."""
    best_pt = None
    best_vol = tol
    for sa in solids_a:
        if sa is None:
            continue
        for sb in solids_b:
            if sb is None:
                continue
            try:
                inter = BooleanOperationsUtils.ExecuteBooleanOperation(
                    sa, sb, BooleanOperationsType.Intersect)
                if inter is None:
                    continue
                vol = inter.Volume
                if vol <= tol:
                    continue
                if vol > best_vol:
                    pt = _solid_reference_point(inter)
                    if pt is not None:
                        best_pt = pt
                        best_vol = vol
            except Exception:
                pass
    return best_pt


def _clash_pair_center(el_a, el_b, link_transform=None, b_is_link=False):
    """Model-space point at the solid-solid clash (largest intersection volume)."""
    solids_a = list(_iter_element_solids(el_a))
    if not solids_a:
        return _clash_pair_bbox_center(
            el_a, el_b, link_transform=link_transform, b_is_link=b_is_link)

    solids_b = list(_iter_element_solids(el_b))
    if b_is_link and link_transform is not None:
        solids_b = _transform_solids(solids_b, link_transform)

    if solids_b:
        pt = _intersection_point_from_solids(solids_a, solids_b)
        if pt is not None:
            return pt

    return _clash_pair_bbox_center(
        el_a, el_b, link_transform=link_transform, b_is_link=b_is_link)


def _bbox_corners(bbox):
    """Return all 8 corners of a BoundingBoxXYZ as XYZ list."""
    mn = bbox.Min
    mx = bbox.Max
    return [
        XYZ(mn.X, mn.Y, mn.Z),
        XYZ(mx.X, mn.Y, mn.Z),
        XYZ(mn.X, mx.Y, mn.Z),
        XYZ(mx.X, mx.Y, mn.Z),
        XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mx.Y, mx.Z),
    ]


def _transform_bbox_to_host(bbox, transform):
    """Transform all 8 corners and return a new axis-aligned BoundingBoxXYZ in host space."""
    corners = _bbox_corners(bbox)
    host_pts = [transform.OfPoint(c) for c in corners]
    mn_x = min(p.X for p in host_pts)
    mn_y = min(p.Y for p in host_pts)
    mn_z = min(p.Z for p in host_pts)
    mx_x = max(p.X for p in host_pts)
    mx_y = max(p.Y for p in host_pts)
    mx_z = max(p.Z for p in host_pts)
    bb = BoundingBoxXYZ()
    bb.Min = XYZ(mn_x, mn_y, mn_z)
    bb.Max = XYZ(mx_x, mx_y, mx_z)
    return bb


def _iter_element_solids(element):
    """Yield Solid objects with Volume > 1e-9 from an element's geometry.

    Recurses through GeometryInstance.GetInstanceGeometry(). Uses options that
    skip references and non-visible geometry for speed.
    """
    if element is None:
        return
    opts = Options()
    opts.ComputeReferences = False
    opts.IncludeNonVisibleObjects = False
    opts.DetailLevel = ViewDetailLevel.Fine
    try:
        geom = element.get_Geometry(opts)
    except Exception:
        return
    if geom is None:
        return
    stack = list(geom)
    while stack:
        obj = stack.pop()
        if obj is None:
            continue
        if isinstance(obj, Solid):
            try:
                if obj.Volume > 1e-9:
                    yield obj
            except Exception:
                pass
        elif isinstance(obj, GeometryInstance):
            try:
                inst_geom = obj.GetInstanceGeometry()
                if inst_geom:
                    stack.extend(inst_geom)
            except Exception:
                pass


def _transform_solids(solids, transform):
    """Transform solids to a new coordinate space using SolidUtils.CreateTransformed.

    Returns list of transformed solids; drops any that fail.
    """
    result = []
    for s in solids:
        try:
            if s is None:
                continue
            ts = SolidUtils.CreateTransformed(s, transform)
            if ts is not None:
                result.append(ts)
        except Exception:
            pass
    return result


def _solids_intersect_info(solids_a, solids_b, tol=1e-9):
    """Return (intersects, clash_point) for solid pairs.

    Single boolean pass: clash_point is centroid of largest intersection volume.
    """
    best_pt = None
    best_vol = tol
    found = False
    for sa in solids_a:
        if sa is None:
            continue
        for sb in solids_b:
            if sb is None:
                continue
            try:
                inter = BooleanOperationsUtils.ExecuteBooleanOperation(
                    sa, sb, BooleanOperationsType.Intersect)
                if inter is None:
                    continue
                vol = inter.Volume
                if vol <= tol:
                    continue
                found = True
                if vol > best_vol:
                    pt = _solid_reference_point(inter)
                    if pt is not None:
                        best_pt = pt
                        best_vol = vol
            except Exception:
                pass
    return found, best_pt


def _solids_intersect(solids_a, solids_b, tol=1e-9):
    """Return True if any solid from A intersects any solid from B."""
    found, _ = _solids_intersect_info(solids_a, solids_b, tol=tol)
    return found


def _element_level_name(document, element):
    """Best-effort level name for an element; 'No Level' if unknown."""
    try:
        lvl_id = element.LevelId
        if lvl_id and lvl_id != ElementId.InvalidElementId:
            lvl = document.GetElement(lvl_id)
            if lvl and lvl.Name:
                return lvl.Name
    except Exception:
        pass

    try:
        host = getattr(element, 'Host', None)
        if host is not None:
            host_lvl_id = host.LevelId
            if host_lvl_id and host_lvl_id != ElementId.InvalidElementId:
                host_lvl = document.GetElement(host_lvl_id)
                if host_lvl and host_lvl.Name:
                    return host_lvl.Name
    except Exception:
        pass

    try:
        lvl_param = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
        if lvl_param and lvl_param.HasValue:
            lvl_id = lvl_param.AsElementId()
            if lvl_id and lvl_id != ElementId.InvalidElementId:
                lvl = document.GetElement(lvl_id)
                if lvl and lvl.Name:
                    return lvl.Name
    except Exception:
        pass

    return "No Level"


# =============================================================================
# View/sheet name sanitation
# =============================================================================

# Revit forbids these characters in view / sheet names.
_REVIT_NAME_FORBIDDEN = u'\\:{}[]<>;?|*`~'


def _sanitize_revit_name(name):
    """Strip characters that Revit rejects in view/sheet names."""
    if name is None:
        return u""
    try:
        s = unicode(name)
    except Exception:
        s = str(name)
    out_chars = []
    for ch in s:
        if ch in _REVIT_NAME_FORBIDDEN:
            out_chars.append(u" ")
        else:
            out_chars.append(ch)
    cleaned = u"".join(out_chars)
    while u"  " in cleaned:
        cleaned = cleaned.replace(u"  ", u" ")
    return cleaned.strip()


def _host_category_id(document, bic):
    """Return the host-document Category.Id for a BuiltInCategory, or None."""
    try:
        cat = Category.GetCategory(document, bic)
        if cat is not None:
            return cat.Id
    except Exception:
        pass
    return None


def _hide_non_target_model_categories(view, keep_cat_id_values):
    """In *view*, hide every overridable Model category except those in keep set.

    *keep_cat_id_values* is an iterable of category-id integer values.
    Host view category visibility also drives linked-model visibility when the
    link uses the "By host view" display mode (the Revit default), so hiding
    non-target categories here scopes both host AND link content to the two
    clashing categories.

    OST_RvtLinks is always exempt — hiding it would suppress all linked models.
    """
    keep_set = set()
    for v in keep_cat_id_values:
        if v is None:
            continue
        try:
            keep_set.add(int(v))
        except Exception:
            continue
    document = view.Document
    # Always exempt OST_RvtLinks regardless of caller's keep set
    try:
        rvtlinks_cat = Category.GetCategory(document, BuiltInCategory.OST_RvtLinks)
        if rvtlinks_cat is not None:
            keep_set.add(int(get_element_id_value(rvtlinks_cat.Id)))
    except Exception:
        pass
    hidden = 0
    skipped = 0
    try:
        categories = document.Settings.Categories
    except Exception:
        return (0, 0)
    for cat in categories:
        try:
            if cat.CategoryType != CategoryType.Model:
                continue
        except Exception:
            continue
        try:
            cid = cat.Id
        except Exception:
            continue
        try:
            cid_val = int(get_element_id_value(cid))
        except Exception:
            try:
                cid_val = int(cid.IntegerValue)
            except Exception:
                continue
        if cid_val in keep_set:
            try:
                if view.CanCategoryBeHidden(cid):
                    view.SetCategoryHidden(cid, False)
            except Exception:
                pass
            continue
        try:
            if not view.CanCategoryBeHidden(cid):
                skipped += 1
                continue
        except Exception:
            pass
        try:
            view.SetCategoryHidden(cid, True)
            hidden += 1
        except Exception:
            skipped += 1
            continue
    return (hidden, skipped)


def _ensure_link_instance_visible(
        view, link_instance_id, document,
        _List=List,
        _ElementId=ElementId,
        _Category=Category,
        _BuiltInCategory=BuiltInCategory,
        _WorksetVisibility=WorksetVisibility,
        _FilteredWorksetCollector=FilteredWorksetCollector,
        _WorksetKind=WorksetKind):
    """Make the target Revit link actually show in *view*.

    HideElements (first pass unhide and second-pass PostCommand) only affects
    currently visible elements. Isolate-to-permanent, a hidden RVT Links
    category, or a hidden host workset will make the link invisible so the
    hide command no-ops.

    Must run inside an open transaction. Idempotent.
    Default-arg captures keep types alive after pyRevit disposes the script
    module (Idling second pass).
    """
    try:
        if view.AreModelCategoriesHidden:
            view.AreModelCategoriesHidden = False
    except Exception:
        pass
    try:
        rvtlinks_cat = _Category.GetCategory(document, _BuiltInCategory.OST_RvtLinks)
        if rvtlinks_cat is not None:
            cid = rvtlinks_cat.Id
            if view.CanCategoryBeHidden(cid):
                view.SetCategoryHidden(cid, False)
    except Exception:
        pass

    li = None
    try:
        li = document.GetElement(link_instance_id)
    except Exception:
        li = None
    if li is None:
        return

    # Host workset of the link instance. UnhideElements cannot show a
    # workset-hidden link; HideElements then has nothing to hide from.
    if getattr(document, "IsWorkshared", False):
        try:
            ws_id = li.WorksetId
            if ws_id is not None:
                try:
                    view.SetWorksetVisibility(ws_id, _WorksetVisibility.Visible)
                except Exception:
                    # View template may lock workset VG. Detach and retry —
                    # clash views already own their VG (isolate, categories).
                    try:
                        view.ViewTemplateId = _ElementId.InvalidElementId
                        view.SetWorksetVisibility(ws_id, _WorksetVisibility.Visible)
                    except Exception:
                        pass
        except Exception:
            pass

    try:
        if li.IsHidden(view):
            unhide = _List[_ElementId]()
            unhide.Add(link_instance_id)
            view.UnhideElements(unhide)
    except Exception:
        pass

    # Nested worksets inside the linked model (VG > Revit Links > Worksets).
    # Only via RevitLinkGraphicsSettings when the API exposes it — never
    # View.SetWorksetVisibility with linked WorksetIds (ids collide with host).
    try:
        settings = view.GetLinkOverrides(link_instance_id)
    except Exception:
        settings = None
    setter = getattr(settings, "SetWorksetVisibility", None) if settings else None
    if setter is not None:
        try:
            link_doc = li.GetLinkDocument()
            if link_doc is not None and getattr(link_doc, "IsWorkshared", False):
                vis = _WorksetVisibility.Visible
                changed = False
                for ws in _FilteredWorksetCollector(link_doc).OfKind(_WorksetKind.UserWorkset):
                    try:
                        setter(ws.Id, vis)
                        changed = True
                    except Exception:
                        continue
                if changed:
                    view.SetLinkOverrides(link_instance_id, settings)
        except Exception:
            pass

    for bic_name in ("OST_VolumeOfInterest", "OST_ScopeBoxes"):
        try:
            bic = getattr(_BuiltInCategory, bic_name, None)
            if bic is None:
                continue
            cat = _Category.GetCategory(document, bic)
            if cat is None:
                continue
            if view.CanCategoryBeHidden(cat.Id):
                view.SetCategoryHidden(cat.Id, True)
        except Exception:
            pass


# =============================================================================
# Refinement pipeline (per-element complement hiding via Idling + PostCommand)
# =============================================================================

RefinementJob = namedtuple("RefinementJob", [
    "view_id",           # ElementId of the created View3D (in host doc)
    "link_instance_id",  # ElementId of RevitLinkInstance (in host doc)
    "a_cat_id",          # ElementId of the A (host) category — hide ALL in link
    "b_cat_id",          # ElementId of the B (link) category (first / primary)
    "b_clash_link_eids", # list[int] — linked-doc ElementId integers that ARE clash B
    "bbox_min",          # XYZ — host-space section-box min
    "bbox_max",          # XYZ — host-space section-box max
    "b_cat_ids",         # list[ElementId] — all B (link) categories in this view
])

# Stored on sys so it survives script scope cleanup between pushbutton runs.
# pyRevit disposes the script module after execution, which would GC module-level
# lists; sys persists for the AppDomain (Revit session) lifetime.
_SYS_KEY = '_pyBS_clash_refinement_drivers'
_SYS_GLOBALS_KEY = '_pyBS_clash_g'
_SYS_TOOL_KEY = '_pyBS_clash_tool_window'

# Last-run inputs so the summary Refresh button can rebuild views after Close().
ClashRunSpec = namedtuple('ClashRunSpec', [
    'categories',
    'scale',
    'view_type_id',
    'name_prefix',
    'sheet_prefix',
    'sheet_name',
    'iteration',
    'group_by_level',
    'link_instance_id',
    'link_categories',
    'combine_link_views',
])


def _snapshot_clash_globals(_sys=sys, _gkey=_SYS_GLOBALS_KEY):
    """Keep a copy of this script's globals; pyRevit may clear the module dict."""
    setattr(_sys, _gkey, dict(globals()))
    setattr(_sys, '_pyBS_clash_restore', _restore_clash_globals)


def _restore_clash_globals(_sys=sys, _key=_SYS_GLOBALS_KEY):
    """Put snapped names back if pyRevit disposed this pushbutton module."""
    snap = getattr(_sys, _key, None)
    if not snap:
        return
    g = globals()
    if g.get('doc') is None or g.get('ClashViewsWindow') is None:
        g.update(snap)


def _pick_non_clash_view(document, doomed_values):
    """Any non-template view not in the set about to be deleted."""
    try:
        views = FilteredElementCollector(document).OfClass(View).ToElements()
    except Exception:
        return None
    fallback = None
    for v in views:
        try:
            if getattr(v, 'IsTemplate', False):
                continue
            vid = int(get_element_id_value(v.Id))
            if vid in doomed_values:
                continue
            if isinstance(v, ViewSheet):
                if fallback is None:
                    fallback = v
                continue
            return v
        except Exception:
            continue
    return fallback


class LinkVisibilityRefinementDriver(object):
    """Idling-event state machine that hides unwanted elements from the link
    in each created clash view, one view per idle tick.

    Two kinds of elements are hidden per view (when A and B differ):
      1. All link-A-category elements (e.g. Walls in the link when B is Floors).
      2. Non-clash link-B-category elements in the section-box area.

    When A and B share a category (e.g. Walls vs Walls), step 1 is skipped —
    otherwise every link wall is hidden, including clash B walls.

    Workaround for View.HideElements rejecting linked element ids (forum-validated):
      build Reference.CreateLinkReference, set as uidoc.Selection, post HideElements.
    """

    MAX_COMPLEMENT_SIZE = 500

    def __init__(self, uiapp, jobs, document, return_to_sheet_id=None,
                 summary_data=None, show_summary_callback=None, progress_close_callback=None,
                 progress_update_callback=None):
        self._uiapp = uiapp
        self._jobs = list(jobs)
        self._idx = 0
        self._doc = document
        self._handler = None
        self._return_to_sheet_id = return_to_sheet_id
        self._summary_data = summary_data
        self._show_summary_callback = show_summary_callback
        self._progress_close_callback = progress_close_callback
        self._progress_update_callback = progress_update_callback
        # Capture module-level names needed by _process at init-time.
        # pyRevit disposes the script module after the pushbutton returns.
        self._ElementId = ElementId
        self._FilteredElementCollector = FilteredElementCollector
        self._BoundingBoxIntersectsFilter = BoundingBoxIntersectsFilter
        self._Outline = Outline
        self._XYZ = XYZ
        self._Reference = Reference
        self._List = List
        self._get_element_id_value = get_element_id_value
        self._RevitCommandId = RevitCommandId
        self._PostableCommand = PostableCommand
        self._Transaction = Transaction
        self._ensure_link_instance_visible = _ensure_link_instance_visible
        self._logger = logger
        self._sys_key = _SYS_KEY

    def _view_ids_equal(self, id_a, id_b):
        """Value compare ElementIds (IronPython `!=` is unreliable on 64-bit ids)."""
        if id_a is None or id_b is None:
            return False
        try:
            return int(self._get_element_id_value(id_a)) == int(
                self._get_element_id_value(id_b))
        except Exception:
            return id_a == id_b

    def start(self):
        if not self._jobs:
            return
        self._handler = self._on_idling
        self._uiapp.Idling += self._handler
        import sys
        if not hasattr(sys, self._sys_key):
            setattr(sys, self._sys_key, [])
        getattr(sys, self._sys_key).append(self)
        # Kick the first view change before pyRevit disposes the script module.
        try:
            uidoc = self._uiapp.ActiveUIDocument
            view = self._doc.GetElement(self._jobs[0].view_id)
            if uidoc is not None and view is not None:
                uidoc.RequestViewChange(view)
        except Exception:
            pass

    def stop(self):
        if self._handler is not None:
            try:
                self._uiapp.Idling -= self._handler
            except Exception:
                pass
            self._handler = None
        import sys
        try:
            getattr(sys, self._sys_key).remove(self)
        except (AttributeError, ValueError):
            pass

    def _on_idling(self, sender, e):
        try:
            self._on_idling_body(sender, e)
        except Exception as ex:
            self._logger.debug("Link refinement idle failed: {}".format(ex))

    def _on_idling_body(self, sender, e):
        if self._idx >= len(self._jobs):
            if self._progress_close_callback is not None:
                try:
                    self._progress_close_callback()
                except Exception as ex:
                    self._logger.debug("Failed to close progress window: {}".format(ex))
            if self._return_to_sheet_id is not None:
                try:
                    uidoc = self._uiapp.ActiveUIDocument
                    if uidoc is not None:
                        sheet = self._doc.GetElement(self._return_to_sheet_id)
                        if sheet is not None:
                            uidoc.RequestViewChange(sheet)
                            if self._show_summary_callback is not None and self._summary_data is not None:
                                try:
                                    self._show_summary_callback(self._summary_data)
                                except Exception as ex:
                                    self._logger.debug("Failed to show summary dialog: {}".format(ex))
                except Exception:
                    pass
            self.stop()
            return
        job = self._jobs[self._idx]
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None:
            self.stop()
            return
        current = uidoc.ActiveView
        if current is None or not self._view_ids_equal(current.Id, job.view_id):
            view = self._doc.GetElement(job.view_id)
            if view is None:
                self._idx += 1
                return
            try:
                uidoc.RequestViewChange(view)
            except Exception:
                self._idx += 1
            try:
                e.SetRaiseWithoutDelay()
            except Exception:
                pass
            return
        self._idx += 1
        if self._progress_update_callback is not None:
            try:
                self._progress_update_callback(self._idx)
            except Exception:
                pass
        try:
            self._process(job, uidoc)
        except Exception as ex:
            self._logger.debug("Link refinement failed: {}".format(ex))

    def _process(self, job, uidoc):
        view = self._doc.GetElement(job.view_id)
        if view is None:
            return
        link_instance = self._doc.GetElement(job.link_instance_id)
        if link_instance is None:
            return
        linked_doc = link_instance.GetLinkDocument()
        if linked_doc is None:
            return
        # HideElements only affects currently visible elements. Isolate and
        # workset VG can leave the link off; show it before posting hide.
        t_vis = self._Transaction(self._doc, "Show clash link")
        try:
            t_vis.Start()
            self._ensure_link_instance_visible(
                view, job.link_instance_id, self._doc)
            t_vis.Commit()
        except Exception as ex:
            try:
                t_vis.RollBack()
            except Exception:
                pass
            self._logger.debug("Could not show clash link before hide: {}".format(ex))
        finally:
            try:
                t_vis.Dispose()
            except Exception:
                pass
        refs = []
        hide_link_eids = []
        _EId = self._ElementId
        _FEC = self._FilteredElementCollector
        _BBF = self._BoundingBoxIntersectsFilter
        _Ol = self._Outline
        _XYZ = self._XYZ
        _Ref = self._Reference
        _giv = self._get_element_id_value
        clash_set = set(job.b_clash_link_eids)
        b_cat_ids = list(job.b_cat_ids or [])
        if not b_cat_ids and job.b_cat_id is not None:
            b_cat_ids = [job.b_cat_id]
        same_cat = False
        if job.a_cat_id is not None:
            for cid in b_cat_ids:
                try:
                    if int(_giv(job.a_cat_id)) == int(_giv(cid)):
                        same_cat = True
                        break
                except Exception:
                    if str(job.a_cat_id) == str(cid):
                        same_cat = True
                        break
        if job.a_cat_id is not None and not same_cat:
            a_cat_fresh = _EId(_giv(job.a_cat_id))
            try:
                a_eids = list(_FEC(linked_doc)
                              .OfCategoryId(a_cat_fresh)
                              .WhereElementIsNotElementType()
                              .ToElementIds())
                for eid in a_eids:
                    el = linked_doc.GetElement(eid)
                    if el is None:
                        continue
                    try:
                        refs.append(_Ref(el).CreateLinkReference(link_instance))
                        hide_link_eids.append(_giv(eid))
                    except Exception:
                        continue
            except Exception:
                pass
        if b_cat_ids:
            transform = None
            link_min = None
            link_max = None
            try:
                transform = link_instance.GetTotalTransform()
                inv = transform.Inverse
                p1 = inv.OfPoint(job.bbox_min)
                p2 = inv.OfPoint(job.bbox_max)
                link_min = _XYZ(min(p1.X, p2.X), min(p1.Y, p2.Y), min(p1.Z, p2.Z))
                link_max = _XYZ(max(p1.X, p2.X), max(p1.Y, p2.Y), max(p1.Z, p2.Z))
            except Exception:
                pass
            for b_cat_id in b_cat_ids:
                all_b_eids = []
                b_cat_fresh = _EId(_giv(b_cat_id))
                try:
                    if link_min is not None and link_max is not None:
                        all_b_eids = list(_FEC(linked_doc)
                                          .OfCategoryId(b_cat_fresh)
                                          .WhereElementIsNotElementType()
                                          .WherePasses(_BBF(_Ol(link_min, link_max)))
                                          .ToElementIds())
                    else:
                        raise Exception("no bbox")
                except Exception:
                    try:
                        all_b_eids = list(_FEC(linked_doc)
                                          .OfCategoryId(b_cat_fresh)
                                          .WhereElementIsNotElementType()
                                          .ToElementIds())
                    except Exception:
                        all_b_eids = []
                for eid in all_b_eids:
                    if _giv(eid) in clash_set:
                        continue
                    el = linked_doc.GetElement(eid)
                    if el is None:
                        continue
                    try:
                        refs.append(_Ref(el).CreateLinkReference(link_instance))
                        hide_link_eids.append(_giv(eid))
                    except Exception:
                        continue
        clash_wrongly_hidden = sorted(clash_set.intersection(set(hide_link_eids)))
        if clash_wrongly_hidden:
            self._logger.warning(
                "Link refinement: {} clash B link elements marked for hide in '{}' "
                "(same_category={})".format(
                    len(clash_wrongly_hidden), view.Name, same_cat))
        outcome = 'ok'
        if not refs:
            outcome = 'no_refs'
        elif len(refs) > self.MAX_COMPLEMENT_SIZE:
            outcome = 'too_many_refs'
            self._logger.warning(
                "Link refinement: too many refs ({}) for view '{}', skipping.".format(
                    len(refs), view.Name))
        set_refs_ok = False
        if outcome == 'ok':
            _L = self._List
            try:
                uidoc.Selection.SetReferences(_L[_Ref](refs))
                set_refs_ok = True
            except Exception as ex:
                outcome = 'setrefs_fail'
                self._logger.warning("Link refinement: SetReferences failed: {}".format(ex))
            if set_refs_ok:
                try:
                    cmd_id = self._RevitCommandId.LookupPostableCommandId(
                        self._PostableCommand.HideElements)
                    if self._uiapp.CanPostCommand(cmd_id):
                        self._uiapp.PostCommand(cmd_id)
                    else:
                        outcome = 'cannot_post'
                except Exception as ex:
                    outcome = 'postcmd_fail'
                    self._logger.warning("Link refinement: PostCommand failed: {}".format(ex))


# =============================================================================
# Color overrides
# =============================================================================

CLASH_COLOR_A = RevitColor(230, 50, 50)
CLASH_COLOR_B = RevitColor(40, 170, 70)

# Clash views use fixed red (side A / host) and green (side B / link) for all pairs.
# Distinctive per-category colors are not enabled yet. To add them later, map each
# selected category to palette slots (e.g. warm pole for host, cool pole for link via
# bic + side key) and pass per-pair overrides into _apply_clash_view_colors — see
# revit.revit_utils.generate_color_range and ColorElements.generate_color_range.


def _darken_color(color, factor=0.7):
    """Return a darkened RevitColor by multiplying RGB values by factor (0.0-1.0).

    Default factor 0.7 makes lines ~30% darker than fill.
    """
    try:
        r = int(color.Red * factor)
        g = int(color.Green * factor)
        b = int(color.Blue * factor)
        # Clamp to valid range
        r = max(0, min(255, r))
        g = max(0, min(255, g))
        b = max(0, min(255, b))
        return RevitColor(r, g, b)
    except Exception:
        return color


def _solid_fill_pattern_id(document):
    """Return the ElementId of any solid fill pattern, or InvalidElementId."""
    try:
        for pat in FilteredElementCollector(document).OfClass(FillPatternElement):
            try:
                if pat.GetFillPattern().IsSolidFill:
                    return pat.Id
            except Exception:
                continue
    except Exception:
        pass
    return ElementId.InvalidElementId


def _build_clash_override(color, solid_fill_id):
    """Build an OverrideGraphicSettings that paints the element solid with color.

    Lines (projection/cut) use a darkened shade (~30% darker) for visual distinction.
    Surfaces and cut patterns use the full-brightness color.
    """
    ogs = OverrideGraphicSettings()
    line_color = _darken_color(color, 0.7)
    try:
        ogs.SetProjectionLineColor(line_color)
        ogs.SetCutLineColor(line_color)
    except Exception:
        pass
    try:
        ogs.SetSurfaceForegroundPatternColor(color)
        ogs.SetCutForegroundPatternColor(color)
        if solid_fill_id is not None and solid_fill_id != ElementId.InvalidElementId:
            ogs.SetSurfaceForegroundPatternId(solid_fill_id)
            ogs.SetCutForegroundPatternId(solid_fill_id)
        try:
            ogs.SetSurfaceForegroundPatternVisible(True)
            ogs.SetCutForegroundPatternVisible(True)
        except Exception:
            pass
    except Exception:
        pass
    return ogs


def _apply_category_color_override(view, cat_id, ogs):
    """Apply OverrideGraphicSettings to a category via SetCategoryOverrides."""
    if cat_id is None:
        return False, "cat_id is None"
    try:
        if not view.IsCategoryOverridable(cat_id):
            return False, "category not overridable"
    except Exception as ex:
        return False, "IsCategoryOverridable: {}".format(ex)
    try:
        view.SetCategoryOverrides(cat_id, ogs)
        return True, None
    except Exception as ex:
        return False, str(ex)


def _same_category_ids(cat_a_id, cat_b_id):
    """True when both sides resolve to the same host Category.Id."""
    if cat_a_id is None or cat_b_id is None:
        return False
    try:
        return int(get_element_id_value(cat_a_id)) == int(get_element_id_value(cat_b_id))
    except Exception:
        return str(cat_a_id) == str(cat_b_id)


def _normalize_cat_ids(cat_b_id, cat_b_ids=None):
    """Unique Category.Id list from optional list plus a single fallback id."""
    out = []
    seen = set()
    seq = list(cat_b_ids or [])
    if cat_b_id is not None:
        seq.append(cat_b_id)
    for cid in seq:
        if cid is None:
            continue
        try:
            v = int(get_element_id_value(cid))
        except Exception:
            continue
        if v in seen:
            continue
        seen.add(v)
        out.append(cid)
    return out


def _same_category_any(cat_a_id, cat_b_ids):
    """True when cat_a matches any id in cat_b_ids."""
    for cid in cat_b_ids or []:
        if _same_category_ids(cat_a_id, cid):
            return True
    return False


def _append_unique_cat_id(id_list, cat_id):
    """Append cat_id to id_list if its integer value is not already present."""
    if cat_id is None:
        return
    try:
        v = int(get_element_id_value(cat_id))
    except Exception:
        return
    for existing in id_list:
        try:
            if int(get_element_id_value(existing)) == v:
                return
        except Exception:
            continue
    id_list.append(cat_id)


def _apply_element_override_list(view, element_ids, ogs):
    """Apply overrides to host-document ElementIds. Returns (success_count, fail_count)."""
    success = 0
    failures = 0
    for eid in element_ids or []:
        try:
            view.SetElementOverrides(eid, ogs)
            success += 1
        except Exception:
            failures += 1
    return success, failures


def _apply_link_category_color_override(view, link_instance_id, cat_id, ogs):
    """Color linked-model category via host view SetCategoryOverrides.

    When link object styles follow the host view (ByHostView, Revit default),
    category overrides on the host view apply to matching linked categories.
    SetLinkOverrides with ObjectStyles=Custom is not supported by the API.
    """
    if link_instance_id is None or cat_id is None:
        return False, "missing link_instance_id or cat_id"

    ok, err = _apply_category_color_override(view, cat_id, ogs)
    if ok:
        return True, None

    try:
        view.SetElementOverrides(link_instance_id, ogs)
        return True, None
    except Exception as ex:
        return False, err or str(ex)


def _apply_clash_view_colors(view, ovr_a, ovr_b, is_link_mode=False,
                             link_instance_id=None, cat_a_id=None, cat_b_id=None,
                             a_eids=None, b_eids=None, cat_b_ids=None):
    """Apply clash colors via category overrides (or per-element fallback)."""
    a_eids = a_eids or []
    b_eids = b_eids or []
    b_cat_list = _normalize_cat_ids(cat_b_id, cat_b_ids)
    if cat_b_id is None and b_cat_list:
        cat_b_id = b_cat_list[0]
    same_cat = _same_category_any(cat_a_id, b_cat_list)
    results = {
        "use_category_link_overrides": USE_CATEGORY_LINK_OVERRIDES,
        "is_link_mode": is_link_mode,
        "same_category": same_cat,
        "a": None,
        "b": None,
        "fallbacks": [],
    }

    if USE_CATEGORY_LINK_OVERRIDES and same_cat:
        if is_link_mode and link_instance_id is not None:
            # API cannot SetElementOverrides on individual link elements.
            # Category green on shared cat + per-element red on host clash A walls.
            ok_b = True
            err_b = None
            for cid in b_cat_list or [cat_b_id]:
                ok_i, err_i = _apply_category_color_override(view, cid, ovr_b)
                if not ok_i:
                    ok_b = False
                    err_b = err_i
            count_a, fail_a = _apply_element_override_list(view, a_eids, ovr_a)
            results["a"] = {
                "method": "SetElementOverrides(same_cat_host_a)",
                "count": count_a,
                "failures": fail_a,
            }
            results["b"] = {
                "method": "SetCategoryOverrides(same_cat_link_b)",
                "ok": ok_b,
                "error": err_b,
            }
            results["fallbacks"].append("same_category_link_hybrid")
        else:
            count_a, fail_a = _apply_element_override_list(view, a_eids, ovr_a)
            count_b, fail_b = _apply_element_override_list(view, b_eids, ovr_b)
            results["a"] = {
                "method": "SetElementOverrides(same_cat_a)",
                "count": count_a,
                "failures": fail_a,
            }
            results["b"] = {
                "method": "SetElementOverrides(same_cat_b)",
                "count": count_b,
                "failures": fail_b,
            }
            results["fallbacks"].append("same_category_host_per_element")
    elif USE_CATEGORY_LINK_OVERRIDES and cat_a_id is not None:
        ok, err = _apply_category_color_override(view, cat_a_id, ovr_a)
        results["a"] = {"method": "SetCategoryOverrides", "ok": ok, "error": err}
    else:
        count, failures = _apply_element_override_list(view, a_eids, ovr_a)
        results["a"] = {
            "method": "SetElementOverrides",
            "count": count,
            "failures": failures,
        }
        results["fallbacks"].append("host_a per-element")

    if not (USE_CATEGORY_LINK_OVERRIDES and same_cat):
        if is_link_mode and link_instance_id is not None:
            if USE_CATEGORY_LINK_OVERRIDES and b_cat_list:
                ok, err = True, None
                for cid in b_cat_list:
                    ok_i, err_i = _apply_link_category_color_override(
                        view, link_instance_id, cid, ovr_b)
                    if not ok_i:
                        ok, err = ok_i, err_i
                results["b"] = {
                    "method": "SetCategoryOverrides(link_b)",
                    "ok": ok,
                    "error": err,
                }
            else:
                try:
                    view.SetElementOverrides(link_instance_id, ovr_b)
                    results["b"] = {"method": "SetElementOverrides(link)", "ok": True}
                except Exception as ex:
                    results["b"] = {
                        "method": "SetElementOverrides(link)",
                        "ok": False,
                        "error": str(ex),
                    }
        elif USE_CATEGORY_LINK_OVERRIDES and b_cat_list:
            ok, err = True, None
            for cid in b_cat_list:
                ok_i, err_i = _apply_category_color_override(view, cid, ovr_b)
                if not ok_i:
                    ok, err = ok_i, err_i
            results["b"] = {"method": "SetCategoryOverrides", "ok": ok, "error": err}
        else:
            count, failures = _apply_element_override_list(view, b_eids, ovr_b)
            results["b"] = {
                "method": "SetElementOverrides",
                "count": count,
                "failures": failures,
            }
            results["fallbacks"].append("host_b per-element")

    if DEBUG_MODE:
        logger.debug("Clash view colors: {}".format(results))
    return results


# =============================================================================
# Clash detection – host only
# =============================================================================

def _clash_pairs(document, bic_a, bic_b):
    """Detect clashing element pairs between two BuiltInCategory values in host doc.

    BoundingBoxIntersectsFilter pre-filter + ElementIntersectsElementFilter confirm.
    Returns list of (element_a, element_b).
    """
    if bic_a == bic_b:
        return []

    try:
        elems_a = list(FilteredElementCollector(document)
                       .OfCategory(bic_a)
                       .WhereElementIsNotElementType()
                       .ToElements())
    except Exception as ex:
        logger.debug("Could not collect category {}: {}".format(bic_a, ex))
        return []

    try:
        b_ids_all = (FilteredElementCollector(document)
                     .OfCategory(bic_b)
                     .WhereElementIsNotElementType()
                     .ToElementIds())
    except Exception as ex:
        logger.debug("Could not collect category {}: {}".format(bic_b, ex))
        return []

    b_ids_list = List[ElementId]()
    for eid in b_ids_all:
        b_ids_list.Add(eid)

    if b_ids_list.Count == 0 or not elems_a:
        return []

    pairs = []
    seen = set()

    for a in elems_a:
        try:
            a_bbox = a.get_BoundingBox(None)
        except Exception:
            a_bbox = None
        if not a_bbox:
            continue

        outline = _outline_from_bbox(a_bbox)
        if outline is None:
            continue

        try:
            bbox_filter = BoundingBoxIntersectsFilter(outline)
            candidates = list(FilteredElementCollector(document, b_ids_list)
                              .WherePasses(bbox_filter)
                              .ToElements())
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("bbox prefilter failed for {}: {}".format(a.Id, ex))
            continue
        if not candidates:
            continue

        try:
            geom_filter = ElementIntersectsElementFilter(a)
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("geom filter could not be built for {}: {}".format(a.Id, ex))
            continue

        for b in candidates:
            if b.Id == a.Id:
                continue
            try:
                if not geom_filter.PassesFilter(document, b.Id):
                    continue
            except Exception as ex:
                if DEBUG_MODE:
                    logger.debug("geom filter failed on {}/{}: {}".format(a.Id, b.Id, ex))
                continue
            key = (min(get_element_id_value(a.Id), get_element_id_value(b.Id)),
                   max(get_element_id_value(a.Id), get_element_id_value(b.Id)))
            if key in seen:
                continue
            seen.add(key)
            center = _clash_pair_center(a, b)
            pairs.append((a, b, center))

    return pairs


# =============================================================================
# Clash detection – host vs linked model (bbox prefilter + solid intersection)
# =============================================================================

def _clash_pairs_with_link(document, bic_host, link_instance, link_doc, bic_link):
    """Detect clashing element pairs between host category and linked-model category.

    Uses bounding-box prefilter to find candidates, then real solid-solid intersection
    in host coordinate space (link solids transformed via SolidUtils.CreateTransformed).

    Returns list of (host_element, link_element) tuples.
    """
    try:
        transform = link_instance.GetTotalTransform()
    except Exception as ex:
        logger.warning("Could not get link transform: {}".format(ex))
        return []

    try:
        host_elems = list(FilteredElementCollector(document)
                          .OfCategory(bic_host)
                          .WhereElementIsNotElementType()
                          .ToElements())
    except Exception as ex:
        logger.debug("Could not collect host category {}: {}".format(bic_host, ex))
        return []

    try:
        link_elems = list(FilteredElementCollector(link_doc)
                          .OfCategory(bic_link)
                          .WhereElementIsNotElementType()
                          .ToElements())
    except Exception as ex:
        logger.debug("Could not collect link category {}: {}".format(bic_link, ex))
        return []

    if not host_elems or not link_elems:
        return []

    # Pre-build host element id list for BoundingBoxIntersectsFilter
    host_ids = List[ElementId]()
    for el in host_elems:
        host_ids.Add(el.Id)

    pairs = []
    seen = set()

    for link_el in link_elems:
        try:
            link_bbox = link_el.get_BoundingBox(None)
        except Exception:
            link_bbox = None
        if not link_bbox:
            continue

        # Transform link bbox to host coordinate space (all 8 corners)
        try:
            host_space_bbox = _transform_bbox_to_host(link_bbox, transform)
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("Transform failed for link elem {}: {}".format(link_el.Id, ex))
            continue

        outline = _outline_from_bbox(host_space_bbox)
        if outline is None:
            continue

        try:
            bbox_filter = BoundingBoxIntersectsFilter(outline)
            candidates = list(FilteredElementCollector(document, host_ids)
                              .WherePasses(bbox_filter)
                              .ToElements())
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("bbox filter failed for link elem {}: {}".format(link_el.Id, ex))
            continue
        if not candidates:
            continue

        # Cache link solids once per link element and transform to host space
        link_solids = list(_iter_element_solids(link_el))
        transformed_link_solids = _transform_solids(link_solids, transform)
        has_link_solids = len(transformed_link_solids) > 0

        for host_el in candidates:
            link_id_val = get_element_id_value(link_el.Id)
            host_id_val = get_element_id_value(host_el.Id)
            key = (host_id_val, link_id_val)
            if key in seen:
                continue

            clash_pt = None
            # If we have transformed link solids, do real solid-solid intersection
            if has_link_solids:
                host_solids = list(_iter_element_solids(host_el))
                if host_solids:
                    intersects, clash_pt = _solids_intersect_info(
                        host_solids, transformed_link_solids)
                    if not intersects:
                        continue
                else:
                    # Host element has no extractable solids (lines, etc.)
                    # Fall back to bbox match so we don't silently drop clashes
                    pass
            # If no link solids, fall back to bbox match (non-solid link elements)

            seen.add(key)
            pairs.append((host_el, link_el, clash_pt))

    return pairs


def _merge_grouped_link_views(grouped, link_name):
    """Collapse host×link-category groups into one group per host category."""
    buckets = OrderedDict()
    cat_names_by_host = OrderedDict()
    for pair_key, by_level in grouped.items():
        host_name = None
        for _level_name, ids_map in by_level.items():
            host_name = ids_map.get("host_name") or host_name
            if host_name:
                break
        if not host_name:
            try:
                host_name = pair_key.split(" vs ")[0]
            except Exception:
                host_name = pair_key
        if host_name not in buckets:
            buckets[host_name] = OrderedDict()
            cat_names_by_host[host_name] = []
        for level_name, ids_map in by_level.items():
            lname = ids_map.get("link_cat_name")
            if lname and lname not in cat_names_by_host[host_name]:
                cat_names_by_host[host_name].append(lname)
            if level_name not in buckets[host_name]:
                buckets[host_name][level_name] = {
                    "a_ids": set(),
                    "b_ids": set(),
                    "link_mode": True,
                    "cat_a_id": ids_map.get("cat_a_id"),
                    "cat_b_id": ids_map.get("cat_b_id"),
                    "cat_b_ids": [],
                    "pair_points": [],
                    "host_name": host_name,
                }
            dest = buckets[host_name][level_name]
            dest["a_ids"].update(ids_map.get("a_ids") or [])
            dest["b_ids"].update(ids_map.get("b_ids") or [])
            dest["pair_points"].extend(list(ids_map.get("pair_points") or []))
            src_b = ids_map.get("cat_b_ids") or []
            if not src_b and ids_map.get("cat_b_id") is not None:
                src_b = [ids_map.get("cat_b_id")]
            for cid in src_b:
                _append_unique_cat_id(dest["cat_b_ids"], cid)
            if dest.get("cat_b_id") is None:
                dest["cat_b_id"] = ids_map.get("cat_b_id")

    merged = OrderedDict()
    for host_name, by_level in buckets.items():
        names = cat_names_by_host.get(host_name) or []
        if len(names) <= 2:
            b_label = " + ".join(names) if names else "combined"
        else:
            b_label = "{} categories".format(len(names))
        pair_key = "{} vs {} [Link: {}]".format(host_name, b_label, link_name)
        merged[pair_key] = by_level
    return merged


# =============================================================================
# Defaults
# =============================================================================

DEFAULT_VIEW_SCALE = 100
DEFAULT_CROP_OFFSET_MM = 500.0
BASE_SHEET_NAME = "Category Clash Views"

_INVALID_ITERATION_CHARS = '\\/{}[]:|<>*?"'


def _sanitize_iteration_token(text):
    """Strip and replace characters unsafe for Revit sheet numbers / view names."""
    if not text:
        return ""
    s = (text or "").strip()
    if not s:
        return ""
    parts = []
    for c in s:
        if c in _INVALID_ITERATION_CHARS or ord(c) < 32:
            parts.append("-")
        else:
            parts.append(c)
    return "".join(parts)


def _view_family_type_display_name(vft):
    if not vft:
        return ""
    try:
        name_param = vft.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if name_param and name_param.HasValue:
            return name_param.AsString() or ""
    except Exception:
        pass
    try:
        return vft.Name or ""
    except Exception:
        return ""


def _default_three_d_view_type_id(document):
    """Return ElementId of first ViewFamilyType with ViewFamily.ThreeDimensional, sorted by name."""
    vfts = []
    for vft in FilteredElementCollector(document).OfClass(ViewFamilyType).ToElements():
        try:
            if vft.ViewFamily == ViewFamily.ThreeDimensional:
                vfts.append(vft)
        except Exception:
            continue
    if not vfts:
        return None
    vfts.sort(key=lambda v: _view_family_type_display_name(v).lower())
    return vfts[0].Id


# =============================================================================
# Data classes
# =============================================================================

class CategoryFilterItem(forms.Reactive):
    """One category row in the category checklist."""

    def __init__(self, category_name, bic, is_selected=False):
        super(CategoryFilterItem, self).__init__()
        self._category_name = category_name
        self._bic = bic
        self._is_selected = is_selected

    @property
    def CategoryName(self):
        return self._category_name

    @property
    def Bic(self):
        return self._bic

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        self.OnPropertyChanged("IsSelected")


class LinkItem(forms.Reactive):
    """One link entry in the link selector ComboBox."""

    def __init__(self, display_name, link_instance_id):
        super(LinkItem, self).__init__()
        self._display_name = display_name
        self._link_instance_id = link_instance_id

    @property
    def DisplayName(self):
        return self._display_name

    @property
    def LinkInstanceId(self):
        return self._link_instance_id

    def ToString(self):
        """WPF ComboBox selection box / text search use .NET ToString when templates omit."""
        return self._display_name or ""


# =============================================================================
# Main window
# =============================================================================

class ClashViewsWindow(forms.WPFWindow):
    """WPF window for creating category clash views on a sheet."""

    GRID_COLUMNS = 20
    GROUP_GAP_ROWS = 1
    VIEWPORT_SPACING = 0.03

    def __init__(self, host_categories=None):
        logger.debug("Initializing Clash Views window")

        xaml_file = op.join(pushbutton_dir, "ClashViewsWindow.xaml")
        forms.WPFWindow.__init__(self, xaml_file)

        try:
            from styles import load_styles_to_window
            load_styles_to_window(self)
        except Exception as ex:
            logger.debug("Could not load styles: {}".format(ex))

        # Host category collections
        self._category_items = ObservableCollection[CategoryFilterItem]()
        self._category_display = ObservableCollection[CategoryFilterItem]()

        # Link item collection
        self._link_items = ObservableCollection[LinkItem]()

        # Link category collections
        self._link_category_items = ObservableCollection[CategoryFilterItem]()
        self._link_category_display = ObservableCollection[CategoryFilterItem]()

        self.clash_views = {}
        self.created_sheets = []
        self._refinement_driver = None
        self._run_spec = None
        self._marker_points_by_view = {}

        self._populate_category_filters(host_categories)
        self._sync_category_list_display()
        self.categoryListBox.ItemsSource = self._category_display
        self._update_category_selection_count()

        # Link selector
        self.linkSelectorComboBox.ItemsSource = self._link_items
        self.linkCategoryListBox.ItemsSource = self._link_category_display

    # -------------------- Host category UI setup --------------------------

    def _populate_category_filters(self, host_categories=None):
        self._category_items.Clear()
        pairs = host_categories if host_categories is not None else _discover_clashable_categories(doc)
        for name, bic in pairs:
            self._category_items.Add(CategoryFilterItem(name, bic, False))

    def _sync_category_list_display(self):
        self._category_display.Clear()
        search = ""
        if hasattr(self, "categorySearchTextBox") and self.categorySearchTextBox:
            search = (self.categorySearchTextBox.Text or "").strip().lower()
        for item in self._category_items:
            if not search or search in item.CategoryName.lower():
                self._category_display.Add(item)

    def _update_category_selection_count(self):
        n_sel = sum(1 for i in self._category_items if i.IsSelected)
        n_tot = self._category_items.Count
        n_pairs = n_sel * (n_sel - 1) // 2 if n_sel > 1 else 0
        if hasattr(self, "categorySelectionCountText") and self.categorySelectionCountText:
            self.categorySelectionCountText.Text = (
                "{} of {} categories selected ({} pairs)".format(n_sel, n_tot, n_pairs))

    def _get_selected_categories(self):
        """Return list of (name, bic) for checked host categories."""
        out = []
        for item in self._category_items:
            if item.IsSelected:
                out.append((item.CategoryName, item.Bic))
        return out

    # -------------------- Link UI setup -----------------------------------

    def _populate_links(self):
        self._link_items.Clear()
        for name, inst_id in _discover_links(doc):
            self._link_items.Add(LinkItem(name, inst_id))

    def _populate_link_categories(self, link_doc):
        self._link_category_items.Clear()
        self._link_category_display.Clear()
        if link_doc is None:
            self._update_link_category_count()
            return
        for name, bic in _discover_clashable_categories(link_doc):
            self._link_category_items.Add(CategoryFilterItem(name, bic, False))
        self._sync_link_category_list_display()
        self._update_link_category_count()

    def _sync_link_category_list_display(self):
        self._link_category_display.Clear()
        search = ""
        if hasattr(self, "linkCategorySearchTextBox") and self.linkCategorySearchTextBox:
            search = (self.linkCategorySearchTextBox.Text or "").strip().lower()
        for item in self._link_category_items:
            if not search or search in item.CategoryName.lower():
                self._link_category_display.Add(item)

    def _update_link_category_count(self):
        if not (hasattr(self, "linkCategorySelectionCountText") and self.linkCategorySelectionCountText):
            return
        n_sel = sum(1 for i in self._link_category_items if i.IsSelected)
        n_tot = self._link_category_items.Count
        if n_tot == 0:
            self.linkCategorySelectionCountText.Text = "Select a link model above"
        else:
            self.linkCategorySelectionCountText.Text = (
                "{} of {} link categories selected".format(n_sel, n_tot))

    def _get_selected_link_categories(self):
        """Return list of (name, bic) for checked link categories."""
        out = []
        for item in self._link_category_items:
            if item.IsSelected:
                out.append((item.CategoryName, item.Bic))
        return out

    def _get_selected_link(self):
        """Return (link_instance, link_doc) or (None, None)."""
        selected = self.linkSelectorComboBox.SelectedItem
        if selected is None:
            return None, None
        inst_id = selected.LinkInstanceId
        inst = doc.GetElement(inst_id)
        if inst is None:
            return None, None
        link_doc = inst.GetLinkDocument()
        return inst, link_doc

    # -------------------- Host category event handlers -------------------

    def categorySearchTextBox_TextChanged(self, sender, args):
        self._sync_category_list_display()

    def categorySelectAllButton_Click(self, sender, args):
        for item in self._category_items:
            item.IsSelected = True
        self._update_category_selection_count()

    def categoryDeselectAllButton_Click(self, sender, args):
        for item in self._category_items:
            item.IsSelected = False
        self._update_category_selection_count()

    def categoryFilterCheckBox_Changed(self, sender, args):
        self._update_category_selection_count()

    # -------------------- Link event handlers ----------------------------

    def againstLinkToggle_Changed(self, sender, args):
        is_checked = bool(self.againstLinkToggle.IsChecked)
        if is_checked and self._link_items.Count == 0:
            self._populate_links()
            if self._link_items.Count == 0:
                forms.alert("No loaded Revit links found in the project.", title="No Links")
        if not is_checked:
            # Reset link categories on uncheck
            self._link_category_items.Clear()
            self._link_category_display.Clear()
            self._update_link_category_count()

    def linkSelectorComboBox_SelectionChanged(self, sender, args):
        inst, link_doc = self._get_selected_link()
        self._populate_link_categories(link_doc)

    def linkCategorySearchTextBox_TextChanged(self, sender, args):
        self._sync_link_category_list_display()

    def linkSelectAllButton_Click(self, sender, args):
        for item in self._link_category_items:
            item.IsSelected = True
        self._update_link_category_count()

    def linkDeselectAllButton_Click(self, sender, args):
        for item in self._link_category_items:
            item.IsSelected = False
        self._update_link_category_count()

    def linkCategoryFilterCheckBox_Changed(self, sender, args):
        self._update_link_category_count()

    # -------------------- Shared handlers --------------------------------

    def cancelButton_Click(self, sender, args):
        self.Close()

    def createButton_Click(self, sender, args):
        selected = self._get_selected_categories()
        against_link = bool(self.againstLinkToggle.IsChecked)

        if against_link:
            if len(selected) < 1:
                forms.alert("Pick at least one host category.", title="Warning")
                return
            link_inst, link_doc = self._get_selected_link()
            if link_inst is None or link_doc is None:
                forms.alert("Select a loaded link model.", title="Warning")
                return
            link_cats = self._get_selected_link_categories()
            if len(link_cats) < 1:
                forms.alert("Pick at least one link category.", title="Warning")
                return
        else:
            if len(selected) < 2:
                forms.alert("Pick at least two categories to clash.", title="Warning")
                return
            link_inst = None
            link_doc = None
            link_cats = []

        view_type_id = _default_three_d_view_type_id(doc)
        if view_type_id is None:
            forms.alert("No 3D view family type found in the project.", title="Error")
            return

        scale = DEFAULT_VIEW_SCALE
        name_prefix = self.namePrefixTextBox.Text or "Clash - "
        sheet_prefix = self.sheetPrefixTextBox.Text or "CV-"
        iteration = _sanitize_iteration_token(
            self.iterationTextBox.Text if hasattr(self, "iterationTextBox") and self.iterationTextBox else "")
        if iteration:
            sheet_name = "{} ({})".format(BASE_SHEET_NAME, iteration)
        else:
            sheet_name = BASE_SHEET_NAME
        group_by_level = bool(self.groupByLevelCheckBox.IsChecked)
        combine_link_views = False
        if against_link and hasattr(self, "combineLinkViewsToggle") and self.combineLinkViewsToggle:
            combine_link_views = bool(self.combineLinkViewsToggle.IsChecked)

        self.Close()

        self._run_spec = ClashRunSpec(
            categories=list(selected),
            scale=scale,
            view_type_id=view_type_id,
            name_prefix=name_prefix,
            sheet_prefix=sheet_prefix,
            sheet_name=sheet_name,
            iteration=iteration,
            group_by_level=group_by_level,
            link_instance_id=link_inst.Id if link_inst is not None else None,
            link_categories=list(link_cats),
            combine_link_views=combine_link_views,
        )
        _snapshot_clash_globals()
        setattr(sys, _SYS_TOOL_KEY, self)

        try:
            if against_link:
                n_pairs = len(selected) * len(link_cats)
                title = "Detecting clashes ({} host cats x {} link cats = {} pairs)".format(
                    len(selected), len(link_cats), n_pairs)
            else:
                n_pairs = len(selected) * (len(selected) - 1) // 2
                title = "Detecting clashes ({} categories, {} pairs)".format(
                    len(selected), n_pairs)

            with forms.ProgressBar(title=title) as pb:
                self._create_clash_views(
                    selected, scale, view_type_id,
                    name_prefix, sheet_prefix, sheet_name, iteration,
                    group_by_level, pb,
                    link_instance=link_inst,
                    link_doc=link_doc,
                    link_categories=link_cats,
                    combine_link_views=combine_link_views,
                )

            # Refinement driver returns to the sheet when it finishes. Switching
            # now would fight RequestViewChange on the clash 3D views.
            if self.created_sheets and not getattr(self, '_refinement_driver', None):
                uidoc.ActiveView = self.created_sheets[0]

        except Exception as ex:
            logger.error("Error creating clash views: {}".format(ex))
            import traceback
            logger.debug(traceback.format_exc())
            forms.alert("Error creating clash views: {}".format(ex), title="Error")

    # -------------------- Orchestration ----------------------------------

    def _collect_clash_output_ids(self):
        """ElementIds of 3D views and sheets from the last successful run."""
        doomed = []
        doomed_values = set()
        for view in (self.clash_views or {}).values():
            try:
                if view is None:
                    continue
                doomed.append(view.Id)
                doomed_values.add(int(get_element_id_value(view.Id)))
            except Exception:
                continue
        for sheet in (self.created_sheets or []):
            try:
                if sheet is None:
                    continue
                doomed.append(sheet.Id)
                doomed_values.add(int(get_element_id_value(sheet.Id)))
            except Exception:
                continue
        return doomed, doomed_values

    def _leave_clash_output_views(self):
        """Move the active view off output we are about to delete."""
        doomed, doomed_values = self._collect_clash_output_ids()
        if not doomed:
            return
        try:
            active = uidoc.ActiveView
            if active is None:
                return
            if int(get_element_id_value(active.Id)) not in doomed_values:
                return
            fallback = _pick_non_clash_view(doc, doomed_values)
            if fallback is not None:
                uidoc.ActiveView = fallback
        except Exception as ex:
            logger.debug("Could not leave clash view before delete: {}".format(ex))

    def _delete_clash_output_ids(self, doomed):
        """Delete previous clash views/sheet. Must run inside an open transaction."""
        if not doomed:
            return
        id_list = List[ElementId]()
        for eid in doomed:
            id_list.Add(eid)
        try:
            doc.Delete(id_list)
        except Exception as ex:
            logger.warning("Bulk delete of clash views failed: {}".format(ex))
            for eid in doomed:
                try:
                    doc.Delete(eid)
                except Exception:
                    pass

    def refresh_clash_views(self, summary, _restore=_restore_clash_globals):
        """Re-run the last spec: detect, replace views/sheet, update summary."""
        _restore()
        spec = getattr(self, '_run_spec', None)
        if spec is None:
            forms.alert("No previous clash run to refresh.", title="Refresh")
            return 'missing_spec'

        driver = getattr(self, '_refinement_driver', None)
        if driver is not None:
            try:
                driver.stop()
            except Exception:
                pass
            self._refinement_driver = None

        link_inst = None
        link_doc = None
        if spec.link_instance_id is not None:
            link_inst = doc.GetElement(spec.link_instance_id)
            if link_inst is None:
                forms.alert(
                    "Linked model from this run is gone. Refresh cancelled.",
                    title="Refresh")
                return 'missing_link'
            try:
                link_doc = link_inst.GetLinkDocument()
            except Exception:
                link_doc = None
            if link_doc is None:
                forms.alert(
                    "Linked model is unloaded. Load it, then Refresh.",
                    title="Refresh")
                return 'unloaded_link'

        view_type_id = spec.view_type_id
        try:
            if view_type_id is None or doc.GetElement(view_type_id) is None:
                view_type_id = _default_three_d_view_type_id(doc)
        except Exception:
            view_type_id = _default_three_d_view_type_id(doc)
        if view_type_id is None:
            forms.alert("No 3D view family type found in the project.", title="Error")
            return 'no_view_type'

        categories = list(spec.categories)
        link_cats = list(spec.link_categories or [])
        against_link = spec.link_instance_id is not None
        if against_link:
            n_pairs = len(categories) * len(link_cats)
            title = "Refreshing clashes ({} host cats x {} link cats = {} pairs)".format(
                len(categories), len(link_cats), n_pairs)
        else:
            n_pairs = len(categories) * (len(categories) - 1) // 2
            title = "Refreshing clashes ({} categories, {} pairs)".format(
                len(categories), n_pairs)

        with forms.ProgressBar(title=title) as pb:
            status = self._create_clash_views(
                categories, spec.scale, view_type_id,
                spec.name_prefix, spec.sheet_prefix, spec.sheet_name,
                spec.iteration, spec.group_by_level, pb,
                link_instance=link_inst,
                link_doc=link_doc,
                link_categories=link_cats,
                combine_link_views=spec.combine_link_views,
                replace_existing=True,
                existing_summary=summary,
            )
        return status or 'ok'

    def _create_clash_views(self, categories, scale, view_type_id,
                            name_prefix, sheet_prefix, sheet_name, iteration,
                            group_by_level, progress_bar=None,
                            link_instance=None, link_doc=None, link_categories=None,
                            combine_link_views=False,
                            replace_existing=False, existing_summary=None):
        """Run clash detection for all category pairs, create views, then place on a sheet.

        When link_instance is provided, pairs are host-cat x link-cat cross-product
        and clash detection uses bbox-only intersection in host coordinate space.
        combine_link_views merges those pairs into one view per host category.
        replace_existing deletes this run's previous sheet/views inside the same
        transaction group (rolled back if create fails).
        """
        link_mode = link_instance is not None and link_doc is not None

        try:
            crop_offset = UnitUtils.ConvertToInternalUnits(
                DEFAULT_CROP_OFFSET_MM, DB.UnitTypeId.Millimeters)
        except Exception:
            crop_offset = DEFAULT_CROP_OFFSET_MM / 304.8

        if link_mode:
            # Cross-product: every host cat vs every link cat
            link_name = link_instance.Name or "Link"
            pair_list = [
                ((hname, hbic), (lname, lbic))
                for (hname, hbic) in categories
                for (lname, lbic) in (link_categories or [])
            ]
        else:
            pair_list = list(combinations(categories, 2))

        grouped = OrderedDict()
        total_clashes = 0
        if not replace_existing:
            self._marker_points_by_view = {}

        link_transform = None
        if link_mode:
            try:
                link_transform = link_instance.GetTotalTransform()
            except Exception:
                link_transform = None

        for idx, ((name_a, bic_a), (name_b, bic_b)) in enumerate(pair_list):
            if progress_bar:
                try:
                    progress_bar.title = "Clash {}/{}: {} vs {}".format(
                        idx + 1, len(pair_list), name_a, name_b)
                except Exception:
                    pass
                progress_bar.update_progress(idx, len(pair_list))

            try:
                if link_mode:
                    pairs = _clash_pairs_with_link(doc, bic_a, link_instance, link_doc, bic_b)
                else:
                    pairs = _clash_pairs(doc, bic_a, bic_b)
            except Exception as ex:
                logger.warning("Clash detection failed for {} vs {}: {}".format(
                    name_a, name_b, ex))
                continue

            if not pairs:
                continue

            if link_mode:
                pair_key = "{} vs {} [Link: {}]".format(name_a, name_b, link_name)
            else:
                pair_key = "{} vs {}".format(name_a, name_b)

            cat_a_id = _host_category_id(doc, bic_a)
            cat_b_id = _host_category_id(doc, bic_b)
            by_level = OrderedDict()
            for pair in pairs:
                a = pair[0]
                b = pair[1]
                cached_center = pair[2] if len(pair) > 2 else None
                if group_by_level:
                    level_key = _element_level_name(doc, a)
                else:
                    level_key = "All"
                if level_key not in by_level:
                    b_ids_list = []
                    if cat_b_id is not None:
                        b_ids_list = [cat_b_id]
                    by_level[level_key] = {
                        "a_ids": set(),
                        "b_ids": set(),
                        "link_mode": link_mode,
                        "cat_a_id": cat_a_id,
                        "cat_b_id": cat_b_id,
                        "cat_b_ids": b_ids_list,
                        "pair_points": [],
                        "host_name": name_a,
                        "link_cat_name": name_b,
                    }
                by_level[level_key]["a_ids"].add(get_element_id_value(a.Id))
                by_level[level_key]["b_ids"].add(get_element_id_value(b.Id))
                center = cached_center
                if center is None and not (link_mode and link_transform is None):
                    center = _clash_pair_center(
                        a, b, link_transform=link_transform, b_is_link=link_mode)
                if center is not None:
                    by_level[level_key]["pair_points"].append({
                        "x": center.X,
                        "y": center.Y,
                        "z": center.Z,
                        "a_id": get_element_id_value(a.Id),
                        "b_id": get_element_id_value(b.Id),
                    })

            grouped[pair_key] = by_level
            total_clashes += len(pairs)

        if link_mode and combine_link_views and grouped:
            grouped = _merge_grouped_link_views(grouped, link_name)

        if progress_bar:
            progress_bar.update_progress(len(pair_list), len(pair_list))

        if not grouped:
            forms.alert("No clashes detected between selected categories.",
                        title="No Clashes")
            return 'no_clashes'

        # Phase 1: compute bboxes and uniform size before the transaction
        natural_info = OrderedDict()
        max_dx = max_dy = max_dz = 0.0

        for pair_key, by_level in grouped.items():
            for level_name, ids_map in by_level.items():
                a_ids = list(ids_map["a_ids"])
                b_ids = list(ids_map["b_ids"])
                is_link = ids_map.get("link_mode", False)

                host_elements = []
                for v in a_ids:
                    el = doc.GetElement(make_element_id(v))
                    if el is not None:
                        host_elements.append(el)

                if is_link and link_transform is not None:
                    link_elements_with_transform = []
                    for v in b_ids:
                        el = link_doc.GetElement(make_element_id(v))
                        if el is not None:
                            link_elements_with_transform.append((el, link_transform))
                    bbox = _union_bbox_mixed(host_elements, link_elements_with_transform,
                                            padding=crop_offset)
                else:
                    all_elements = list(host_elements)
                    for v in b_ids:
                        el = doc.GetElement(make_element_id(v))
                        if el is not None:
                            all_elements.append(el)
                    bbox = _union_bbox(all_elements, padding=crop_offset)

                if bbox is None:
                    continue
                dx = bbox.Max.X - bbox.Min.X
                dy = bbox.Max.Y - bbox.Min.Y
                dz = bbox.Max.Z - bbox.Min.Z
                if dx > max_dx: max_dx = dx
                if dy > max_dy: max_dy = dy
                if dz > max_dz: max_dz = dz
                natural_info[(pair_key, level_name)] = {
                    "a_ids": a_ids,
                    "b_ids": b_ids,
                    "link_mode": is_link,
                    "cat_a_id": ids_map.get("cat_a_id"),
                    "cat_b_id": ids_map.get("cat_b_id"),
                    "cat_b_ids": list(ids_map.get("cat_b_ids") or (
                        [ids_map.get("cat_b_id")] if ids_map.get("cat_b_id") else [])),
                    "pair_points": list(ids_map.get("pair_points") or []),
                    "center": XYZ(
                        (bbox.Min.X + bbox.Max.X) / 2.0,
                        (bbox.Min.Y + bbox.Max.Y) / 2.0,
                        (bbox.Min.Z + bbox.Max.Z) / 2.0,
                    ),
                }

        if not natural_info:
            forms.alert("Clashes were detected but no usable bounding boxes.",
                        title="No Views")
            return 'no_views'

        if max_dx < 1.0: max_dx = 1.0
        if max_dy < 1.0: max_dy = 1.0
        if max_dz < 1.0: max_dz = 1.0

        self._uniform_dx = max_dx
        self._uniform_dy = max_dy
        self._uniform_dz = max_dz

        total_views_to_create = len(natural_info)
        logger.debug(
            "Found {} clashes across {} pairs; {} views to create; "
            "uniform box = ({:.2f} x {:.2f} x {:.2f}) ft".format(
                total_clashes, len(grouped), total_views_to_create,
                max_dx, max_dy, max_dz))

        solid_fill_id = _solid_fill_pattern_id(doc)
        ovr_a = _build_clash_override(CLASH_COLOR_A, solid_fill_id)
        ovr_b = _build_clash_override(CLASH_COLOR_B, solid_fill_id)

        refinement_queue = []

        old_view_map = dict(self.clash_views) if replace_existing else None
        old_sheets = list(self.created_sheets) if replace_existing else None
        old_markers = dict(getattr(self, '_marker_points_by_view', {}) or {}) if replace_existing else None
        doomed_ids = []
        if replace_existing:
            doomed_ids, _doomed_values = self._collect_clash_output_ids()
            self._leave_clash_output_views()

        tg = TransactionGroup(doc, "Create Category Clash Views")
        tg.Start()
        try:
            if replace_existing:
                with Transaction(doc, "Remove previous clash views") as t:
                    t.Start()
                    self._delete_clash_output_ids(doomed_ids)
                    t.Commit()
                self.clash_views = {}
                self.created_sheets = []
                self._marker_points_by_view = {}

            processed = 0
            total_views = len(natural_info)
            for (pair_key, level_name), info in natural_info.items():
                processed += 1
                if progress_bar:
                    try:
                        progress_bar.title = "Building view {}/{}".format(
                            processed, total_views)
                    except Exception:
                        pass
                    progress_bar.update_progress(processed, total_views)

                center = info["center"]
                uniform = BoundingBoxXYZ()
                uniform.Min = XYZ(
                    center.X - max_dx / 2.0,
                    center.Y - max_dy / 2.0,
                    center.Z - max_dz / 2.0,
                )
                uniform.Max = XYZ(
                    center.X + max_dx / 2.0,
                    center.Y + max_dy / 2.0,
                    center.Z + max_dz / 2.0,
                )

                shell = None
                vg = TransactionGroup(doc, "Clash View")
                vg.Start()
                try:
                    with Transaction(doc, "Create Clash 3D View") as t:
                        t.Start()
                        shell = self._create_clash_view_shell(
                            pair_key, level_name,
                            view_type_id, scale, name_prefix, iteration,
                            info["a_ids"], info["b_ids"],
                            uniform_bbox=uniform)
                        t.Commit()

                    if shell is None:
                        vg.RollBack()
                        continue

                    view = shell['view']
                    self.clash_views[(pair_key, level_name)] = view
                    pair_pts = info.get("pair_points") or []
                    if pair_pts:
                        try:
                            vid = get_element_id_value(view.Id)
                            self._marker_points_by_view[vid] = pair_pts
                        except Exception:
                            pass
                    shell['ovr_a'] = ovr_a
                    shell['ovr_b'] = ovr_b
                    shell['is_link_mode'] = info.get("link_mode", False)
                    shell['link_instance_id'] = (
                        link_instance.Id if link_instance is not None else None)
                    shell['cat_a_id'] = info.get("cat_a_id")
                    shell['cat_b_id'] = info.get("cat_b_id")
                    shell['cat_b_ids'] = list(info.get("cat_b_ids") or [])
                    shell['uniform_bbox'] = uniform
                    shell['info'] = info

                    with Transaction(doc, "Configure Clash 3D View") as t:
                        t.Start()
                        try:
                            self._configure_clash_view(shell)
                        except Exception as ex:
                            logger.warning(
                                "Configure clash view failed for {} / {}: {}".format(
                                    shell.get('pair_key'), shell.get('level_name'), ex))
                        t.Commit()

                    vg.Assimilate()
                finally:
                    vg.Dispose()

                if info.get("link_mode") and shell.get('link_instance_id') is not None:
                    refinement_queue.append(RefinementJob(
                        view_id=shell['view'].Id,
                        link_instance_id=shell['link_instance_id'],
                        a_cat_id=info.get("cat_a_id"),
                        b_cat_id=info.get("cat_b_id"),
                        b_clash_link_eids=list(info.get("b_ids") or []),
                        bbox_min=uniform.Min,
                        bbox_max=uniform.Max,
                        b_cat_ids=list(info.get("cat_b_ids") or []),
                    ))

            with Transaction(doc, "Create Sheet and Place Viewports") as t:
                t.Start()
                self._create_sheet_and_place_viewports(
                    grouped, sheet_prefix, iteration, sheet_name, scale)
                t.Commit()

            tg.Assimilate()
        except Exception:
            try:
                tg.RollBack()
            except Exception:
                pass
            if replace_existing:
                if old_view_map is not None:
                    self.clash_views = old_view_map
                if old_sheets is not None:
                    self.created_sheets = old_sheets
                if old_markers is not None:
                    self._marker_points_by_view = old_markers
            raise
        finally:
            tg.Dispose()

        # Show progress window immediately if link refinement is needed
        # This gives user feedback during the summary data building and before idling
        progress_window = None
        if refinement_queue:
            progress_window = RefinementProgressWindow(len(refinement_queue))
            progress_window.Show()

        # Build summary data for the results dialog
        view_names = []
        clash_pairs = []
        for pair_key, by_level in grouped.items():
            pair_count = 0
            for level_name, ids_map in by_level.items():
                if (pair_key, level_name) in self.clash_views:
                    view = self.clash_views[(pair_key, level_name)]
                    try:
                        view_names.append(view.Name)
                    except Exception:
                        view_names.append(pair_key + " - " + level_name)
                    pair_count += len(ids_map.get("a_ids", set()))
            clash_pairs.append((pair_key, pair_count))

        summary_data = {
            'sheet': self.created_sheets[0] if self.created_sheets else None,
            'view_names': view_names,
            'total_clashes': total_clashes,
            'clash_pairs': clash_pairs,
            'marker_data': {
                'available': _CLASH_MARKERS_AVAILABLE,
                'view_points': dict(self._marker_points_by_view),
            },
        }
        if register_marker_session and self._marker_points_by_view:
            register_marker_session(
                doc, self._marker_points_by_view, toggle_active=True)

        # Callback to show or refresh the summary dialog (after refinement)
        def _show_summary(data, _Summary=ClashViewsSummaryWindow, _logger=logger,
                          _existing=existing_summary, _tool=self):
            try:
                import sys as _sys
                if _existing is not None:
                    _existing.apply_run_results(data)
                    return
                summary_window = _Summary(
                    data['sheet'], data['view_names'], data['total_clashes'],
                    data['clash_pairs'], data.get('marker_data'),
                    tool_window=_tool)
                _sys._pyBS_clash_summary_window = summary_window
                summary_window.Show()
            except Exception as ex:
                _logger.debug("Failed to show summary window: {}".format(ex))

        # Start the per-element link refinement pipeline after the transaction
        # group has been committed (elements exist in the model)
        if refinement_queue:
            # Get the first created sheet to return to after refinement
            return_sheet_id = None
            if self.created_sheets:
                return_sheet_id = self.created_sheets[0].Id

            # Callback to update progress from driver
            def _update_progress(current):
                try:
                    if progress_window:
                        progress_window.update_progress(current)
                except Exception:
                    pass

            # Callback to close progress when done
            def _close_progress():
                try:
                    if progress_window:
                        progress_window.Close()
                except Exception:
                    pass

            driver = LinkVisibilityRefinementDriver(
                __revit__, refinement_queue, doc,
                return_to_sheet_id=return_sheet_id,
                summary_data=summary_data,
                show_summary_callback=_show_summary,
                progress_close_callback=_close_progress,
                progress_update_callback=_update_progress)
            self._refinement_driver = driver
            driver.start()
            return 'pending_refine'
        else:
            # No refinement needed - close progress if shown and show summary
            if progress_window:
                try:
                    progress_window.Close()
                except Exception:
                    pass
            if self.created_sheets:
                _show_summary(summary_data)
            return 'ok'

    # -------------------- View creation ----------------------------------

    def _create_clash_view_shell(self, pair_key, level_name,
                                 view_type_id, scale, name_prefix, iteration,
                                 a_id_values, b_id_values, uniform_bbox=None):
        """Create 3D view with section box applied before first regen."""
        has_a = False
        for v in a_id_values or []:
            if doc.GetElement(make_element_id(v)) is not None:
                has_a = True
                break
        if not has_a:
            return None

        view = View3D.CreateIsometric(doc, view_type_id)
        if view is None:
            return None

        try:
            view.Scale = scale
        except Exception:
            pass

        section_box_set = False
        if uniform_bbox is not None:
            try:
                view.SetSectionBox(uniform_bbox)
                try:
                    view.IsSectionBoxActive = True
                except Exception:
                    pass
                section_box_set = True
            except Exception as ex:
                logger.debug("SetSectionBox failed for {} / {}: {}".format(
                    pair_key, level_name, ex))

        try:
            view.DetailLevel = ViewDetailLevel.Fine
        except Exception as ex:
            logger.debug("Set DetailLevel Fine failed for {} / {}: {}".format(
                pair_key, level_name, ex))

        if iteration:
            raw_name = "{0}[{1}] {2} - {3}".format(
                name_prefix, iteration, pair_key, level_name)
        else:
            raw_name = "{0}{1} - {2}".format(name_prefix, pair_key, level_name)
        base_name = _sanitize_revit_name(raw_name)
        try:
            view.Name = base_name
        except Exception:
            for i in range(1, 100):
                try:
                    view.Name = u"{} ({})".format(base_name, i)
                    break
                except Exception:
                    continue

        return {
            'view': view,
            'pair_key': pair_key,
            'level_name': level_name,
            'a_id_values': list(a_id_values or []),
            'b_id_values': list(b_id_values or []),
            'section_box_set': section_box_set,
        }

    def _configure_clash_view(self, shell):
        """Apply isolate, category scope, and colors after view exists.

        Deferred from view creation so Revit finishes link GRep for the new
        view before isolate/overrides run (avoids regen conflicts).
        """
        view = shell['view']
        pair_key = shell['pair_key']
        level_name = shell['level_name']
        a_id_values = shell['a_id_values']
        b_id_values = shell['b_id_values']
        ovr_a = shell['ovr_a']
        ovr_b = shell['ovr_b']
        is_link_mode = shell.get('is_link_mode', False)
        link_instance_id = shell.get('link_instance_id')
        cat_a_id = shell.get('cat_a_id')
        cat_b_id = shell.get('cat_b_id')
        cat_b_ids = list(shell.get('cat_b_ids') or [])
        if not cat_b_ids and cat_b_id is not None:
            cat_b_ids = [cat_b_id]
        uniform_bbox = shell.get('uniform_bbox')

        if uniform_bbox is not None and not shell.get('section_box_set'):
            try:
                view.SetSectionBox(uniform_bbox)
                try:
                    view.IsSectionBoxActive = True
                except Exception:
                    pass
            except Exception as ex:
                logger.debug("SetSectionBox failed for {} / {}: {}".format(
                    pair_key, level_name, ex))

        isolate_list = List[ElementId]()
        a_eids = []
        b_eids = []

        for v in a_id_values:
            eid = make_element_id(v)
            if doc.GetElement(eid) is None:
                continue
            a_eids.append(eid)
            isolate_list.Add(eid)

        if not is_link_mode:
            for v in b_id_values:
                eid = make_element_id(v)
                if doc.GetElement(eid) is None:
                    continue
                b_eids.append(eid)
                isolate_list.Add(eid)
        # Link mode: isolate host clash elements only. Adding the whole
        # RvtLinkInstance forces generateViewSpecificGRep for every link view
        # (journal crash: Arr.cpp Invalid array after mass link regen).

        if isolate_list.Count > 0:
            try:
                view.IsolateElementsTemporary(isolate_list)
            except Exception as ex:
                if DEBUG_MODE:
                    logger.debug("IsolateElementsTemporary failed: {}".format(ex))
            try:
                view.ConvertTemporaryHideIsolateToPermanent()
            except Exception as ex:
                if DEBUG_MODE:
                    logger.debug(
                        "ConvertTemporaryHideIsolateToPermanent failed: {}".format(ex))

        if is_link_mode and link_instance_id is not None:
            _ensure_link_instance_visible(view, link_instance_id, doc)

        try:
            _set_annotation_categories_visible(view, False)
        except Exception as ex:
            logger.debug("Failed to hide annotations: {}".format(ex))

        if is_link_mode:
            keep_values = []
            keep_cats = [cat_a_id] + list(cat_b_ids or [])
            if cat_b_id is not None:
                keep_cats.append(cat_b_id)
            for cid in keep_cats:
                if cid is not None:
                    try:
                        keep_values.append(get_element_id_value(cid))
                    except Exception:
                        pass
            try:
                _hide_non_target_model_categories(view, keep_values)
            except Exception as ex:
                logger.debug("_hide_non_target_model_categories failed: {}".format(ex))
            if link_instance_id is not None:
                _ensure_link_instance_visible(view, link_instance_id, doc)

        try:
            _hide_scope_boxes(view)
        except Exception as ex:
            logger.debug("Failed to hide scope boxes: {}".format(ex))

        _apply_clash_view_colors(
            view, ovr_a, ovr_b,
            is_link_mode=is_link_mode,
            link_instance_id=link_instance_id,
            cat_a_id=cat_a_id,
            cat_b_id=cat_b_id,
            a_eids=a_eids,
            b_eids=b_eids,
            cat_b_ids=cat_b_ids,
        )

    def _create_clash_view(self, pair_key, level_name,
                           a_id_values, b_id_values, uniform_bbox,
                           view_type_id, scale, name_prefix, iteration,
                           ovr_a, ovr_b, is_link_mode=False,
                           link_instance_id=None,
                           cat_a_id=None, cat_b_id=None):
        """Create and configure an isolated isometric 3D clash view (single step)."""
        shell = self._create_clash_view_shell(
            pair_key, level_name,
            view_type_id, scale, name_prefix, iteration,
            a_id_values, b_id_values,
            uniform_bbox=uniform_bbox)
        if shell is None:
            return None
        shell['uniform_bbox'] = uniform_bbox
        shell['ovr_a'] = ovr_a
        shell['ovr_b'] = ovr_b
        shell['is_link_mode'] = is_link_mode
        shell['link_instance_id'] = link_instance_id
        shell['cat_a_id'] = cat_a_id
        shell['cat_b_id'] = cat_b_id
        shell['cat_b_ids'] = [cat_b_id] if cat_b_id is not None else []
        self._configure_clash_view(shell)
        return shell['view']

    # -------------------- Grid layout ------------------------------------

    _ISO_COS30 = 0.8660254037844386
    _ISO_SIN30 = 0.5

    def _uniform_viewport_size(self, scale):
        dx = getattr(self, '_uniform_dx', 0.0) or 0.0
        dy = getattr(self, '_uniform_dy', 0.0) or 0.0
        dz = getattr(self, '_uniform_dz', 0.0) or 0.0
        sc = float(scale) if scale else 1.0
        if sc <= 0:
            sc = 1.0
        if dx <= 0 and dy <= 0 and dz <= 0:
            return (0.5, 0.5)
        paper_width_model = (dx + dy) * self._ISO_COS30
        paper_height_model = dz + (dx + dy) * self._ISO_SIN30
        w = paper_width_model / sc
        h = paper_height_model / sc
        if w < 0.05: w = 0.05
        if h < 0.05: h = 0.05
        return (w, h)

    def _calculate_grid_layout(self, grouped_items, cell_width, cell_height, sheet_margin):
        total_viewports = len(self.clash_views)
        cols = min(self.GRID_COLUMNS, total_viewports)
        if cols < 1:
            cols = 1
        content_width = cols * cell_width

        row_count = 0
        col = 0
        for pair_key, by_level in grouped_items.items():
            for level_name, _ in by_level.items():
                if (pair_key, level_name) not in self.clash_views:
                    continue
                if col >= cols:
                    col = 0
                    row_count += 1
                col += 1
            if col > 0:
                col = 0
                row_count += 1 + self.GROUP_GAP_ROWS

        row_count = max(1, row_count - self.GROUP_GAP_ROWS)
        content_height = row_count * cell_height
        sheet_width = content_width + 2 * sheet_margin
        sheet_height = content_height + 2 * sheet_margin
        sheet_width = max(sheet_width, 1.5)
        sheet_height = max(sheet_height, 1.0)
        start_x = sheet_margin + cell_width / 2
        start_y = sheet_height - sheet_margin - cell_height / 2
        return (cols, row_count, sheet_width, sheet_height, start_x, start_y)

    def _calculate_viewport_positions(self, grouped_items, cols,
                                      cell_width, cell_height, start_x, start_y):
        positions = []
        col = 0
        row = 0
        for pair_key, by_level in grouped_items.items():
            for level_name, _ in by_level.items():
                if (pair_key, level_name) not in self.clash_views:
                    continue
                if col >= cols:
                    col = 0
                    row += 1
                x = start_x + col * cell_width
                y = start_y - row * cell_height
                positions.append((pair_key, level_name, x, y))
                col += 1
            if col > 0:
                col = 0
                row += 1 + self.GROUP_GAP_ROWS
        return positions

    def _create_sheet_and_place_viewports(self, grouped_items, sheet_prefix,
                                          iteration, sheet_name, scale):
        vp_w, vp_h = self._uniform_viewport_size(scale)

        if vp_w < 0.05:
            vp_w = 0.05
        if vp_h < 0.05:
            vp_h = 0.05

        cell_padding = 0.12
        cell_width = vp_w + self.VIEWPORT_SPACING + cell_padding
        cell_height = vp_h + self.VIEWPORT_SPACING + cell_padding
        sheet_margin = 0.15

        cols, row_count, sheet_width, sheet_height, start_x, start_y = \
            self._calculate_grid_layout(
                grouped_items, cell_width, cell_height, sheet_margin)

        positions = self._calculate_viewport_positions(
            grouped_items, cols, cell_width, cell_height, start_x, start_y)

        if not positions:
            logger.warning("No positions calculated for clash viewports")
            return

        sheet = ViewSheet.Create(doc, ElementId.InvalidElementId)
        num_prefix = sheet_prefix + (iteration + "-" if iteration else "")
        assigned = False
        for i in range(1, 1000):
            try:
                sheet.SheetNumber = "{}{:03d}".format(num_prefix, i)
                assigned = True
                break
            except Exception:
                continue
        if not assigned:
            logger.warning("Could not assign a unique sheet number with prefix {!r}".format(
                num_prefix))
            try:
                doc.Delete(sheet.Id)
            except Exception:
                pass
            return
        try:
            sheet.Name = sheet_name
        except Exception:
            pass
        self.created_sheets.append(sheet)

        _vp_placed = 0
        _vp_failed = 0
        for pair_key, level_name, x, y in positions:
            view = self.clash_views.get((pair_key, level_name))
            if view is None:
                continue
            try:
                Viewport.Create(doc, sheet.Id, view.Id, XYZ(x, y, 0))
                _vp_placed += 1
            except Exception as ex:
                _vp_failed += 1
                logger.warning("Could not place viewport for {} / {}: {}".format(
                    pair_key, level_name, ex))


# =============================================================================
# Annotation categories (for hide/show toggle)
# =============================================================================

_ANNOTATION_CATEGORIES = {
    "OST_Dimensions", "OST_TextNotes", "OST_Tags", "OST_SpotElevations",
    "OST_SpotCoordinates", "OST_SpotSlopes", "OST_AnnotationCrop",
    "OST_AnnotationCutlines", "OST_AnnotationObjects", "OST_ReferenceViewer",
    "OST_ReferenceViewerSymbol", "OST_Viewports", "OST_TitleBlocks",
    "OST_GridHeads", "OST_LevelHeads", "OST_SectionHeads", "OST_ElevationMarks",
    "OST_CalloutHeads", "OST_CropBoundary", "OST_CropRegions",
    "OST_Annotation_SketchLines", "OST_Annotation_Lines", "OST_CenterLines",
    "OST_HiddenLines", "OST_DemolishedLines", "OST_OverheadLines",
    "OST_Lines", "OST_Curves", "OST_CurveGroups",
    "OST_Levels", "OST_ScopeBoxes", "OST_VolumeOfInterest",
}

# Revit Scope Boxes are OST_VolumeOfInterest; OST_ScopeBoxes exists in some builds.
_SCOPE_BOX_CATEGORIES = {
    "OST_VolumeOfInterest",
    "OST_ScopeBoxes",
}


def _hide_scope_boxes(
        view,
        _Category=Category,
        _BuiltInCategory=BuiltInCategory):
    """Force-hide scope boxes in *view*. Must run inside a transaction."""
    if view is None:
        return 0
    document = view.Document
    hidden = 0
    seen = set()
    for bic_name in ("OST_VolumeOfInterest", "OST_ScopeBoxes"):
        try:
            bic = getattr(_BuiltInCategory, bic_name, None)
            if bic is None:
                continue
            cat = _Category.GetCategory(document, bic)
            if cat is None:
                continue
            cid = cat.Id
            cid_key = int(get_element_id_value(cid))
            if cid_key in seen:
                continue
            seen.add(cid_key)
            if view.CanCategoryBeHidden(cid):
                view.SetCategoryHidden(cid, True)
                hidden += 1
        except Exception:
            continue
    try:
        for cat in document.Settings.Categories:
            try:
                if _bic_key(cat.BuiltInCategory) not in _SCOPE_BOX_CATEGORIES:
                    continue
                cid = cat.Id
                cid_key = int(get_element_id_value(cid))
                if cid_key in seen:
                    continue
                seen.add(cid_key)
                if view.CanCategoryBeHidden(cid):
                    view.SetCategoryHidden(cid, True)
                    hidden += 1
            except Exception:
                continue
    except Exception:
        pass
    return hidden


def _set_annotation_categories_visible(view, visible):
    """Show or hide all annotation categories in the given view.

    Returns (shown_count, hidden_count).
    """
    document = view.Document
    shown = 0
    hidden = 0
    for cat in document.Settings.Categories:
        try:
            bic = cat.BuiltInCategory
            bic_str = _bic_key(bic)
        except Exception:
            continue
        if bic_str not in _ANNOTATION_CATEGORIES:
            continue
        # Scope boxes stay hidden even when the annotation toggle is on.
        if bic_str in _SCOPE_BOX_CATEGORIES:
            try:
                if view.CanCategoryBeHidden(cat.Id):
                    view.SetCategoryHidden(cat.Id, True)
                    hidden += 1
            except Exception:
                pass
            continue
        try:
            if not view.CanCategoryBeHidden(cat.Id):
                continue
        except Exception:
            continue
        try:
            view.SetCategoryHidden(cat.Id, not visible)
            if visible:
                shown += 1
            else:
                hidden += 1
        except Exception:
            pass
    return (shown, hidden)


# =============================================================================
# Refinement progress dialog
# =============================================================================

class RefinementProgressWindow(forms.WPFWindow):
    """Indeterminate progress window shown during link refinement."""

    def __init__(self, total_views):
        xaml_content = '''<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Updating clash views" Width="480" Height="200"
        WindowStartupLocation="CenterScreen"
        ShowActivated="False"
        Focusable="False"
        Topmost="True"
        Background="{DynamicResource WindowBackgroundBrush}"
        ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Hiding non-clash elements in linked models..."
                   Style="{DynamicResource HeaderTextBlockStyle}" Margin="0,0,0,10"/>

        <TextBlock Grid.Row="1" x:Name="statusText"
                   Text="View 0 of 0"
                   Style="{DynamicResource BodyTextBlockStyle}"
                   Foreground="{DynamicResource TextSecondaryBrush}" Margin="0,0,0,15"/>

        <ProgressBar Grid.Row="2" x:Name="progressBar" Height="16"
                     IsIndeterminate="True"
                     Background="{DynamicResource ControlBackgroundBrush}"
                     Foreground="{DynamicResource AccentBrush}"/>

        <TextBlock Grid.Row="3" TextWrapping="Wrap" Margin="0,15,0,0"
                   Style="{DynamicResource SecondaryTextBlockStyle}" FontSize="11" LineHeight="16">
            <Run FontWeight="SemiBold">What is happening:</Run>
            <LineBreak/>
            <Run>Revit is hiding elements in the linked model that are not part of a clash. Only clashing elements stay visible.</Run>
            <LineBreak/>
            <Run>The tool switches between views automatically. Please wait.</Run>
        </TextBlock>
    </Grid>
</Window>'''
        # Write XAML to temp file
        import tempfile
        xaml_path = tempfile.mktemp(suffix='.xaml')
        with open(xaml_path, 'w') as f:
            f.write(xaml_content)
        forms.WPFWindow.__init__(self, xaml_path)
        self._total_views = total_views
        self._update_status(0)

        try:
            from styles import load_styles_to_window
            load_styles_to_window(self)
        except Exception:
            pass

    def _update_status(self, current):
        try:
            if hasattr(self, 'statusText') and self.statusText:
                self.statusText.Text = "View {} of {}".format(current, self._total_views)
        except Exception:
            pass

    def update_progress(self, current):
        """Update the status text with current progress."""
        self._update_status(current)


# =============================================================================
# Summary dialog
# =============================================================================

class ClashViewsSummaryWindow(forms.WPFWindow):
    """Summary dialog shown after clash views are created."""

    def __init__(self, sheet, view_names, total_clashes, clash_pairs,
                 marker_data=None, tool_window=None):
        """
        Args:
            sheet: The created ViewSheet
            view_names: List of created view names
            total_clashes: Total number of clash pairs detected
            clash_pairs: List of (pair_key, count) tuples
            marker_data: Optional dict with available flag and view_points map
            tool_window: ClashViewsWindow that ran create (kept for Refresh)
        """
        xaml_file = op.join(pushbutton_dir, "ClashViewsSummary.xaml")
        forms.WPFWindow.__init__(self, xaml_file)

        try:
            from styles import load_styles_to_window
            load_styles_to_window(self)
        except Exception as ex:
            logger.debug("Could not load styles: {}".format(ex))

        self._tool_window = tool_window
        self._refresh_busy = False
        self._marker_driver = None
        self._annotations_visible = False
        self.Closing += self._on_window_closing
        self._bind_summary_data(
            sheet, view_names, total_clashes, clash_pairs, marker_data or {},
            start_markers=True)

    def _bind_summary_data(self, sheet, view_names, total_clashes, clash_pairs,
                           marker_data, start_markers=True):
        """Fill header, list, and marker toggle from a run result."""
        self._sheet = sheet
        self._view_names = list(view_names or [])
        self._total_clashes = total_clashes
        self._clash_pairs = list(clash_pairs or [])
        marker_data = marker_data or {}
        self._markers_available = bool(marker_data.get('available'))
        self._marker_view_points = dict(marker_data.get('view_points') or {})

        try:
            sheet_num = sheet.SheetNumber if sheet is not None else "-"
        except Exception:
            sheet_num = "-"
        try:
            sheet_name = sheet.Name if sheet is not None else "-"
        except Exception:
            sheet_name = "-"
        self.sheetInfoText.Text = "{} - {}".format(sheet_num, sheet_name)
        self.viewsCountText.Text = str(len(self._view_names))
        self.clashesCountText.Text = str(total_clashes)

        mode_text = "host-only"
        for p, _count in self._clash_pairs:
            if "[Link:" in p:
                mode_text = "cross-document (host vs link)"
                break
        self.summaryHeaderText.Text = (
            "Created {} clash view(s) in {} mode. "
            "{} unique element pair(s) detected across {} category combination(s).".format(
                len(self._view_names), mode_text, total_clashes,
                len(self._clash_pairs)))

        self._populate_views_list()
        self._setup_markers_toggle(start_markers=start_markers)

        self._annotations_visible = False
        if hasattr(self, "annotationsToggle") and self.annotationsToggle:
            self.annotationsToggle.IsChecked = False

    def _setup_markers_toggle(self, start_markers=True):
        if not (hasattr(self, "markersToggle") and self.markersToggle):
            return
        if not self._markers_available:
            self.markersToggle.IsChecked = False
            self.markersToggle.IsEnabled = False
            if hasattr(self, "markersHelpText") and self.markersHelpText:
                self.markersHelpText.Text = (
                    "Clash markers require Revit 2022 or newer "
                    "(TemporaryGraphicsManager API).")
            return
        if self._marker_view_points:
            self.markersToggle.IsEnabled = True
            toggle_on = True
            if is_marker_toggle_active is not None:
                toggle_on = is_marker_toggle_active(doc)
            self.markersToggle.IsChecked = toggle_on
            if start_markers and toggle_on:
                self._start_marker_driver()
        else:
            self.markersToggle.IsChecked = False
            self.markersToggle.IsEnabled = False

    def apply_run_results(self, data, _restore=_restore_clash_globals):
        """Replace summary contents after a Refresh rebuild."""
        _restore()
        data = data or {}
        self._stop_marker_driver()
        self._bind_summary_data(
            data.get('sheet'),
            data.get('view_names') or [],
            data.get('total_clashes') or 0,
            data.get('clash_pairs') or [],
            data.get('marker_data') or {},
            start_markers=True)
        self._refresh_busy = False
        self._reset_refresh_button()

    def _reset_refresh_button(self):
        if hasattr(self, "refreshButton") and self.refreshButton:
            self.refreshButton.IsEnabled = True
            self.refreshButton.Content = "Refresh clashes"

    def refreshButton_Click(self, sender, args):
        """Re-run clash detection with the same categories and settings."""
        import sys as _sys
        _fn = getattr(_sys, '_pyBS_clash_restore', None)
        if _fn is not None:
            _fn()
        if getattr(self, '_refresh_busy', False):
            return
        tool = getattr(self, '_tool_window', None)
        if tool is None:
            tool = getattr(_sys, '_pyBS_clash_tool_window', None)
        if tool is None or not hasattr(tool, 'refresh_clash_views'):
            forms.alert(
                "Cannot refresh: original clash run is no longer available.",
                title="Refresh")
            return

        self._refresh_busy = True
        if hasattr(self, "refreshButton") and self.refreshButton:
            self.refreshButton.IsEnabled = False
            self.refreshButton.Content = "Refreshing..."
        try:
            self._stop_marker_driver()
            status = tool.refresh_clash_views(self)
        except Exception as ex:
            logger.error("Refresh clashes failed: {}".format(ex))
            import traceback
            logger.debug(traceback.format_exc())
            forms.alert("Refresh failed: {}".format(ex), title="Error")
            self._refresh_busy = False
            self._reset_refresh_button()
            if hasattr(self, "markersToggle") and self.markersToggle \
                    and bool(self.markersToggle.IsChecked):
                self._start_marker_driver()
            return

        # pending_refine: apply_run_results runs when the second pass finishes
        if status == 'pending_refine':
            return
        if status != 'ok':
            self._refresh_busy = False
            self._reset_refresh_button()
            if hasattr(self, "markersToggle") and self.markersToggle \
                    and bool(self.markersToggle.IsChecked):
                self._start_marker_driver()
        elif getattr(self, '_refresh_busy', False):
            self._refresh_busy = False
            self._reset_refresh_button()

    def _populate_views_list(self):
        """Add view entries to the scrollable list."""
        import clr
        clr.AddReference("PresentationFramework")
        from System.Windows.Controls import TextBlock, Separator
        from System.Windows import Thickness
        from System.Windows.Media import FontWeights

        self.viewsListPanel.Children.Clear()

        for i, name in enumerate(self._view_names, 1):
            # View name with number
            tb = TextBlock()
            tb.Text = "{}. {}".format(i, name)
            tb.TextWrapping = True
            tb.Margin = Thickness(0, 2, 0, 2)
            tb.FontSize = 12
            self.viewsListPanel.Children.Add(tb)

            # Add separator except for last item
            if i < len(self._view_names):
                sep = Separator()
                sep.Margin = Thickness(0, 4, 0, 4)
                self.viewsListPanel.Children.Add(sep)

    def AnnotationsToggle_Changed(self, sender, args):
        """Handle annotation visibility toggle."""
        import sys as _sys
        _fn = getattr(_sys, '_pyBS_clash_restore', None)
        if _fn is not None:
            _fn()
        is_checked = bool(self.annotationsToggle.IsChecked)
        if is_checked == self._annotations_visible:
            return
        self._annotations_visible = is_checked

        if self._sheet is None:
            return

        # Update all created views
        with Transaction(doc, "Toggle Annotations Visibility") as t:
            t.Start()
            try:
                viewport_ids = self._sheet.GetAllViewports()
                for vp_id in viewport_ids:
                    try:
                        viewport = doc.GetElement(vp_id)
                        if viewport is not None:
                            view_id = viewport.ViewId
                            view_obj = doc.GetElement(view_id)
                            if view_obj is not None:
                                _set_annotation_categories_visible(view_obj, is_checked)
                                _hide_scope_boxes(view_obj)
                    except Exception as ex:
                        logger.debug("Failed to toggle annotations for viewport: {}".format(ex))
            except Exception as ex:
                logger.debug("Failed to get viewports: {}".format(ex))
            t.Commit()

        status = "shown" if is_checked else "hidden"
        logger.debug("Annotations {} in all clash views".format(status))

    def _start_marker_driver(self, _restore=_restore_clash_globals):
        _restore()
        if not _CLASH_MARKERS_AVAILABLE or start_or_get_driver is None:
            return
        if not self._marker_view_points:
            return
        try:
            driver = start_or_get_driver(
                __revit__, doc, self._marker_view_points,
                get_element_id_value, logger=None)
            if driver is None:
                return
            enabled = True
            if hasattr(self, "markersToggle") and self.markersToggle:
                enabled = bool(self.markersToggle.IsChecked)
            driver.set_enabled(enabled)
            self._marker_driver = driver
            if set_marker_toggle_active is not None and hasattr(self, "markersToggle"):
                set_marker_toggle_active(doc, enabled)
        except Exception as ex:
            logger.warning("Failed to start clash marker driver: {}".format(ex))

    def _stop_marker_driver(self, _restore=_restore_clash_globals):
        """Stop global clash marker session (ribbon OFF / explicit cleanup)."""
        _restore()
        if clean_marker_session is not None:
            clean_marker_session(doc)
        else:
            if self._marker_driver is not None:
                try:
                    self._marker_driver.stop()
                except Exception:
                    pass
        self._marker_driver = None

    def _release_marker_driver_ref(self):
        """Drop summary-window ref without stopping the shared session driver."""
        self._marker_driver = None

    def MarkersToggle_Changed(self, sender, args):
        import sys as _sys
        _fn = getattr(_sys, '_pyBS_clash_restore', None)
        if _fn is not None:
            _fn()
        if not self._markers_available:
            return
        is_checked = bool(self.markersToggle.IsChecked)
        if set_marker_toggle_active is not None:
            set_marker_toggle_active(doc, is_checked)
        if not is_checked:
            self._stop_marker_driver()
        else:
            if self._marker_driver is None and find_marker_driver is not None:
                self._marker_driver = find_marker_driver(doc)
            if self._marker_driver is None:
                self._start_marker_driver()
            elif self._marker_driver is not None:
                self._marker_driver.set_enabled(True)

    def _on_window_closing(self, sender, args):
        self._release_marker_driver_ref()

    def closeButton_Click(self, sender, args):
        """Close the dialog."""
        self._release_marker_driver_ref()
        self.Close()


# =============================================================================
# Entry point
# =============================================================================

if __name__ == '__main__':
    if _default_three_d_view_type_id(doc) is None:
        forms.alert("No 3D view family types found in the project.", title="Error")
    else:
        discovered = _discover_clashable_categories(doc)
        if not discovered:
            forms.alert("No clashable model categories with instances were found.",
                        title="No Categories")
        else:
            window = ClashViewsWindow(host_categories=discovered)
            window.ShowDialog()
