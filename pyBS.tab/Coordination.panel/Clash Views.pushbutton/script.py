# -*- coding: utf-8 -*-
"""Detect clashes between Revit categories and place isolated 3D views on a sheet."""

__title__ = "Clash\nViews"
__author__ = "pyByggstyrning"
__highlight__ = "new"
__doc__ = """Pick two or more model categories. All unique pairs (C(N,2)) are clash-detected
using a bounding-box pre-filter plus geometric confirmation in the host document.
Each pair (sub-grouped by level) becomes an isolated 3D view; every view is placed
end-to-end on a single sheet in a grid layout.

Optionally clash against a linked model: enable "Against link model", pick a link,
select link categories. Pairs are then host-category x link-category (cross-product).
Cross-doc clash uses bbox-only intersection (no geometric confirmation across docs)."""

import clr
clr.AddReference("System")
clr.AddReference("System.Collections")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")

from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection

from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementId, BuiltInCategory,
    ViewFamilyType, ViewFamily, View3D, ViewSheet, Viewport,
    XYZ, BoundingBoxXYZ, Outline,
    BoundingBoxIntersectsFilter, ElementIntersectsElementFilter,
    UnitUtils, Transaction, TransactionGroup, BuiltInParameter,
    OverrideGraphicSettings, FillPatternElement,
    Color as RevitColor,
    RevitLinkInstance, RevitLinkType,
)

import sys
import os.path as op
from collections import OrderedDict
from itertools import combinations

from pyrevit import revit, DB
from pyrevit import forms, script

script_path = __file__
pushbutton_dir = op.dirname(script_path)
panel_dir = op.dirname(pushbutton_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from revit.compat import get_element_id_value, make_element_id

logger = script.get_logger()

DEBUG_MODE = False

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# =============================================================================
# Category discovery
# =============================================================================

# Categories without physical geometry that should never be clashed
_NON_GEOMETRIC_CATEGORIES = {
    "OST_Grids", "OST_Levels", "OST_SectionBox", "OST_Viewers",
    "OST_CLines", "OST_SiteProperty", "OST_SitePropertyLineSegment",
    "OST_Cameras", "OST_Sections", "OST_Elev", "OST_AreaSchemes",
    "OST_Matchline", "OST_ReferenceLines", "OST_SketchLines",
    "OST_ScheduleGraphics", "OST_CenterLines", "OST_DecalElement",
}


def _is_clashable_category(cat):
    """Heuristic filter: keep model categories that can have 3D geometry."""
    if not cat:
        return False
    try:
        if cat.CategoryType != DB.CategoryType.Model:
            return False
    except Exception:
        return False

    try:
        bic = cat.BuiltInCategory
        if bic is None:
            return False
        if str(bic) in _NON_GEOMETRIC_CATEGORIES:
            return False
    except Exception:
        pass

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


def _discover_clashable_categories(document):
    """Return sorted list of (name, BuiltInCategory) for model categories with instances."""
    results = []
    seen_names = set()
    try:
        for cat in document.Settings.Categories:
            if not _is_clashable_category(cat):
                continue
            bic = _get_category_bic(cat)
            if bic is None:
                continue
            try:
                first_id = (FilteredElementCollector(document)
                            .OfCategory(bic)
                            .WhereElementIsNotElementType()
                            .FirstElementId())
                if first_id == ElementId.InvalidElementId:
                    continue
            except Exception:
                continue

            name = cat.Name
            if not name or name in seen_names:
                continue
            seen_names.add(name)
            results.append((name, bic))
    except Exception as ex:
        logger.warning("Error discovering categories: {}".format(ex))

    results.sort(key=lambda pair: pair[0].lower())
    return results


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
# Color overrides
# =============================================================================

CLASH_COLOR_A = RevitColor(230, 50, 50)
CLASH_COLOR_B = RevitColor(40, 170, 70)


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
    """Build an OverrideGraphicSettings that paints the element solid with color."""
    ogs = OverrideGraphicSettings()
    try:
        ogs.SetProjectionLineColor(color)
        ogs.SetCutLineColor(color)
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
            candidates = (FilteredElementCollector(document, b_ids_list)
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
            pairs.append((a, b))

    return pairs


# =============================================================================
# Clash detection – host vs linked model (bbox only)
# =============================================================================

def _clash_pairs_with_link(document, bic_host, link_instance, link_doc, bic_link):
    """Detect clashing element pairs between host category and linked-model category.

    ElementIntersectsElementFilter cannot cross documents, so this uses
    bounding-box intersection only (transformed to host coordinate space).

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
            candidates = (FilteredElementCollector(document, host_ids)
                          .WherePasses(bbox_filter)
                          .ToElements())
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("bbox filter failed for link elem {}: {}".format(link_el.Id, ex))
            continue

        for host_el in candidates:
            link_id_val = get_element_id_value(link_el.Id)
            host_id_val = get_element_id_value(host_el.Id)
            key = (host_id_val, link_id_val)
            if key in seen:
                continue
            seen.add(key)
            pairs.append((host_el, link_el))

    return pairs


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


# =============================================================================
# Main window
# =============================================================================

class ClashViewsWindow(forms.WPFWindow):
    """WPF window for creating category clash views on a sheet."""

    GRID_COLUMNS = 20
    GROUP_GAP_ROWS = 1
    VIEWPORT_SPACING = 0.03

    def __init__(self):
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

        self._populate_category_filters()
        self._sync_category_list_display()
        self.categoryListBox.ItemsSource = self._category_display
        self._update_category_selection_count()

        # Link selector
        self.linkSelectorComboBox.ItemsSource = self._link_items
        self.linkCategoryListBox.ItemsSource = self._link_category_display

    # -------------------- Host category UI setup --------------------------

    def _populate_category_filters(self):
        self._category_items.Clear()
        for name, bic in _discover_clashable_categories(doc):
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

        self.Close()

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
                )

            if self.created_sheets:
                uidoc.ActiveView = self.created_sheets[0]

        except Exception as ex:
            logger.error("Error creating clash views: {}".format(ex))
            import traceback
            logger.debug(traceback.format_exc())
            forms.alert("Error creating clash views: {}".format(ex), title="Error")

    # -------------------- Orchestration ----------------------------------

    def _create_clash_views(self, categories, scale, view_type_id,
                            name_prefix, sheet_prefix, sheet_name, iteration,
                            group_by_level, progress_bar=None,
                            link_instance=None, link_doc=None, link_categories=None):
        """Run clash detection for all category pairs, create views, then place on a sheet.

        When link_instance is provided, pairs are host-cat x link-cat cross-product
        and clash detection uses bbox-only intersection in host coordinate space.
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

            by_level = OrderedDict()
            for (a, b) in pairs:
                if group_by_level:
                    level_key = _element_level_name(doc, a)
                else:
                    level_key = "All"
                if level_key not in by_level:
                    by_level[level_key] = {"a_ids": set(), "b_ids": set(), "link_mode": link_mode}
                by_level[level_key]["a_ids"].add(get_element_id_value(a.Id))
                if link_mode:
                    by_level[level_key]["b_ids"].add(get_element_id_value(b.Id))
                else:
                    by_level[level_key]["b_ids"].add(get_element_id_value(b.Id))

            grouped[pair_key] = by_level
            total_clashes += len(pairs)

        if progress_bar:
            progress_bar.update_progress(len(pair_list), len(pair_list))

        if not grouped:
            forms.alert("No clashes detected between selected categories.",
                        title="No Clashes")
            return

        # Phase 1: compute bboxes and uniform size before the transaction
        natural_info = OrderedDict()
        max_dx = max_dy = max_dz = 0.0

        link_transform = None
        if link_mode:
            try:
                link_transform = link_instance.GetTotalTransform()
            except Exception:
                link_transform = None

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
                    "center": XYZ(
                        (bbox.Min.X + bbox.Max.X) / 2.0,
                        (bbox.Min.Y + bbox.Max.Y) / 2.0,
                        (bbox.Min.Z + bbox.Max.Z) / 2.0,
                    ),
                }

        if not natural_info:
            forms.alert("Clashes were detected but no usable bounding boxes.",
                        title="No Views")
            return

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

        with TransactionGroup(doc, "Create Category Clash Views") as tg:
            tg.Start()

            with Transaction(doc, "Create Clash 3D Views") as t:
                t.Start()
                processed = 0
                for (pair_key, level_name), info in natural_info.items():
                    processed += 1
                    if progress_bar:
                        try:
                            progress_bar.title = "Creating view {}/{}".format(
                                processed, total_views_to_create)
                        except Exception:
                            pass
                        progress_bar.update_progress(
                            processed, total_views_to_create)

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

                    try:
                        view = self._create_clash_view(
                            pair_key, level_name,
                            info["a_ids"], info["b_ids"], uniform,
                            view_type_id, scale, name_prefix, iteration,
                            ovr_a, ovr_b,
                            is_link_mode=info.get("link_mode", False),
                        )
                        if view:
                            self.clash_views[(pair_key, level_name)] = view
                    except Exception as ex:
                        logger.warning(
                            "Failed to create view for {} / {}: {}".format(
                                pair_key, level_name, ex))
                t.Commit()

            with Transaction(doc, "Create Sheet and Place Viewports") as t:
                t.Start()
                self._create_sheet_and_place_viewports(
                    grouped, sheet_prefix, iteration, sheet_name, scale)
                t.Commit()

            tg.Assimilate()

    # -------------------- View creation ----------------------------------

    def _create_clash_view(self, pair_key, level_name,
                           a_id_values, b_id_values, uniform_bbox,
                           view_type_id, scale, name_prefix, iteration,
                           ovr_a, ovr_b, is_link_mode=False):
        """Create an isolated isometric 3D view.

        Host mode: isolate A+B, color A=red, B=green.
        Link mode: isolate host (A) elements only (link elements cannot be isolated
        in host views), color A=red. Section box scopes the view to the clash area.
        """
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

        if not a_eids and not b_eids:
            return None

        view = View3D.CreateIsometric(doc, view_type_id)
        if view is None:
            return None

        try:
            view.SetSectionBox(uniform_bbox)
        except Exception as ex:
            logger.debug("SetSectionBox failed for {} / {}: {}".format(
                pair_key, level_name, ex))

        try:
            view.Scale = scale
        except Exception:
            pass

        if iteration:
            base_name = "{0}[{1}] {2} - {3}".format(
                name_prefix, iteration, pair_key, level_name)
        else:
            base_name = "{0}{1} - {2}".format(name_prefix, pair_key, level_name)
        try:
            view.Name = base_name
        except Exception:
            for i in range(1, 100):
                try:
                    view.Name = "{} ({})".format(base_name, i)
                    break
                except Exception:
                    continue

        if isolate_list.Count > 0:
            try:
                view.IsolateElementsTemporary(isolate_list)
            except Exception as ex:
                if DEBUG_MODE:
                    logger.debug("IsolateElementsTemporary failed: {}".format(ex))

            try:
                view.ConvertTemporaryHideIsolateToPermanent()
            except Exception:
                pass

        # Color overrides
        for eid in a_eids:
            try:
                view.SetElementOverrides(eid, ovr_a)
            except Exception:
                continue
        for eid in b_eids:
            try:
                view.SetElementOverrides(eid, ovr_b)
            except Exception:
                continue

        return view

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

        for pair_key, level_name, x, y in positions:
            view = self.clash_views.get((pair_key, level_name))
            if view is None:
                continue
            try:
                Viewport.Create(doc, sheet.Id, view.Id, XYZ(x, y, 0))
            except Exception as ex:
                logger.warning("Could not place viewport for {} / {}: {}".format(
                    pair_key, level_name, ex))


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
            window = ClashViewsWindow()
            window.ShowDialog()
