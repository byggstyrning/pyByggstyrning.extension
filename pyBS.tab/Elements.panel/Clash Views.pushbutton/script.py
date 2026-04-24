# -*- coding: utf-8 -*-
"""Clash selected categories pairwise and place one isolated 3D view per clashing pair on a sheet."""

__title__ = "Clash\nViews"
__author__ = "pyByggstyrning"
__highlight__ = "new"
__doc__ = """Create isolated 3D clash views for every pairwise combination of selected categories.

Pick two or more model categories. The tool clashes every unordered pair (A x B).
For each pair with clashes, it creates one isolated 3D view (section-boxed, temporarily
isolated to the clashing elements) and places every pair-view on a sheet in a grid layout.
"""

# Import .NET libraries
import clr
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Collections")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")

from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection

# Import Revit API
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInParameter,
    ViewFamilyType, ViewFamily, View, View3D, ViewSheet, Viewport,
    XYZ, ElementId, BoundingBoxXYZ, Outline,
    UnitUtils, Transaction, TransactionGroup,
    BoundingBoxIntersectsFilter, ElementIntersectsElementFilter,
    ElementCategoryFilter,
)

# Import pyRevit libraries
import sys
import os.path as op
from collections import OrderedDict
from pyrevit import DB
from pyrevit import forms, script

# Add the extension directory to the path
script_path = __file__
pushbutton_dir = op.dirname(script_path)
panel_dir = op.dirname(pushbutton_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from revit.compat import get_element_id_value

logger = script.get_logger()

DEBUG_MODE = False

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# =============================================================================
# Helpers
# =============================================================================

def _is_model_category(cat):
    """True if category is a model (non-annotation, non-internal) category usable for clash."""
    if not cat:
        return False
    try:
        if cat.CategoryType != DB.CategoryType.Model:
            return False
    except Exception:
        return False
    try:
        if cat.Id is None:
            return False
    except Exception:
        return False
    return True


def _discover_clashable_categories():
    """Return OrderedDict name -> Category for model categories that have >=1 instance."""
    found = {}
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for el in collector:
        try:
            cat = el.Category
            if not cat or not cat.Name:
                continue
            if not _is_model_category(cat):
                continue
            if cat.Name not in found:
                found[cat.Name] = cat
        except Exception:
            continue
    ordered = OrderedDict()
    for name in sorted(found.keys()):
        ordered[name] = found[name]
    return ordered


def _expand_bbox(bbox, pad):
    """Return new BoundingBoxXYZ padded by `pad` on each side."""
    out = BoundingBoxXYZ()
    out.Min = XYZ(bbox.Min.X - pad, bbox.Min.Y - pad, bbox.Min.Z - pad)
    out.Max = XYZ(bbox.Max.X + pad, bbox.Max.Y + pad, bbox.Max.Z + pad)
    return out


def _union_bbox(bbox_a, bbox_b):
    """Return BoundingBoxXYZ enclosing both input bboxes."""
    out = BoundingBoxXYZ()
    out.Min = XYZ(
        min(bbox_a.Min.X, bbox_b.Min.X),
        min(bbox_a.Min.Y, bbox_b.Min.Y),
        min(bbox_a.Min.Z, bbox_b.Min.Z),
    )
    out.Max = XYZ(
        max(bbox_a.Max.X, bbox_b.Max.X),
        max(bbox_a.Max.Y, bbox_b.Max.Y),
        max(bbox_a.Max.Z, bbox_b.Max.Z),
    )
    return out


def _collect_union_bbox(element_ids):
    """Union bbox over elements by Id. Returns None if no valid bbox."""
    out = None
    for eid in element_ids:
        el = doc.GetElement(eid)
        if not el:
            continue
        try:
            bb = el.get_BoundingBox(None)
        except Exception:
            bb = None
        if not bb:
            continue
        if out is None:
            out = BoundingBoxXYZ()
            out.Min = XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z)
            out.Max = XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z)
        else:
            out = _union_bbox(out, bb)
    return out


# =============================================================================
# Data Classes
# =============================================================================

class CategoryFilterItem(forms.Reactive):
    """One category row in the category filter list."""

    def __init__(self, category_name, is_selected=False):
        super(CategoryFilterItem, self).__init__()
        self._category_name = category_name
        self._is_selected = is_selected

    @property
    def CategoryName(self):
        return self._category_name

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        self.OnPropertyChanged("IsSelected")


class TemplateItem:
    """Wrapper for a 3D view template."""

    def __init__(self, view=None):
        self.view = view
        self.Name = view.Name if view else "<None>"
        self.Id = view.Id if view else None

    def __str__(self):
        return self.Name

    def __repr__(self):
        return self.Name


class ViewTypeItem:
    """Wrapper for a 3D ViewFamilyType."""

    def __init__(self, vft):
        self.vft = vft
        try:
            p = vft.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            self.Name = p.AsString() if p else vft.Name
        except Exception:
            self.Name = getattr(vft, "Name", "3D View")
        self.Id = vft.Id if vft else None


# =============================================================================
# Main Window
# =============================================================================

class ClashViewsWindow(forms.WPFWindow):
    """WPF window for building clash 3D views."""

    VIEWPORT_SPACING = 0.03
    GROUP_GAP_ROWS = 1

    def __init__(self):
        logger.debug("Initializing Clash Views window")

        xaml_file = op.join(pushbutton_dir, "ClashViewsWindow.xaml")
        forms.WPFWindow.__init__(self, xaml_file)

        try:
            from styles import load_styles_to_window
            load_styles_to_window(self)
        except Exception as ex:
            logger.debug("Could not load styles: {}".format(ex))

        self._categories = OrderedDict()
        self._category_items = ObservableCollection[CategoryFilterItem]()
        self._category_display = ObservableCollection[CategoryFilterItem]()

        self.pair_views = OrderedDict()
        self.created_sheets = []

        self._setup_scale_options()
        self._setup_templates()
        self._setup_view_types()
        self._populate_category_filters()
        self._sync_category_list_display()
        self.categoryListBox.ItemsSource = self._category_display
        self._update_category_selection_count()

    # -------------------------------------------------------------------------
    # Setup
    # -------------------------------------------------------------------------

    def _setup_scale_options(self):
        for scale in ["1:10", "1:20", "1:50", "1:100", "1:200"]:
            self.scaleComboBox.Items.Add(scale)
        self.scaleComboBox.SelectedIndex = 2

    def _setup_templates(self):
        self.templateComboBox.Items.Add(TemplateItem(None))
        all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
        templates = [v for v in all_views
                     if v.IsTemplate and v.ViewType == DB.ViewType.ThreeD]
        for template in sorted(templates, key=lambda x: x.Name):
            self.templateComboBox.Items.Add(TemplateItem(template))
        self.templateComboBox.SelectedIndex = 0

    def _setup_view_types(self):
        vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
        three_d_types = []
        for vft in vfts:
            try:
                if vft.ViewFamily == ViewFamily.ThreeDimensional:
                    three_d_types.append(vft)
            except Exception as ex:
                logger.debug("Error checking ViewFamilyType: {}".format(ex))

        for vft in three_d_types:
            self.viewTypeComboBox.Items.Add(ViewTypeItem(vft))

        if three_d_types:
            self.viewTypeComboBox.SelectedIndex = 0
        else:
            logger.warning("No 3D ViewFamilyType found in the project.")

    def _populate_category_filters(self):
        self._categories = _discover_clashable_categories()
        self._category_items.Clear()
        for name in self._categories.keys():
            self._category_items.Add(CategoryFilterItem(name, False))

    def _sync_category_list_display(self):
        """Refresh filtered category list (search) without changing check state."""
        self._category_display.Clear()
        search = ""
        if hasattr(self, "categorySearchTextBox") and self.categorySearchTextBox:
            search = (self.categorySearchTextBox.Text or "").strip().lower()
        for item in self._category_items:
            if not search or search in item.CategoryName.lower():
                self._category_display.Add(item)

    def _get_selected_category_names(self):
        return [i.CategoryName for i in self._category_items if i.IsSelected]

    def _update_category_selection_count(self):
        sel_names = self._get_selected_category_names()
        n_sel = len(sel_names)
        n_tot = self._category_items.Count
        pair_count = n_sel * (n_sel - 1) // 2
        if hasattr(self, "categorySelectionCountText") and self.categorySelectionCountText:
            self.categorySelectionCountText.Text = \
                "{} of {} categories selected ({} pairs to clash)".format(
                    n_sel, n_tot, pair_count)

    # -------------------------------------------------------------------------
    # Event Handlers
    # -------------------------------------------------------------------------

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

    def cancelButton_Click(self, sender, args):
        self.Close()

    def createButton_Click(self, sender, args):
        selected_names = self._get_selected_category_names()
        if len(selected_names) < 2:
            forms.alert("Select at least two categories to clash.", title="Warning")
            return

        if self.viewTypeComboBox.SelectedItem is None:
            forms.alert("No 3D view type available in the project.", title="Error")
            return

        scale = self._get_scale_value()
        view_type_id = self.viewTypeComboBox.SelectedItem.Id
        template_item = self.templateComboBox.SelectedItem
        template_id = template_item.Id if template_item and template_item.Id else None
        name_prefix = self.namePrefixTextBox.Text or "Clash - "
        sheet_prefix = self.sheetPrefixTextBox.Text or "CV-"
        sheet_name = self.sheetNameTextBox.Text or "Clash Views"

        try:
            crop_offset_mm = float(self.cropOffsetTextBox.Text or "300")
        except ValueError:
            crop_offset_mm = 300.0

        try:
            grid_cols = int(self.gridColumnsTextBox.Text or "20")
            if grid_cols < 1:
                grid_cols = 1
        except ValueError:
            grid_cols = 20

        try:
            crop_offset = UnitUtils.ConvertToInternalUnits(
                crop_offset_mm, DB.UnitTypeId.Millimeters)
        except Exception:
            crop_offset = crop_offset_mm / 304.8

        selected_cats = [(n, self._categories[n]) for n in selected_names
                         if n in self._categories]

        self.Close()

        try:
            title = "Creating Clash Views ({} categories)".format(len(selected_cats))
            with forms.ProgressBar(title=title) as pb:
                self._run(selected_cats, scale, view_type_id, template_id,
                          name_prefix, crop_offset, sheet_prefix, sheet_name,
                          grid_cols, pb)

            if self.created_sheets:
                uidoc.ActiveView = self.created_sheets[0]
        except Exception as ex:
            logger.error("Error creating clash views: {}".format(ex))
            import traceback
            logger.debug(traceback.format_exc())
            forms.alert("Error creating clash views: {}".format(ex), title="Error")

    # -------------------------------------------------------------------------
    # Core Run
    # -------------------------------------------------------------------------

    def _get_scale_value(self):
        scale_text = self.scaleComboBox.SelectedItem
        if scale_text:
            return int(scale_text.split(":")[1])
        return 50

    def _run(self, selected_cats, scale, view_type_id, template_id,
             name_prefix, crop_offset, sheet_prefix, sheet_name,
             grid_cols, progress_bar=None):
        """Clash every unordered pair; build isolated 3D views; place on sheet."""

        pairs = []
        for i in range(len(selected_cats)):
            for j in range(i + 1, len(selected_cats)):
                pairs.append((selected_cats[i], selected_cats[j]))

        total_pairs = len(pairs)
        if total_pairs == 0:
            forms.alert("No category pairs to clash.", title="Warning")
            return

        with TransactionGroup(doc, "Create Clash Views") as tg:
            tg.Start()

            with Transaction(doc, "Create Clash 3D Views") as t:
                t.Start()

                for idx, ((name_a, cat_a), (name_b, cat_b)) in enumerate(pairs, start=1):
                    if progress_bar:
                        progress_bar.update_progress(idx, total_pairs)

                    try:
                        clashing_ids = self._find_clashing_ids(cat_a, cat_b)
                    except Exception as ex:
                        logger.warning("Clash pass failed for {} x {}: {}".format(
                            name_a, name_b, ex))
                        clashing_ids = []

                    if not clashing_ids:
                        logger.debug("No clashes between {} and {}".format(name_a, name_b))
                        continue

                    try:
                        view = self._create_pair_view(
                            name_a, name_b, clashing_ids,
                            view_type_id, scale, name_prefix, template_id, crop_offset)
                    except Exception as ex:
                        logger.warning("Failed to create view for {} x {}: {}".format(
                            name_a, name_b, ex))
                        continue

                    if view is not None:
                        self.pair_views[(name_a, name_b)] = view

                t.Commit()

            if not self.pair_views:
                tg.RollBack()
                forms.alert("No clashes were found between the selected categories.",
                            title="No Clashes")
                return

            with Transaction(doc, "Create Clash Sheet and Place Viewports") as t:
                t.Start()
                self._create_sheet_and_place_viewports(
                    sheet_prefix, sheet_name, grid_cols)
                t.Commit()

            tg.Assimilate()

    # -------------------------------------------------------------------------
    # Clash detection
    # -------------------------------------------------------------------------

    def _collect_category_elements(self, category):
        """Return list of non-type elements for a category."""
        try:
            cat_filter = ElementCategoryFilter(category.Id)
            return list(FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .WherePasses(cat_filter)
                        .ToElements())
        except Exception as ex:
            logger.debug("Could not collect category {}: {}".format(
                getattr(category, "Name", "?"), ex))
            return []

    def _find_clashing_ids(self, cat_a, cat_b):
        """Return list of ElementIds (A + B) that participate in any clash between cats."""
        a_elements = self._collect_category_elements(cat_a)
        b_elements = self._collect_category_elements(cat_b)
        if not a_elements or not b_elements:
            return []

        b_id_list = List[ElementId]()
        for be in b_elements:
            b_id_list.Add(be.Id)
        if b_id_list.Count == 0:
            return []

        a_hit_ids = set()
        b_hit_ids = set()

        for a_elem in a_elements:
            try:
                a_bbox = a_elem.get_BoundingBox(None)
            except Exception:
                a_bbox = None
            if not a_bbox:
                continue

            try:
                outline = Outline(a_bbox.Min, a_bbox.Max)
                bbox_filter = BoundingBoxIntersectsFilter(outline)
            except Exception as ex:
                if DEBUG_MODE:
                    logger.debug("Outline/bbox filter failed for {}: {}".format(
                        get_element_id_value(a_elem.Id), ex))
                continue

            try:
                ei_filter = ElementIntersectsElementFilter(a_elem)
            except Exception as ex:
                if DEBUG_MODE:
                    logger.debug("ElementIntersectsElementFilter unsupported for {}: {}".format(
                        get_element_id_value(a_elem.Id), ex))
                continue

            try:
                candidates = (FilteredElementCollector(doc, b_id_list)
                              .WherePasses(bbox_filter)
                              .WherePasses(ei_filter)
                              .ToElements())
            except Exception as ex:
                if DEBUG_MODE:
                    logger.debug("Clash filter pass failed for {}: {}".format(
                        get_element_id_value(a_elem.Id), ex))
                continue

            if not candidates:
                continue

            a_hit_ids.add(a_elem.Id)
            for c in candidates:
                b_hit_ids.add(c.Id)

        out = List[ElementId]()
        seen_ints = set()
        for eid in a_hit_ids:
            key = get_element_id_value(eid)
            if key in seen_ints:
                continue
            seen_ints.add(key)
            out.Add(eid)
        for eid in b_hit_ids:
            key = get_element_id_value(eid)
            if key in seen_ints:
                continue
            seen_ints.add(key)
            out.Add(eid)

        return out

    # -------------------------------------------------------------------------
    # View creation
    # -------------------------------------------------------------------------

    def _create_pair_view(self, name_a, name_b, clashing_ids,
                          view_type_id, scale, name_prefix, template_id,
                          crop_offset):
        """Create one isolated 3D isometric view for a clashing category pair."""

        view = View3D.CreateIsometric(doc, view_type_id)
        if view is None:
            logger.error("View3D.CreateIsometric returned None for {} x {}".format(
                name_a, name_b))
            return None

        base_name = "{}{} x {}".format(name_prefix, name_a, name_b)
        try:
            view.Name = base_name
        except Exception:
            for i in range(1, 100):
                try:
                    view.Name = "{} ({})".format(base_name, i)
                    break
                except Exception:
                    continue

        if template_id:
            try:
                view.ViewTemplateId = template_id
            except Exception as ex:
                if DEBUG_MODE:
                    logger.debug("Could not apply template: {}".format(ex))

        try:
            view.Scale = scale
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("Could not set scale: {}".format(ex))

        union_bbox = _collect_union_bbox(clashing_ids)
        if union_bbox is not None:
            padded = _expand_bbox(union_bbox, crop_offset)
            try:
                view.SetSectionBox(padded)
            except Exception as ex:
                logger.debug("SetSectionBox failed for {}: {}".format(base_name, ex))
            try:
                view.IsSectionBoxActive = True
            except Exception:
                pass
        else:
            logger.debug("No usable bbox for {} x {}".format(name_a, name_b))

        try:
            view.IsolateElementsTemporary(clashing_ids)
        except Exception as ex:
            logger.debug("IsolateElementsTemporary failed for {}: {}".format(base_name, ex))

        return view

    # -------------------------------------------------------------------------
    # Sheet + grid
    # -------------------------------------------------------------------------

    def _get_viewport_size(self, view):
        """Approximate viewport width/height on sheet (feet) from section box / scale."""
        try:
            sb = view.GetSectionBox()
            if sb is not None:
                mx = abs(sb.Max.X - sb.Min.X)
                my = abs(sb.Max.Y - sb.Min.Y)
                mz = abs(sb.Max.Z - sb.Min.Z)
                sc = max(view.Scale, 1)
                plan = max(mx, my)
                width = plan / float(sc)
                height = max(plan, mz) / float(sc)
                if width > 0.0 and height > 0.0:
                    return (width, height)
        except Exception:
            pass
        try:
            ol = view.Outline
            if ol:
                width = abs(ol.Max.U - ol.Min.U)
                height = abs(ol.Max.V - ol.Min.V)
                if width > 0.0 and height > 0.0:
                    return (width, height)
        except Exception:
            pass
        return (0.3, 0.3)

    def _create_sheet_and_place_viewports(self, sheet_prefix, sheet_name, grid_cols):
        """Create a sheet and lay every pair-view out left-to-right, top-to-bottom."""

        if not self.pair_views:
            return

        max_width = 0.0
        max_height = 0.0
        view_sizes = OrderedDict()
        for pair_key, view in self.pair_views.items():
            w, h = self._get_viewport_size(view)
            view_sizes[pair_key] = (w, h)
            if w > max_width:
                max_width = w
            if h > max_height:
                max_height = h

        if max_width < 0.01:
            max_width = 0.3
        if max_height < 0.01:
            max_height = 0.3

        cell_padding = 0.08
        cell_width = max_width + self.VIEWPORT_SPACING + cell_padding
        cell_height = max_height + self.VIEWPORT_SPACING + cell_padding
        sheet_margin = 0.2

        total = len(self.pair_views)
        cols = min(grid_cols, total)
        rows = (total + cols - 1) // cols

        content_width = cols * cell_width
        content_height = rows * cell_height
        sheet_width = max(content_width + 2 * sheet_margin, 1.5)
        sheet_height = max(content_height + 2 * sheet_margin, 1.0)

        start_x = sheet_margin + cell_width / 2.0
        start_y = sheet_height - sheet_margin - cell_height / 2.0

        sheet = ViewSheet.Create(doc, ElementId.InvalidElementId)
        try:
            sheet.SheetNumber = "{}001".format(sheet_prefix)
        except Exception:
            for i in range(2, 100):
                try:
                    sheet.SheetNumber = "{}{:03d}".format(sheet_prefix, i)
                    break
                except Exception:
                    continue
        try:
            sheet.Name = sheet_name
        except Exception:
            pass
        self.created_sheets.append(sheet)

        col = 0
        row = 0
        for pair_key, view in self.pair_views.items():
            if col >= cols:
                col = 0
                row += 1
            x = start_x + col * cell_width
            y = start_y - row * cell_height
            try:
                Viewport.Create(doc, sheet.Id, view.Id, XYZ(x, y, 0))
            except Exception as ex:
                logger.warning("Could not place viewport for {}: {}".format(pair_key, ex))
            col += 1


# =============================================================================
# Main Entry Point
# =============================================================================

if __name__ == '__main__':
    vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    three_d_types = [vft for vft in vfts
                     if vft.ViewFamily == ViewFamily.ThreeDimensional]

    if not three_d_types:
        forms.alert("No 3D view types found in the project.", title="Error")
    else:
        categories = _discover_clashable_categories()
        if not categories:
            forms.alert(
                "No model categories with elements found in the project.",
                title="No Categories",
            )
        else:
            window = ClashViewsWindow()
            window.ShowDialog()
