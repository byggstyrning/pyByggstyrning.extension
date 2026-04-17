# -*- coding: utf-8 -*-
"""Create section elevation views for family instances and place them on sheets in a grid layout."""

__title__ = "Family\nElevations"
__author__ = "pyByggstyrning"
__highlight__ = "new"
__doc__ = """Create elevation views for family instances and place them on sheets.

Select categories and instances. Annotation categories (e.g. title blocks, tags) are not listed.
Views use instance bounding box and facing orientation; sheets use a grid layout (20 wide).
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
    ViewFamilyType, ViewFamily, View, ViewSheet, Viewport, ViewSection,
    XYZ, ElementId, Transform, BoundingBoxXYZ,
    UnitUtils, Transaction, TransactionGroup,
    FamilyInstance, LocationPoint, LocationCurve
)

# Import pyRevit libraries
import sys
import os.path as op
from collections import OrderedDict
from pyrevit import revit, DB
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

# Get logger
logger = script.get_logger()

# Performance: Debug mode toggle - set True for verbose logging during development
DEBUG_MODE = False

# Get Revit document, app and UIDocument
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# =============================================================================
# Helpers
# =============================================================================

def _get_bbox_crop_dimensions(element):
    """Return (plan_width, height) in feet from axis-aligned bbox. Plan width = max(X,Y extent)."""
    bbox = element.get_BoundingBox(None)
    if not bbox:
        return 3.0, 7.0
    x_ext = abs(bbox.Max.X - bbox.Min.X)
    y_ext = abs(bbox.Max.Y - bbox.Min.Y)
    z_ext = abs(bbox.Max.Z - bbox.Min.Z)
    plan_width = max(x_ext, y_ext)
    height = z_ext
    if plan_width < 0.01:
        plan_width = 3.0
    if height < 0.01:
        height = 7.0
    return plan_width, height


def _get_bbox_center(element):
    """World-space center of axis-aligned bounding box."""
    bbox = element.get_BoundingBox(None)
    if not bbox:
        return XYZ(0, 0, 0)
    return XYZ(
        (bbox.Min.X + bbox.Max.X) / 2.0,
        (bbox.Min.Y + bbox.Max.Y) / 2.0,
        (bbox.Min.Z + bbox.Max.Z) / 2.0
    )


def _get_family_instance_location_point(instance):
    """Location point or curve midpoint for a FamilyInstance; None if unavailable."""
    try:
        loc = instance.Location
        if loc is None:
            return None
        if isinstance(loc, LocationPoint):
            return loc.Point
        if isinstance(loc, LocationCurve):
            crv = loc.Curve
            if crv:
                return crv.Evaluate(0.5, True)
    except Exception as ex:
        logger.debug("FamilyInstance location error: {}".format(ex))
    return None


def _is_annotation_category(cat):
    """True if category is annotation-type (title blocks, tags, generic annotations, etc.)."""
    if not cat:
        return False
    try:
        return cat.CategoryType == DB.CategoryType.Annotation
    except Exception:
        return False


def _discover_family_instance_category_names():
    """Sorted unique category names for model FamilyInstances (excludes annotation categories)."""
    names = set()
    for fi in FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements():
        try:
            cat = fi.Category
            if not cat or not cat.Name:
                continue
            if _is_annotation_category(cat):
                continue
            names.add(cat.Name)
        except Exception:
            pass
    return sorted(names)


# =============================================================================
# Data Classes
# =============================================================================

class CategoryFilterItem(forms.Reactive):
    """One category row in the category filter list."""

    def __init__(self, category_name, is_selected=True):
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


class FamilyInstanceData(forms.Reactive):
    """Class for family instance data binding with WPF UI."""

    def __init__(self, instance):
        """Initialize with a Revit FamilyInstance."""
        super(FamilyInstanceData, self).__init__()
        self.instance = instance
        self.instance_id = instance.Id

        self._mark = self._get_param_string(instance, BuiltInParameter.ALL_MODEL_MARK) or ""
        try:
            sym = instance.Symbol
            self._family_name = sym.Family.Name if sym and sym.Family else ""
            self._type_name = instance.Name if instance.Name else ""
        except Exception:
            self._family_name = ""
            self._type_name = ""

        try:
            cat = instance.Category
            self._category_name = cat.Name if cat else ""
        except Exception:
            self._category_name = ""

        level = doc.GetElement(instance.LevelId) if instance.LevelId else None
        self._level_name = level.Name if level else ""

        self._comments = self._get_param_string(instance, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS) or ""

    def _get_param_string(self, element, param):
        """Get parameter value as string."""
        try:
            p = element.get_Parameter(param)
            if p and p.HasValue:
                return p.AsString() or ""
        except Exception:
            pass
        return ""

    @property
    def Mark(self):
        return self._mark

    @property
    def CategoryName(self):
        return self._category_name

    @property
    def FamilyName(self):
        return self._family_name

    @property
    def TypeName(self):
        return self._type_name

    @property
    def LevelName(self):
        return self._level_name

    @property
    def Comments(self):
        return self._comments

    def get_group_value(self, group_param):
        """Get value for grouping based on parameter name."""
        if group_param == "Level":
            return self._level_name
        elif group_param == "Category":
            return self._category_name
        elif group_param == "Type":
            return self._type_name
        elif group_param == "Family":
            return self._family_name
        elif group_param == "Mark":
            return self._mark
        elif group_param == "Comments":
            return self._comments
        return ""

    def get_sort_value(self, sort_param):
        """Get value for sorting based on parameter name."""
        return self.get_group_value(sort_param)

    def matches_search(self, search_text):
        """Check if instance matches search text."""
        if not search_text:
            return True
        search_lower = search_text.lower()
        return (search_lower in self._mark.lower() or
                search_lower in self._family_name.lower() or
                search_lower in self._type_name.lower() or
                search_lower in self._level_name.lower() or
                search_lower in self._comments.lower() or
                search_lower in self._category_name.lower())


class TemplateItem:
    """Wrapper for view template."""

    def __init__(self, view=None):
        self.view = view
        self.Name = view.Name if view else "<None>"
        self.Id = view.Id if view else None

    def __str__(self):
        return self.Name

    def __repr__(self):
        return self.Name


class SectionTypeItem:
    """Wrapper for section view family type."""

    def __init__(self, vft):
        self.vft = vft
        self.Name = vft.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() if vft else ""
        self.Id = vft.Id if vft else None


# =============================================================================
# Main Window Class
# =============================================================================

class FamilyElevationsWindow(forms.WPFWindow):
    """WPF window for creating family elevations."""

    GRID_COLUMNS = 20
    GROUP_GAP_ROWS = 1
    VIEWPORT_SPACING = 0.03

    def __init__(self):
        logger.debug("Initializing Family Elevations window")

        xaml_file = op.join(pushbutton_dir, "FamilyElevationsWindow.xaml")
        forms.WPFWindow.__init__(self, xaml_file)

        try:
            from styles import load_styles_to_window
            load_styles_to_window(self)
        except Exception as ex:
            logger.debug("Could not load styles: {}".format(ex))

        self.all_instances_data = []
        self.instances_data = ObservableCollection[FamilyInstanceData]()
        self.elevation_views = {}
        self.created_sheets = []

        self._category_items = ObservableCollection[CategoryFilterItem]()
        self._category_display = ObservableCollection[CategoryFilterItem]()

        self._setup_scale_options()
        self._setup_templates()
        self._setup_elevation_types()
        self._setup_grouping_options()
        self._populate_category_filters()
        self._sync_category_list_display()
        self.categoryListBox.ItemsSource = self._category_display
        self._load_instances()

        self.instancesDataGrid.ItemsSource = self.instances_data
        self.instancesDataGrid.SelectionChanged += self.instancesDataGrid_SelectionChanged
        try:
            self.instancesDataGrid.UnselectAll()
        except Exception:
            pass

        self._update_category_selection_count()
        self._update_selection_count()

    def _setup_scale_options(self):
        scales = ["1:10", "1:20", "1:50", "1:100"]
        for scale in scales:
            self.scaleComboBox.Items.Add(scale)
        self.scaleComboBox.SelectedIndex = 2

    def _setup_templates(self):
        self.templateComboBox.Items.Add(TemplateItem(None))
        all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
        templates = [v for v in all_views if v.IsTemplate and v.ViewType == DB.ViewType.Section]
        for template in sorted(templates, key=lambda x: x.Name):
            self.templateComboBox.Items.Add(TemplateItem(template))
        self.templateComboBox.SelectedIndex = 0

    def _setup_elevation_types(self):
        vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
        section_types = []
        for vft in vfts:
            try:
                if vft.ViewFamily == ViewFamily.Section:
                    section_types.append(vft)
            except Exception as ex:
                logger.debug("Error checking ViewFamilyType: {}".format(ex))

        for vft in section_types:
            self.elevationTypeComboBox.Items.Add(SectionTypeItem(vft))

        if section_types:
            self.elevationTypeComboBox.SelectedIndex = 0
        else:
            logger.warning("No section view family types found!")

    def _setup_grouping_options(self):
        options = ["Level", "Category", "Type", "Family", "Mark", "Comments"]
        for opt in options:
            self.groupByComboBox.Items.Add(opt)
            self.sortByComboBox.Items.Add(opt)
        self.groupByComboBox.SelectedIndex = 0
        self.sortByComboBox.SelectedIndex = 4

    def _populate_category_filters(self):
        self._category_items.Clear()
        for name in _discover_family_instance_category_names():
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

    def _update_category_selection_count(self):
        n_sel = sum(1 for i in self._category_items if i.IsSelected)
        n_tot = self._category_items.Count
        if hasattr(self, "categorySelectionCountText") and self.categorySelectionCountText:
            self.categorySelectionCountText.Text = "{} of {} categories included".format(n_sel, n_tot)

    def _get_selected_category_names(self):
        selected = []
        for item in self._category_items:
            if item.IsSelected:
                selected.append(item.CategoryName)
        return selected

    def _load_instances(self):
        """Load FamilyInstance rows for checked categories."""
        self.all_instances_data = []
        allowed = set(self._get_selected_category_names())

        for fi in FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements():
            try:
                cat = fi.Category
                if not cat or not cat.Name:
                    continue
                if _is_annotation_category(cat):
                    continue
                if cat.Name not in allowed:
                    continue
                row = FamilyInstanceData(fi)
                self.all_instances_data.append(row)
            except Exception as ex:
                logger.debug("Skipping instance {}: {}".format(fi.Id, ex))

        self.all_instances_data.sort(
            key=lambda r: (r.CategoryName, r.FamilyName, r.Mark, r.TypeName)
        )

        self._apply_search_filter()
        self._update_category_selection_count()

        logger.debug("Loaded {} instances for selected categories".format(len(self.all_instances_data)))

    def _apply_search_filter(self):
        search_text = ""
        if hasattr(self, "instanceSearchTextBox") and self.instanceSearchTextBox:
            search_text = self.instanceSearchTextBox.Text or ""

        self.instances_data.Clear()
        for item in self.all_instances_data:
            if item.matches_search(search_text):
                self.instances_data.Add(item)
        try:
            if hasattr(self, "instancesDataGrid") and self.instancesDataGrid:
                self.instancesDataGrid.UnselectAll()
        except Exception:
            pass
        self._update_selection_count()

    def _update_selection_count(self):
        sel = 0
        try:
            if hasattr(self, "instancesDataGrid") and self.instancesDataGrid:
                sel = self.instancesDataGrid.SelectedItems.Count
        except Exception:
            sel = 0
        filtered_count = self.instances_data.Count
        total_count = len(self.all_instances_data)

        if filtered_count < total_count:
            self.selectionCountText.Text = "{} selected ({} shown of {})".format(
                sel, filtered_count, total_count)
        else:
            self.selectionCountText.Text = "{} of {} instances selected".format(
                sel, total_count)

    def _get_scale_value(self):
        scale_text = self.scaleComboBox.SelectedItem
        if scale_text:
            return int(scale_text.split(":")[1])
        return 50

    def _get_selected_instances(self):
        """Rows chosen in the DataGrid (Extended selection: Ctrl/Shift multi-select)."""
        out = []
        try:
            if hasattr(self, "instancesDataGrid") and self.instancesDataGrid:
                for obj in self.instancesDataGrid.SelectedItems:
                    out.append(obj)
        except Exception:
            pass
        return out

    def _get_plan_view(self):
        active = doc.ActiveView
        if active.ViewType in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.AreaPlan]:
            return active

        views = FilteredElementCollector(doc).OfClass(DB.ViewPlan).ToElements()
        for view in views:
            if not view.IsTemplate and view.ViewType == DB.ViewType.FloorPlan:
                return view
        return None

    # =========================================================================
    # Event Handlers
    # =========================================================================

    def categorySearchTextBox_TextChanged(self, sender, args):
        """Filter which categories appear in the list (does not change inclusion)."""
        self._sync_category_list_display()

    def categorySelectAllButton_Click(self, sender, args):
        for item in self._category_items:
            item.IsSelected = True
        self._load_instances()

    def categoryDeselectAllButton_Click(self, sender, args):
        for item in self._category_items:
            item.IsSelected = False
        self._load_instances()

    def categoryFilterCheckBox_Changed(self, sender, args):
        """Reload instances when category inclusion changes."""
        self._load_instances()

    def instanceSearchTextBox_TextChanged(self, sender, args):
        self._apply_search_filter()

    def instancesDataGrid_SelectionChanged(self, sender, args):
        self._update_selection_count()

    def instanceSelectAllFilteredButton_Click(self, sender, args):
        try:
            self.instancesDataGrid.SelectAll()
        except Exception:
            pass
        self._update_selection_count()

    def instanceDeselectAllFilteredButton_Click(self, sender, args):
        try:
            self.instancesDataGrid.UnselectAll()
        except Exception:
            pass
        self._update_selection_count()

    def cancelButton_Click(self, sender, args):
        self.Close()

    def createButton_Click(self, sender, args):
        selected = self._get_selected_instances()

        if not selected:
            forms.alert("No instances selected.", title="Warning")
            return

        if self.elevationTypeComboBox.SelectedItem is None:
            forms.alert("No section type available in the project.", title="Error")
            return

        plan_view = self._get_plan_view()
        if not plan_view:
            forms.alert("No floor plan view found for creating elevations.", title="Error")
            return

        scale = self._get_scale_value()
        elev_type_id = self.elevationTypeComboBox.SelectedItem.Id
        template_item = self.templateComboBox.SelectedItem
        template_id = template_item.Id if template_item and template_item.Id else None
        name_prefix = self.namePrefixTextBox.Text or "Elev - "
        crop_offset_mm = float(self.cropOffsetTextBox.Text or "300")
        sheet_prefix = self.sheetPrefixTextBox.Text or "FE-"
        group_by = self.groupByComboBox.SelectedItem or "Level"
        sort_by = self.sortByComboBox.SelectedItem or "Mark"

        self.Close()

        try:
            title = "Creating Family Elevations ({} instances)".format(len(selected))
            with forms.ProgressBar(title=title) as pb:
                self._create_family_elevations(
                    selected, scale, elev_type_id, template_id,
                    name_prefix, crop_offset_mm,
                    sheet_prefix, group_by, sort_by, pb
                )

            if self.created_sheets:
                uidoc.ActiveView = self.created_sheets[0]

        except Exception as ex:
            logger.error("Error creating elevations: {}".format(ex))
            import traceback
            logger.debug(traceback.format_exc())
            forms.alert("Error creating elevations: {}".format(ex), title="Error")

    # =========================================================================
    # Elevation Creation
    # =========================================================================

    def _create_family_elevations(self, selected_items, scale, elev_type_id, template_id,
                                  name_prefix, crop_offset_mm,
                                  sheet_prefix, group_by, sort_by, progress_bar=None):
        """Create section views for selected instances and place on sheets."""

        try:
            crop_offset = UnitUtils.ConvertToInternalUnits(crop_offset_mm, DB.UnitTypeId.Millimeters)
        except Exception:
            crop_offset = crop_offset_mm / 304.8

        grouped = self._group_instances(selected_items, group_by)
        for group_name in grouped:
            grouped[group_name] = self._sort_instances(grouped[group_name], sort_by)

        total = len(selected_items)
        processed = 0

        with TransactionGroup(doc, "Create Family Elevations") as tg:
            tg.Start()

            with Transaction(doc, "Create Elevation Views") as t:
                t.Start()

                for group_name, rows in grouped.items():
                    for row in rows:
                        processed += 1
                        if progress_bar:
                            progress_bar.update_progress(processed, total)

                        try:
                            view = self._create_single_elevation(
                                row.instance, elev_type_id, scale,
                                name_prefix, template_id, crop_offset
                            )
                            if view:
                                self.elevation_views[row.instance_id] = view
                        except Exception as ex:
                            logger.warning("Failed to create elevation for {}: {}".format(
                                row.instance_id, ex))

                t.Commit()

            with Transaction(doc, "Create Sheets and Place Views") as t:
                t.Start()
                self._create_sheets_and_place_viewports(grouped, sheet_prefix, scale)
                t.Commit()

            tg.Assimilate()

    def _create_single_elevation(self, instance, section_type_id, scale,
                                 name_prefix, template_id, crop_offset):
        """Create a section view for a family instance using ViewSection.CreateSection."""

        mark_param = instance.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        mark = mark_param.AsString() if mark_param and mark_param.HasValue else str(get_element_id_value(instance.Id))

        if DEBUG_MODE:
            logger.debug("Creating section for instance: {} (Id={})".format(mark, get_element_id_value(instance.Id)))

        inst_loc = _get_family_instance_location_point(instance)
        if inst_loc is None:
            inst_loc = _get_bbox_center(instance)

        door_width, door_height = _get_bbox_crop_dimensions(instance)

        if DEBUG_MODE:
            logger.debug("  Crop plan width={:.2f} ft, height={:.2f} ft".format(door_width, door_height))

        facing = XYZ(0, 1, 0)
        try:
            if isinstance(instance, FamilyInstance):
                fo = instance.FacingOrientation
                facing = XYZ(fo.X, fo.Y, fo.Z)
        except Exception as ex:
            logger.warning("  Could not get facing: {}, using +Y".format(ex))

        view_direction = XYZ(facing.X, facing.Y, 0)
        if view_direction.GetLength() < 1e-9:
            view_direction = XYZ(0, 1, 0)
        else:
            view_direction = view_direction.Normalize()

        up_direction = XYZ(0, 0, 1)
        right_direction = up_direction.CrossProduct(view_direction)

        view_offset = 3.0
        section_origin = XYZ(
            inst_loc.X - view_direction.X * view_offset,
            inst_loc.Y - view_direction.Y * view_offset,
            inst_loc.Z
        )

        half_width = door_width / 2.0 + crop_offset
        bottom = -crop_offset
        top = door_height + crop_offset
        near_clip = 0.0
        far_clip = view_offset + 2.0

        section_box = BoundingBoxXYZ()
        section_box.Min = XYZ(-half_width, bottom, near_clip)
        section_box.Max = XYZ(half_width, top, far_clip)

        transform = Transform.Identity
        transform.Origin = section_origin
        transform.BasisX = right_direction
        transform.BasisY = up_direction
        transform.BasisZ = view_direction
        section_box.Transform = transform

        section_view = ViewSection.CreateSection(doc, section_type_id, section_box)
        if not section_view:
            logger.error("  Failed to create section view!")
            return None

        section_view.Scale = scale

        view_name = "{}{}".format(name_prefix, mark)
        try:
            section_view.Name = view_name
        except Exception:
            for i in range(1, 100):
                try:
                    section_view.Name = "{} ({})".format(view_name, i)
                    break
                except Exception:
                    continue

        if template_id:
            try:
                section_view.ViewTemplateId = template_id
            except Exception:
                pass

        self._set_view_phase_from_element(section_view, instance)

        section_view.CropBoxActive = True
        section_view.CropBoxVisible = False

        try:
            ann_crop_param = section_view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
            if ann_crop_param and not ann_crop_param.IsReadOnly:
                ann_crop_param.Set(1)
        except Exception:
            pass

        try:
            eids = List[ElementId]()
            eids.Add(instance.Id)
            section_view.IsolateElementsTemporary(eids)
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("  Could not isolate element: {}".format(ex))

        return section_view

    def _set_view_phase_from_element(self, view, element):
        """Set the view's phase to match the element's phase created."""
        try:
            phase_param = element.get_Parameter(BuiltInParameter.PHASE_CREATED)
            if phase_param and phase_param.HasValue:
                phase_id = phase_param.AsElementId()
                view_phase_param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
                if view_phase_param and not view_phase_param.IsReadOnly:
                    view_phase_param.Set(phase_id)
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("Could not set view phase: {}".format(ex))

    def _group_instances(self, rows, group_param):
        groups = OrderedDict()
        for row in rows:
            value = row.get_group_value(group_param) or "Ungrouped"
            if value not in groups:
                groups[value] = []
            groups[value].append(row)
        return groups

    def _sort_instances(self, rows, sort_param):
        return sorted(rows, key=lambda r: r.get_sort_value(sort_param) or "")

    def _get_viewport_size(self, view):
        try:
            if view.CropBoxActive:
                crop_box = view.CropBox
                model_width = abs(crop_box.Max.X - crop_box.Min.X)
                model_height = abs(crop_box.Max.Y - crop_box.Min.Y)
                sc = view.Scale
                return (model_width / sc, model_height / sc)
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("Could not get viewport size: {}".format(ex))
        return (0.1, 0.15)

    def _calculate_grid_layout(self, grouped_items, cell_width, cell_height, sheet_margin):
        total_viewports = len(self.elevation_views)
        cols = min(self.GRID_COLUMNS, total_viewports)
        content_width = cols * cell_width

        row_count = 0
        col = 0
        for group_name, rows in grouped_items.items():
            for row in rows:
                if row.instance_id not in self.elevation_views:
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

    def _calculate_viewport_positions(self, grouped_items, cols, cell_width, cell_height, start_x, start_y):
        positions = []
        col = 0
        row = 0
        for group_name, rows in grouped_items.items():
            for row_data in rows:
                if row_data.instance_id not in self.elevation_views:
                    continue
                if col >= cols:
                    col = 0
                    row += 1
                x = start_x + col * cell_width
                y = start_y - row * cell_height
                positions.append((row_data.instance_id, x, y))
                col += 1
            if col > 0:
                col = 0
                row += 1 + self.GROUP_GAP_ROWS
        return positions

    def _create_sheets_and_place_viewports(self, grouped_items, sheet_prefix, scale):
        max_width = 0
        max_height = 0
        for inst_id, view in self.elevation_views.items():
            w, h = self._get_viewport_size(view)
            max_width = max(max_width, w)
            max_height = max(max_height, h)

        if max_width < 0.01:
            max_width = 0.15
        if max_height < 0.01:
            max_height = 0.2

        cell_padding = 0.08
        cell_width = max_width + self.VIEWPORT_SPACING + cell_padding
        cell_height = max_height + self.VIEWPORT_SPACING + cell_padding
        sheet_margin = 0.15

        cols, row_count, sheet_width, sheet_height, start_x, start_y = self._calculate_grid_layout(
            grouped_items, cell_width, cell_height, sheet_margin)

        positions = self._calculate_viewport_positions(
            grouped_items, cols, cell_width, cell_height, start_x, start_y)

        if not positions:
            logger.warning("No positions calculated for viewports")
            return

        sheet = ViewSheet.Create(doc, ElementId.InvalidElementId)
        sheet.SheetNumber = "{}001".format(sheet_prefix)
        sheet.Name = "Family Elevations"
        self.created_sheets.append(sheet)

        for inst_id, x, y in positions:
            elev_view = self.elevation_views.get(inst_id)
            if not elev_view:
                continue
            try:
                Viewport.Create(doc, sheet.Id, elev_view.Id, XYZ(x, y, 0))
            except Exception as ex:
                logger.warning("Could not place viewport for instance {}: {}".format(inst_id, ex))


# =============================================================================
# Main Entry Point
# =============================================================================

if __name__ == '__main__':
    instances = FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements()
    if not instances:
        forms.alert("No family instances found in the project.", title="No Instances")
    else:
        vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
        section_types = [vft for vft in vfts if vft.ViewFamily == ViewFamily.Section]

        if DEBUG_MODE:
            for vft in vfts:
                try:
                    name = vft.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                    logger.debug("ViewFamilyType: {} - ViewFamily: {}".format(name, vft.ViewFamily))
                except Exception:
                    pass

        if not section_types:
            forms.alert("No section view types found in the project.", title="Error")
        else:
            model_categories = _discover_family_instance_category_names()
            if not model_categories:
                forms.alert(
                    "No model family instance categories found. Annotation categories "
                    "(e.g. title blocks, tags, generic annotations) are excluded.",
                    title="No Categories",
                )
            else:
                window = FamilyElevationsWindow()
                window.ShowDialog()
