# -*- coding: utf-8 -*-
"""Create elevation views for door instances and place them on sheets in a grid layout."""

__title__ = "Door\nElevations"
__author__ = "pyByggstyrning"
__highlight__ = "new"
__doc__ = """Create elevation views for door instances and place them on sheets.

Elevations are oriented to face the door opening direction.
Views are placed on sheets in a grid layout (20 wide), grouped and sorted by user-selected parameters.
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
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    ViewFamilyType, ViewFamily, View, ViewSheet, Viewport, ViewSection,
    ElementTransformUtils, Line, XYZ, ElementId, Transform, BoundingBoxXYZ,
    UnitUtils, Transaction, TransactionGroup, CurveLoop
)

# Import pyRevit libraries
import os
import sys
import os.path as op
import math
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

# Get logger
logger = script.get_logger()

# Performance: Debug mode toggle - set True for verbose logging during development
DEBUG_MODE = False

# Get Revit document, app and UIDocument
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument


# =============================================================================
# Data Classes
# =============================================================================

class DoorItemData(forms.Reactive):
    """Class for door data binding with WPF UI."""
    
    def __init__(self, door, phase):
        """Initialize with a Revit door instance."""
        super(DoorItemData, self).__init__()
        self.door = door
        self.door_id = door.Id
        self._is_selected = True
        self.phase = phase
        
        # Get door properties
        self._mark = self._get_param_string(door, BuiltInParameter.ALL_MODEL_MARK) or ""
        self._family_name = door.Symbol.Family.Name if door.Symbol and door.Symbol.Family else ""
        self._type_name = door.Name if door.Name else ""
        
        # Level
        level = doc.GetElement(door.LevelId) if door.LevelId else None
        self._level_name = level.Name if level else ""
        
        # From Room
        self._from_room = self._get_from_room_name(door, phase)
        
        # Comments
        self._comments = self._get_param_string(door, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS) or ""
    
    def _get_param_string(self, element, param):
        """Get parameter value as string."""
        try:
            p = element.get_Parameter(param)
            if p and p.HasValue:
                return p.AsString() or ""
        except:
            pass
        return ""
    
    def _get_from_room_name(self, door, phase):
        """Get the From Room name/number for the door."""
        try:
            from_room = door.FromRoom[phase] if phase else None
            if from_room:
                room_number = from_room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                room_name = from_room.get_Parameter(BuiltInParameter.ROOM_NAME)
                number = room_number.AsString() if room_number and room_number.HasValue else ""
                name = room_name.AsString() if room_name and room_name.HasValue else ""
                if number and name:
                    return "{} - {}".format(number, name)
                return number or name or ""
        except Exception as ex:
            logger.debug("Error getting FromRoom: {}".format(ex))
        return ""
    
    @property
    def IsSelected(self):
        return self._is_selected
    
    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        self.OnPropertyChanged("IsSelected")
    
    @property
    def Mark(self):
        return self._mark
    
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
    def FromRoom(self):
        return self._from_room
    
    @property
    def Comments(self):
        return self._comments
    
    def get_group_value(self, group_param):
        """Get value for grouping based on parameter name."""
        if group_param == "Level":
            return self._level_name
        elif group_param == "From Room":
            return self._from_room
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
        """Check if door matches search text."""
        if not search_text:
            return True
        search_lower = search_text.lower()
        # Search across all visible properties
        return (search_lower in self._mark.lower() or
                search_lower in self._family_name.lower() or
                search_lower in self._type_name.lower() or
                search_lower in self._level_name.lower() or
                search_lower in self._from_room.lower() or
                search_lower in self._comments.lower())


class TemplateItem:
    """Wrapper for view template."""
    def __init__(self, view=None):
        self.view = view
        self.Name = view.Name if view else "<None>"
        self.Id = view.Id if view else None
    
    def __str__(self):
        """String representation for ComboBox display."""
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

class DoorElevationsWindow(forms.WPFWindow):
    """WPF window for creating door elevations."""
    
    # Grid layout constants
    GRID_COLUMNS = 20
    GROUP_GAP_ROWS = 1  # Single row gap between groups
    VIEWPORT_SPACING = 0.03  # feet between viewports (~10mm)
    
    def __init__(self):
        """Initialize the window."""
        logger.debug("Initializing Door Elevations window")
        
        xaml_file = op.join(pushbutton_dir, "DoorElevationsWindow.xaml")
        forms.WPFWindow.__init__(self, xaml_file)
        
        # Load styles AFTER window initialization
        try:
            from styles import load_styles_to_window
            load_styles_to_window(self)
        except Exception as ex:
            logger.debug("Could not load styles: {}".format(ex))
        
        # Get active phase
        self.active_phase = self._get_active_phase()
        
        # Initialize data
        self.all_doors_data = []  # Master list of all doors
        self.doors_data = ObservableCollection[DoorItemData]()  # Filtered list for display
        self.elevation_views = {}  # Maps door.Id to created view
        self.created_sheets = []
        
        # Set up UI
        self._setup_scale_options()
        self._setup_templates()
        self._setup_elevation_types()
        self._setup_grouping_options()
        self._load_doors()
        
        # Bind data
        self.doorsDataGrid.ItemsSource = self.doors_data
        self._update_selection_count()
    
    def _get_active_phase(self):
        """Get the active phase from the active view."""
        try:
            active_view = doc.ActiveView
            phase_param = active_view.get_Parameter(BuiltInParameter.VIEW_PHASE)
            if phase_param and phase_param.HasValue:
                phase_id = phase_param.AsElementId()
                return doc.GetElement(phase_id)
        except:
            pass
        
        # Fallback to last phase in project
        try:
            phases = FilteredElementCollector(doc).OfClass(DB.Phase).ToElements()
            if phases:
                return phases[-1]  # Last phase
        except:
            pass
        return None
    
    def _setup_scale_options(self):
        """Set up view scale options."""
        scales = ["1:10", "1:20", "1:50", "1:100"]
        for scale in scales:
            self.scaleComboBox.Items.Add(scale)
        self.scaleComboBox.SelectedIndex = 2  # Default 1:50
    
    def _setup_templates(self):
        """Set up view template options."""
        # Add None option first
        self.templateComboBox.Items.Add(TemplateItem(None))
        
        # Find section templates (since we're using sections now)
        all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
        templates = [v for v in all_views if v.IsTemplate and v.ViewType == DB.ViewType.Section]
        
        for template in sorted(templates, key=lambda x: x.Name):
            self.templateComboBox.Items.Add(TemplateItem(template))
        
        self.templateComboBox.SelectedIndex = 0  # Default None
    
    def _setup_elevation_types(self):
        """Set up section view family type options (using sections for door views)."""
        vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
        
        # Filter for Section view family types
        section_types = []
        for vft in vfts:
            try:
                if vft.ViewFamily == ViewFamily.Section:
                    section_types.append(vft)
                    logger.debug("Found section type: {} (ViewFamily={})".format(
                        vft.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString(),
                        vft.ViewFamily))
            except Exception as ex:
                logger.debug("Error checking ViewFamilyType: {}".format(ex))
        
        for vft in section_types:
            self.elevationTypeComboBox.Items.Add(SectionTypeItem(vft))
        
        if section_types:
            self.elevationTypeComboBox.SelectedIndex = 0
        else:
            logger.warning("No section view family types found!")
    
    def _setup_grouping_options(self):
        """Set up grouping and sorting options."""
        options = ["Level", "From Room", "Type", "Family", "Mark", "Comments"]
        
        for opt in options:
            self.groupByComboBox.Items.Add(opt)
            self.sortByComboBox.Items.Add(opt)
        
        self.groupByComboBox.SelectedIndex = 0  # Default Level
        self.sortByComboBox.SelectedIndex = 4   # Default Mark
    
    def _load_doors(self):
        """Load all doors from the project."""
        doors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors)\
            .WhereElementIsNotElementType().ToElements()
        
        for door in doors:
            # Skip doors without host (curtain wall doors etc.)
            try:
                if door.Host:
                    door_item = DoorItemData(door, self.active_phase)
                    self.all_doors_data.append(door_item)
                    self.doors_data.Add(door_item)
            except Exception as ex:
                logger.debug("Skipping door {}: {}".format(door.Id, ex))
        
        logger.debug("Loaded {} doors".format(len(self.all_doors_data)))
    
    def _update_selection_count(self):
        """Update the selection count text."""
        # Count selected from ALL doors (not just filtered)
        selected_count = sum(1 for d in self.all_doors_data if d.IsSelected)
        filtered_count = self.doors_data.Count
        total_count = len(self.all_doors_data)
        
        if filtered_count < total_count:
            self.selectionCountText.Text = "{} selected ({} shown of {})".format(
                selected_count, filtered_count, total_count)
        else:
            self.selectionCountText.Text = "{} of {} doors selected".format(
                selected_count, total_count)
    
    def _get_scale_value(self):
        """Get the selected scale as integer."""
        scale_text = self.scaleComboBox.SelectedItem
        if scale_text:
            # Parse "1:50" -> 50
            return int(scale_text.split(":")[1])
        return 50
    
    def _get_selected_doors(self):
        """Get list of selected door data items from ALL doors (not just filtered)."""
        return [d for d in self.all_doors_data if d.IsSelected]
    
    def _get_plan_view(self):
        """Get a plan view for creating elevations."""
        # Try active view first if it's a plan
        active = doc.ActiveView
        if active.ViewType in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.AreaPlan]:
            return active
        
        # Find any floor plan
        views = FilteredElementCollector(doc).OfClass(DB.ViewPlan).ToElements()
        for view in views:
            if not view.IsTemplate and view.ViewType == DB.ViewType.FloorPlan:
                return view
        
        return None
    
    # =========================================================================
    # Event Handlers
    # =========================================================================
    
    def doorSearchTextBox_TextChanged(self, sender, args):
        """Handle door search text changed."""
        search_text = self.doorSearchTextBox.Text or ""
        
        # Clear and repopulate the filtered list
        self.doors_data.Clear()
        for door_item in self.all_doors_data:
            if door_item.matches_search(search_text):
                self.doors_data.Add(door_item)
        
        self._update_selection_count()
    
    def selectAllCheckBox_Checked(self, sender, args):
        """Handle select all checked - applies to filtered doors."""
        for door_data in self.doors_data:
            door_data.IsSelected = True
        self._update_selection_count()
    
    def selectAllCheckBox_Unchecked(self, sender, args):
        """Handle select all unchecked - applies to filtered doors."""
        for door_data in self.doors_data:
            door_data.IsSelected = False
        self._update_selection_count()
    
    def doorCheckBox_Changed(self, sender, args):
        """Handle individual door checkbox change."""
        self._update_selection_count()
    
    def cancelButton_Click(self, sender, args):
        """Handle cancel button click."""
        self.Close()
    
    def createButton_Click(self, sender, args):
        """Handle create button click."""
        selected_doors = self._get_selected_doors()
        
        if not selected_doors:
            forms.alert("No doors selected.", title="Warning")
            return
        
        # Validate settings
        if self.elevationTypeComboBox.SelectedItem is None:
            forms.alert("No elevation type available in the project.", title="Error")
            return
        
        plan_view = self._get_plan_view()
        if not plan_view:
            forms.alert("No floor plan view found for creating elevations.", title="Error")
            return
        
        # Get settings
        scale = self._get_scale_value()
        elev_type_id = self.elevationTypeComboBox.SelectedItem.Id
        template_item = self.templateComboBox.SelectedItem
        template_id = template_item.Id if template_item and template_item.Id else None
        name_prefix = self.namePrefixTextBox.Text or "Door Elev - "
        crop_offset_mm = float(self.cropOffsetTextBox.Text or "300")
        crop_to_door = self.cropToDoorCheckBox.IsChecked
        sheet_prefix = self.sheetPrefixTextBox.Text or "DE-"
        group_by = self.groupByComboBox.SelectedItem or "Level"
        sort_by = self.sortByComboBox.SelectedItem or "Mark"
        
        # Close window before showing progress bar (forms.ProgressBar is modal)
        self.Close()
        
        try:
            # Create elevations and place on sheets with progress bar
            with forms.ProgressBar(title="Creating Door Elevations ({} doors)".format(len(selected_doors))) as pb:
                self._create_door_elevations(
                    selected_doors, scale, elev_type_id, template_id,
                    name_prefix, crop_offset_mm, crop_to_door,
                    sheet_prefix, group_by, sort_by, plan_view, pb
                )
            
            # Open the first created sheet
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
    
    def _create_door_elevations(self, selected_doors, scale, elev_type_id, template_id,
                                 name_prefix, crop_offset_mm, crop_to_door,
                                 sheet_prefix, group_by, sort_by, plan_view, progress_bar=None):
        """Create elevation views for selected doors and place on sheets."""
        
        # Convert offset to internal units (feet)
        try:
            # Revit 2022+
            crop_offset = UnitUtils.ConvertToInternalUnits(crop_offset_mm, DB.UnitTypeId.Millimeters)
        except:
            # Older Revit
            crop_offset = crop_offset_mm / 304.8  # mm to feet
        
        # Performance: Pre-cache door type dimensions (many doors share same type)
        type_dimensions = {}  # type_id.IntegerValue -> (width, height)
        for door_data in selected_doors:
            type_id = door_data.door.GetTypeId()
            if type_id.IntegerValue not in type_dimensions:
                door_type = doc.GetElement(type_id)
                width = 3.0  # default feet
                height = 7.0  # default feet
                if door_type:
                    w_param = door_type.get_Parameter(BuiltInParameter.DOOR_WIDTH)
                    h_param = door_type.get_Parameter(BuiltInParameter.DOOR_HEIGHT)
                    if w_param and w_param.HasValue:
                        width = w_param.AsDouble()
                    if h_param and h_param.HasValue:
                        height = h_param.AsDouble()
                type_dimensions[type_id.IntegerValue] = (width, height)
        
        if DEBUG_MODE:
            logger.debug("Cached {} unique door types".format(len(type_dimensions)))
        
        # Performance: Cache section type info for logging
        section_type = doc.GetElement(elev_type_id)
        section_type_name = section_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() if section_type else "Unknown"
        
        # Group and sort doors
        grouped_doors = self._group_doors(selected_doors, group_by)
        for group_name in grouped_doors:
            grouped_doors[group_name] = self._sort_doors(grouped_doors[group_name], sort_by)
        
        # Count total doors for progress
        total_doors = len(selected_doors)
        processed = 0
        
        with TransactionGroup(doc, "Create Door Elevations") as tg:
            tg.Start()
            
            # Step 1: Create all elevation views
            with Transaction(doc, "Create Elevation Views") as t:
                t.Start()
                
                for group_name, doors in grouped_doors.items():
                    for door_data in doors:
                        processed += 1
                        if progress_bar:
                            progress_bar.update_progress(processed, total_doors)
                        
                        try:
                            view = self._create_single_elevation(
                                door_data.door, elev_type_id, scale, plan_view,
                                name_prefix, template_id, crop_to_door, crop_offset,
                                type_dimensions, section_type_name
                            )
                            if view:
                                self.elevation_views[door_data.door_id] = view
                        except Exception as ex:
                            logger.warning("Failed to create elevation for door {}: {}".format(
                                door_data.door_id, ex))
                
                t.Commit()
            
            # Step 2: Calculate grid positions and create sheets
            with Transaction(doc, "Create Sheets and Place Views") as t:
                t.Start()
                
                self._create_sheets_and_place_viewports(
                    grouped_doors, sheet_prefix, scale
                )
                
                t.Commit()
            
            tg.Assimilate()
    
    def _create_single_elevation(self, door, section_type_id, scale, plan_view,
                                  name_prefix, template_id, crop_to_door, crop_offset,
                                  type_dimensions=None, section_type_name=None):
        """Create a section view for a door using ViewSection.CreateSection.
        
        This gives us full control over the view direction, unlike ElevationMarker.
        
        Args:
            type_dimensions: Pre-cached dict of type_id.IntegerValue -> (width, height)
            section_type_name: Pre-cached section type name for logging
        """
        
        door_mark_param = door.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        door_mark = door_mark_param.AsString() if door_mark_param and door_mark_param.HasValue else str(door.Id.IntegerValue)
        
        if DEBUG_MODE:
            logger.debug("=" * 60)
            logger.debug("Creating section for door: {} (Id={})".format(door_mark, door.Id.IntegerValue))
            logger.debug("  ViewFamilyType: {} (Id={})".format(
                section_type_name or "Unknown", section_type_id.IntegerValue))
        
        # Get door location
        door_location = door.Location.Point
        if DEBUG_MODE:
            logger.debug("  Door location: ({:.2f}, {:.2f}, {:.2f})".format(
                door_location.X, door_location.Y, door_location.Z))
        
        # Get host wall
        host_wall = door.Host
        if not host_wall:
            logger.warning("  Door {} has no host wall - SKIPPING".format(door.Id))
            return None
        if DEBUG_MODE:
            logger.debug("  Host wall: {} (Id={})".format(host_wall.Name, host_wall.Id.IntegerValue))
        
        # Get door dimensions from cache or fallback to defaults
        if type_dimensions:
            door_width, door_height = type_dimensions.get(
                door.GetTypeId().IntegerValue, (3.0, 7.0))
        else:
            # Fallback: fetch dimensions (slower path for backwards compatibility)
            door_height = 7.0
            door_width = 3.0
            try:
                door_type = doc.GetElement(door.GetTypeId())
                height_param = door_type.get_Parameter(BuiltInParameter.DOOR_HEIGHT)
                if height_param and height_param.HasValue:
                    door_height = height_param.AsDouble()
                width_param = door_type.get_Parameter(BuiltInParameter.DOOR_WIDTH)
                if width_param and width_param.HasValue:
                    door_width = width_param.AsDouble()
            except:
                pass
        
        if DEBUG_MODE:
            logger.debug("  Door dimensions: width={:.2f} ft, height={:.2f} ft".format(door_width, door_height))
        
        # Get door facing orientation - this points into the "To Room" (swing direction)
        # We want to view from the FROM room side, looking at the door
        try:
            facing = door.FacingOrientation
            if DEBUG_MODE:
                logger.debug("  Door FacingOrientation: ({:.3f}, {:.3f}, {:.3f})".format(
                    facing.X, facing.Y, facing.Z))
        except Exception as ex:
            logger.warning("  Could not get door facing: {}, using +Y".format(ex))
            facing = XYZ(0, 1, 0)
        
        # View direction = opposite of facing (looking FROM the From Room TOWARDS the door)
        # If facing points into the To Room, we want to look in that direction (towards To Room)
        # Actually, let's look at the door from the From Room side, which means view_direction = facing
        view_direction = XYZ(facing.X, facing.Y, 0).Normalize()
        
        # Up direction is always Z
        up_direction = XYZ(0, 0, 1)
        
        # Right direction = up cross view_direction
        right_direction = up_direction.CrossProduct(view_direction)
        
        if DEBUG_MODE:
            logger.debug("  View direction: ({:.3f}, {:.3f}, {:.3f})".format(
                view_direction.X, view_direction.Y, view_direction.Z))
            logger.debug("  Up direction: ({:.3f}, {:.3f}, {:.3f})".format(
                up_direction.X, up_direction.Y, up_direction.Z))
            logger.debug("  Right direction: ({:.3f}, {:.3f}, {:.3f})".format(
                right_direction.X, right_direction.Y, right_direction.Z))
        
        # Section origin - place it in front of the door (on the From Room side)
        # Offset from door location in the opposite of view_direction
        view_offset = 3.0  # feet - distance from door to place section origin
        section_origin = XYZ(
            door_location.X - view_direction.X * view_offset,
            door_location.Y - view_direction.Y * view_offset,
            door_location.Z
        )
        if DEBUG_MODE:
            logger.debug("  Section origin: ({:.2f}, {:.2f}, {:.2f})".format(
                section_origin.X, section_origin.Y, section_origin.Z))
        
        # Create the section bounding box (in view coordinates)
        # The BoundingBoxXYZ for a section defines the crop region
        # Min/Max are in the view's local coordinate system:
        #   X = right direction (width)
        #   Y = up direction (height)
        #   Z = view direction (depth - towards what we're looking at)
        
        half_width = door_width / 2 + crop_offset
        bottom = -crop_offset  # Below floor level
        top = door_height + crop_offset
        near_clip = 0.0  # At the section origin
        far_clip = view_offset + 2.0  # Past the door
        
        section_box = BoundingBoxXYZ()
        section_box.Min = XYZ(-half_width, bottom, near_clip)
        section_box.Max = XYZ(half_width, top, far_clip)
        
        # Set the transform for the section box
        # This defines the orientation of the section in world space
        transform = Transform.Identity
        transform.Origin = section_origin
        transform.BasisX = right_direction
        transform.BasisY = up_direction
        transform.BasisZ = view_direction
        
        section_box.Transform = transform
        
        if DEBUG_MODE:
            logger.debug("  Section box: Min=({:.2f}, {:.2f}, {:.2f}), Max=({:.2f}, {:.2f}, {:.2f})".format(
                section_box.Min.X, section_box.Min.Y, section_box.Min.Z,
                section_box.Max.X, section_box.Max.Y, section_box.Max.Z))
            logger.debug("  Creating ViewSection...")
        
        # Create the section view
        section_view = ViewSection.CreateSection(doc, section_type_id, section_box)
        
        if not section_view:
            logger.error("  Failed to create section view!")
            return None
        
        if DEBUG_MODE:
            logger.debug("  Section view created: {} (Id={}, ViewType={})".format(
                section_view.Name, section_view.Id.IntegerValue, section_view.ViewType))
        
        # Set the scale
        section_view.Scale = scale
        if DEBUG_MODE:
            logger.debug("  Set scale to 1:{}".format(scale))
        
        if DEBUG_MODE:
            # Verify view direction
            actual_view_dir = section_view.ViewDirection
            logger.debug("  Actual View Direction: ({:.3f}, {:.3f}, {:.3f})".format(
                actual_view_dir.X, actual_view_dir.Y, actual_view_dir.Z))
        
        # Set view name
        view_name = "{}{}".format(name_prefix, door_mark)
        try:
            section_view.Name = view_name
            if DEBUG_MODE:
                logger.debug("  View renamed to: {}".format(view_name))
        except:
            for i in range(1, 100):
                try:
                    new_name = "{} ({})".format(view_name, i)
                    section_view.Name = new_name
                    if DEBUG_MODE:
                        logger.debug("  View renamed to: {} (duplicate avoided)".format(new_name))
                    break
                except:
                    continue
        
        # Apply view template
        if template_id:
            try:
                section_view.ViewTemplateId = template_id
                if DEBUG_MODE:
                    logger.debug("  Applied view template: Id={}".format(template_id.IntegerValue))
            except Exception as ex:
                if DEBUG_MODE:
                    logger.debug("  Could not apply template: {}".format(ex))
        
        # Set view phase to match door's created phase
        self._set_view_phase_from_door(section_view, door)
        
        # Activate crop box
        section_view.CropBoxActive = True
        section_view.CropBoxVisible = False
        
        # Set annotation crop offset (small value to keep annotations tight)
        try:
            ann_crop_param = section_view.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
            if ann_crop_param and not ann_crop_param.IsReadOnly:
                ann_crop_param.Set(1)  # Enable annotation crop
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("  Could not set annotation crop: {}".format(ex))
        
        # Isolate the door element in the view (temporary isolation - keeps cyan border)
        try:
            # Create a list with just the door element ID
            door_ids = List[ElementId]()
            door_ids.Add(door.Id)
            
            if DEBUG_MODE:
                logger.debug("  Attempting to isolate door Id={} in view".format(door.Id.IntegerValue))
            
            # Apply temporary isolation - this shows the cyan border and can be reset
            # Do NOT call ConvertTemporaryHideIsolateToPermanent() to keep it temporary
            section_view.IsolateElementsTemporary(door_ids)
            if DEBUG_MODE:
                logger.debug("  Door temporarily isolated in view")
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("  Could not isolate door: {}".format(ex))
        
        if DEBUG_MODE:
            logger.debug("  Section created successfully!")
        return section_view
    
    def _set_view_phase_from_door(self, view, door):
        """Set the view's phase to match the door's phase created."""
        try:
            # Get door's phase created
            door_phase_param = door.get_Parameter(BuiltInParameter.PHASE_CREATED)
            if door_phase_param and door_phase_param.HasValue:
                door_phase_id = door_phase_param.AsElementId()
                
                # Set view phase
                view_phase_param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
                if view_phase_param and not view_phase_param.IsReadOnly:
                    view_phase_param.Set(door_phase_id)
                    if DEBUG_MODE:
                        logger.debug("Set view phase to match door phase: {}".format(door_phase_id))
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("Could not set view phase: {}".format(ex))
    
    # =========================================================================
    # Grouping and Sorting
    # =========================================================================
    
    def _group_doors(self, doors, group_param):
        """Group doors by parameter value."""
        groups = OrderedDict()
        for door_data in doors:
            value = door_data.get_group_value(group_param) or "Ungrouped"
            if value not in groups:
                groups[value] = []
            groups[value].append(door_data)
        return groups
    
    def _sort_doors(self, doors, sort_param):
        """Sort doors by parameter value."""
        return sorted(doors, key=lambda d: d.get_sort_value(sort_param) or "")
    
    # =========================================================================
    # Sheet Creation and Viewport Placement
    # =========================================================================
    
    def _get_viewport_size(self, view):
        """Get the viewport size on sheet based on view's crop box and scale."""
        try:
            if view.CropBoxActive:
                crop_box = view.CropBox
                # Get model-space dimensions
                model_width = abs(crop_box.Max.X - crop_box.Min.X)
                model_height = abs(crop_box.Max.Y - crop_box.Min.Y)
                
                # Convert to sheet space (divide by scale)
                scale = view.Scale
                sheet_width = model_width / scale
                sheet_height = model_height / scale
                
                return (sheet_width, sheet_height)
        except Exception as ex:
            if DEBUG_MODE:
                logger.debug("Could not get viewport size: {}".format(ex))
        
        # Fallback to estimate
        return (0.1, 0.15)  # ~3" x 4.5" default
    
    def _calculate_grid_layout(self, grouped_doors, cell_width, cell_height, sheet_margin):
        """Calculate grid layout dimensions and row count.
        
        Pure Python calculation - no Revit API calls.
        
        Returns:
            tuple: (cols, row_count, sheet_width, sheet_height, start_x, start_y)
        """
        total_viewports = len(self.elevation_views)
        cols = min(self.GRID_COLUMNS, total_viewports)
        
        # Calculate required width based on content
        content_width = cols * cell_width
        
        # Count rows needed (including group gaps)
        row_count = 0
        col = 0
        for group_name, doors in grouped_doors.items():
            for door_data in doors:
                if door_data.door_id not in self.elevation_views:
                    continue
                if col >= cols:
                    col = 0
                    row_count += 1
                col += 1
            if col > 0:  # End of group, start new row + gap
                col = 0
                row_count += 1 + self.GROUP_GAP_ROWS
        
        # Remove trailing gap rows
        row_count = max(1, row_count - self.GROUP_GAP_ROWS)
        
        content_height = row_count * cell_height
        
        # Sheet size = content + margins
        sheet_width = content_width + 2 * sheet_margin
        sheet_height = content_height + 2 * sheet_margin
        
        # Ensure minimum sheet size (at least A3 landscape ~420mm x 297mm = ~1.38ft x 0.97ft)
        sheet_width = max(sheet_width, 1.5)
        sheet_height = max(sheet_height, 1.0)
        
        # Starting position (top-left of usable area, centered in cell)
        start_x = sheet_margin + cell_width / 2
        start_y = sheet_height - sheet_margin - cell_height / 2
        
        return (cols, row_count, sheet_width, sheet_height, start_x, start_y)
    
    def _calculate_viewport_positions(self, grouped_doors, cols, cell_width, cell_height, start_x, start_y):
        """Calculate all viewport positions upfront.
        
        Pure Python calculation - no Revit API calls.
        
        Returns:
            list: List of (door_id, x, y) tuples
        """
        positions = []
        col = 0
        row = 0
        
        for group_name, doors in grouped_doors.items():
            for door_data in doors:
                # Skip if no elevation was created
                if door_data.door_id not in self.elevation_views:
                    continue
                
                # Check if we need a new row (column overflow)
                if col >= cols:
                    col = 0
                    row += 1
                
                # Calculate position
                x = start_x + col * cell_width
                y = start_y - row * cell_height
                
                positions.append((door_data.door_id, x, y))
                col += 1
            
            # Add gap rows between groups (start new row + gap)
            if col > 0:  # Only add gap if we placed items in this group
                col = 0
                row += 1 + self.GROUP_GAP_ROWS  # Move to new row + gap
        
        return positions
    
    def _create_sheets_and_place_viewports(self, grouped_doors, sheet_prefix, scale):
        """Create a single sheet and place all elevation viewports in grid layout."""
        
        # First, calculate actual viewport sizes for all views
        max_width = 0
        max_height = 0
        
        for door_id, view in self.elevation_views.items():
            width, height = self._get_viewport_size(view)
            max_width = max(max_width, width)
            max_height = max(max_height, height)
        
        # Ensure minimum cell size
        if max_width < 0.01:
            max_width = 0.15  # Default ~45mm
        if max_height < 0.01:
            max_height = 0.2  # Default ~60mm
        
        # Add padding for viewport labels/titles and spacing between viewports
        cell_padding = 0.08  # feet (~25mm padding for title block)
        cell_width = max_width + self.VIEWPORT_SPACING + cell_padding
        cell_height = max_height + self.VIEWPORT_SPACING + cell_padding
        sheet_margin = 0.15  # feet margin around edges (~45mm)
        
        if DEBUG_MODE:
            logger.debug("Max viewport size: {:.3f} x {:.3f} ft, cell size: {:.3f} x {:.3f} ft".format(
                max_width, max_height, cell_width, cell_height))
        
        # Performance: Pre-calculate all grid layout dimensions (pure Python)
        cols, row_count, sheet_width, sheet_height, start_x, start_y = self._calculate_grid_layout(
            grouped_doors, cell_width, cell_height, sheet_margin)
        
        if DEBUG_MODE:
            logger.debug("Grid: {} cols x {} rows, sheet size: {:.3f} x {:.3f} ft".format(
                cols, row_count, sheet_width, sheet_height))
        
        # Performance: Pre-calculate all viewport positions (pure Python)
        positions = self._calculate_viewport_positions(
            grouped_doors, cols, cell_width, cell_height, start_x, start_y)
        
        if not positions:
            logger.warning("No positions calculated for viewports")
            return
        
        # Create a single sheet (Revit API call)
        sheet = ViewSheet.Create(doc, ElementId.InvalidElementId)
        sheet.SheetNumber = "{}001".format(sheet_prefix)
        sheet.Name = "Door Elevations"
        self.created_sheets.append(sheet)
        
        if DEBUG_MODE:
            logger.debug("Created 1 sheet (size: {:.2f} x {:.2f} ft), placing {} viewports".format(
                sheet_width, sheet_height, len(positions)))
        
        # Place viewports (Revit API calls - tight loop)
        for door_id, x, y in positions:
            elev_view = self.elevation_views.get(door_id)
            if not elev_view:
                continue
            
            viewport_center = XYZ(x, y, 0)
            
            try:
                Viewport.Create(doc, sheet.Id, elev_view.Id, viewport_center)
            except Exception as ex:
                logger.warning("Could not place viewport for door {}: {}".format(
                    door_id, ex))


# =============================================================================
# Main Entry Point
# =============================================================================

if __name__ == '__main__':
    # Check for doors in project
    doors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors)\
        .WhereElementIsNotElementType().ToElements()
    
    if not doors:
        forms.alert("No doors found in the project.", title="No Doors")
    else:
        # Check for section view types
        vfts = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
        section_types = [vft for vft in vfts if vft.ViewFamily == ViewFamily.Section]
        
        # Debug: log all view family types
        for vft in vfts:
            try:
                name = vft.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
                logger.debug("ViewFamilyType: {} - ViewFamily: {}".format(name, vft.ViewFamily))
            except:
                pass
        
        if not section_types:
            forms.alert("No section view types found in the project.", title="Error")
        else:
            window = DoorElevationsWindow()
            window.ShowDialog()
