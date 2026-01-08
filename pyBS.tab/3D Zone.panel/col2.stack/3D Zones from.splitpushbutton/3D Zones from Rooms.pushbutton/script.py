# -*- coding: utf-8 -*-
"""Create 3D Zone Generic Model family instances from Room boundaries.

Creates Generic Model family instances using the 3DZone.rfa template,
replacing the extrusion profile with each room's boundary loops.
"""

__title__ = "Create 3D Zones from Rooms"
__author__ = "Byggstyrning AB"
__doc__ = "Create Generic Model family instances from Room boundaries using 3DZone.rfa template"
__highlight__ = 'new'

# Import standard libraries
import sys
import os.path as op

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.DB.Structure import StructuralType

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Import WPF for UI
import System.Windows

# Add the extension directory to the path
logger_temp = script.get_logger()
script_dir = script.get_script_path()
# Calculate extension directory by going up from script directory
# Structure: extension_root/pyBS.tab/3D Zone.panel/col2.stack/3D Zones from Rooms.pushbutton/
pushbutton_dir = script_dir
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)  # Go up one more level to get extension root
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Initialize logger
logger = script.get_logger()

# Import shared modules
from zone3d.spatial_adapter import RoomAdapter
from zone3d.selector_dialog import SpatialSelectorWindow, SpatialElementItem
from zone3d.zone_creator import create_zones_from_spatial_elements

# --- Room-Specific Classes ---

class RoomItem(SpatialElementItem):
    """Represents a room item in the filter dialog."""
    def __init__(self, room, doc, has_zone=False):
        adapter = RoomAdapter()
        super(RoomItem, self).__init__(room, doc, has_zone, adapter, "room")


class RoomSelectorWindow(SpatialSelectorWindow):
    """Custom WPF window for selecting rooms with search functionality."""
    
    def __init__(self, element_items, pushbutton_dir, extension_dir, doc=None):
        """Initialize the room selector window.
        
        Args:
            element_items: List of RoomItem objects
            pushbutton_dir: Pushbutton directory (for XAML files)
            extension_dir: Extension root directory (for styles)
            doc: Revit document (optional, kept for backward compatibility)
        """
        # Call parent init
        super(RoomSelectorWindow, self).__init__(element_items, pushbutton_dir, extension_dir)
        
        # Bind DataGrid instead of ListView
        if hasattr(self, 'roomsDataGrid'):
            self.roomsDataGrid.ItemsSource = self.filtered_items
    
    def get_xaml_filename(self):
        """Return XAML filename."""
        return "RoomSelector.xaml"
    
    def get_listview_name(self):
        """Return ListView control name (for backward compatibility)."""
        return "roomsListView"
    
    def get_element_attribute(self):
        """Return element attribute name."""
        return "room"
    
    def apply_filters(self):
        """Apply search text filter to items."""
        try:
            # Start with all items
            filtered = list(self.all_items)
            
            # Apply search text filter
            if hasattr(self, 'searchTextBox') and self.searchTextBox.Text:
                search_text = self.searchTextBox.Text.lower()
                filtered = [
                    item for item in filtered
                    if (search_text in (item.element_number or "").lower() or
                        search_text in (item.element_name or "").lower() or
                        search_text in (item.level_name or "").lower() or
                        search_text in (item.phase_name or "").lower())
                ]
            
            # Update filtered items
            self.filtered_items = filtered
            
            # Update DataGrid
            if hasattr(self, 'roomsDataGrid'):
                self.roomsDataGrid.ItemsSource = self.filtered_items
        except Exception as e:
            logger.debug("Error applying filters: {}".format(e))
    
    def searchTextBox_TextChanged(self, sender, args):
        """Handle search text box text changed event."""
        self.apply_filters()
    
    
    def create_button_click(self, sender, args):
        """Handle Create button click - collect selected elements and close."""
        # Collect all selected items from DataGrid
        selected_items = []
        if hasattr(self, 'roomsDataGrid'):
            for item in self.roomsDataGrid.SelectedItems:
                selected_items.append(item)
        
        # Extract elements from selected items
        element_attr = self.get_element_attribute()
        if selected_items:
            # Try to get element via attribute, fallback to .element
            self.selected_elements = []
            for item in selected_items:
                if hasattr(item, element_attr):
                    self.selected_elements.append(getattr(item, element_attr))
                elif hasattr(item, 'element'):
                    self.selected_elements.append(item.element)
        else:
            self.selected_elements = []
        
        # Close window
        self.Close()


def show_room_filter_dialog(rooms, doc):
    """Show a custom WPF dialog to filter and select rooms.
    
    Args:
        rooms: List of Room elements
        doc: Revit document
        
    Returns:
        List of selected Room elements, or None if cancelled
    """
    adapter = RoomAdapter()
    
    # OPTIMIZATION: Collect Generic Model instances ONCE and filter to only 3DZone families
    # This avoids collecting instances multiple times AND reduces iteration overhead
    all_generic_instances = FilteredElementCollector(doc)\
        .OfClass(FamilyInstance)\
        .OfCategory(BuiltInCategory.OST_GenericModel)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Filter to only 3DZone families (significant optimization - reduces iteration from ~1555 to ~21 instances)
    zone_instances_cache = []
    for instance in all_generic_instances:
        try:
            family_name = instance.Symbol.Family.Name
            if family_name.startswith("3DZone_Room-"):
                zone_instances_cache.append(instance)
        except:
            pass
    
    # Check for existing zones and create room items
    logger.debug("Checking for existing 3D zones...")
    room_items = []
    for room in rooms:
        has_zone = adapter.check_existing_zone(room, doc, zone_instances_cache)
        room_item = RoomItem(room, doc, has_zone)
        room_items.append(room_item)
    
    # Count existing zones
    existing_count = sum(1 for item in room_items if item.has_zone)
    logger.debug("Found {} rooms with existing 3D zones".format(existing_count))
    
    # Show custom WPF selection dialog
    dialog = RoomSelectorWindow(room_items, pushbutton_dir, extension_dir, doc)
    dialog.ShowDialog()
    
    # Return selected rooms (or None if cancelled)
    return dialog.selected_elements


# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Check user selection
    selection = revit.get_selection()
    selected_ids = selection.element_ids if selection else []
    
    # Filter selected elements to only Room instances
    selected_rooms = []
    if selected_ids:
        selected_elements = [doc.GetElement(eid) for eid in selected_ids]
        selected_rooms = [elem for elem in selected_elements if isinstance(elem, Room)]
    
    # If Rooms are selected, use only those; otherwise get all Rooms
    if selected_rooms:
        logger.debug("Found {} Rooms in selection".format(len(selected_rooms)))
        # Filter to only placed rooms
        placed_rooms = [room for room in selected_rooms if room.Area > 0]
        unplaced_count = len(selected_rooms) - len(placed_rooms)
        if unplaced_count > 0:
            logger.debug("Filtered out {} unplaced rooms from selection (Area = 0)".format(unplaced_count))
    else:
        # Get all Rooms
        spatial_elements = FilteredElementCollector(doc)\
            .OfClass(SpatialElement)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        # Filter for Room instances only, and only include placed rooms
        all_rooms = [elem for elem in spatial_elements if isinstance(elem, Room)]
        placed_rooms = [room for room in all_rooms if room.Area > 0]
    
    if not placed_rooms:
        if selected_rooms:
            forms.alert("Found {} Rooms in selection, but none are placed (all have Area = 0).\n\nPlease place rooms in the model before running this tool.".format(len(selected_rooms)),
                       title="No Placed Rooms", exitscript=True)
        else:
            # Check if there are any rooms at all
            spatial_elements = FilteredElementCollector(doc)\
                .OfClass(SpatialElement)\
                .WhereElementIsNotElementType()\
                .ToElements()
            all_rooms = [elem for elem in spatial_elements if isinstance(elem, Room)]
            if all_rooms:
                forms.alert("Found {} Rooms, but none are placed (all have Area = 0).\n\nPlease place rooms in the model before running this tool.".format(len(all_rooms)),
                           title="No Placed Rooms", exitscript=True)
            else:
                forms.alert("No Rooms found in the model.", title="No Rooms", exitscript=True)
    
    if selected_rooms:
        logger.debug("Using {} placed Rooms from selection".format(len(placed_rooms)))
    else:
        logger.debug("Found {} placed Rooms in model".format(len(placed_rooms)))
    
    # Create adapter
    adapter = RoomAdapter()
    
    # Create zones using shared orchestration function
    success_count, fail_count, failed_elements, created_instance_ids = create_zones_from_spatial_elements(
        placed_rooms,
        doc,
        adapter,
        extension_dir,
        pushbutton_dir,
        show_room_filter_dialog,
        "Rooms"
    )
