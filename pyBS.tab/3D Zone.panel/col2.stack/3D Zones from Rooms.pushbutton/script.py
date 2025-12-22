# -*- coding: utf-8 -*-
"""Create 3D Zone Generic Model family instances from Room boundaries.

Creates Generic Model family instances using the 3DZone.rfa template,
replacing the extrusion profile with each room's boundary loops.
"""

__title__ = "Create 3D Zones from Rooms"
__author__ = "Byggstyrning AB"
__doc__ = "Create Generic Model family instances from Room boundaries using 3DZone.rfa template"

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
extension_dir = tab_dir
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

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
    
    def get_xaml_filename(self):
        """Return XAML filename."""
        return "RoomSelector.xaml"
    
    def get_listview_name(self):
        """Return ListView control name."""
        return "roomsListView"
    
    def get_element_attribute(self):
        """Return element attribute name."""
        return "room"


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
    dialog = RoomSelectorWindow(room_items, pushbutton_dir, extension_dir)
    dialog.ShowDialog()
    
    # Return selected rooms (or None if cancelled)
    return dialog.selected_elements


# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Get all Rooms
    spatial_elements = FilteredElementCollector(doc)\
        .OfClass(SpatialElement)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Filter for Room instances only, and only include placed rooms
    all_rooms = [elem for elem in spatial_elements if isinstance(elem, Room)]
    placed_rooms = [room for room in all_rooms if room.Area > 0]
    
    unplaced_count = len(all_rooms) - len(placed_rooms)
    if unplaced_count > 0:
        logger.debug("Filtered out {} unplaced rooms (Area = 0)".format(unplaced_count))
    
    if not placed_rooms:
        if all_rooms:
            forms.alert("Found {} Rooms, but none are placed (all have Area = 0).\n\nPlease place rooms in the model before running this tool.".format(len(all_rooms)),
                       title="No Placed Rooms", exitscript=True)
        else:
            forms.alert("No Rooms found in the model.", title="No Rooms", exitscript=True)
    
    logger.debug("Found {} placed Rooms ({} total, {} unplaced)".format(
        len(placed_rooms), len(all_rooms), unplaced_count))
    
    # Create adapter
    adapter = RoomAdapter()
    
    # Create zones using shared orchestration function
    create_zones_from_spatial_elements(
        placed_rooms,
        doc,
        adapter,
        extension_dir,
        pushbutton_dir,
        show_room_filter_dialog,
        "Rooms"
    )
