# -*- coding: utf-8 -*-
"""Create 3D Zone Generic Model family instances from Area boundaries.

Creates Generic Model family instances using the 3DZone.rfa template,
replacing the extrusion profile with each area's boundary loops.
"""

__title__ = "Create 3D Zones from Areas"
__author__ = "Byggstyrning AB"
__doc__ = "Create Generic Model family instances from Area boundaries using 3DZone.rfa template"

# Import standard libraries
import sys
import os.path as op

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import Room
# Area is imported from Autodesk.Revit.DB (not a submodule)

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Add the extension directory to the path
logger_temp = script.get_logger()
script_dir = script.get_script_path()
# Calculate extension directory by going up from script directory
# Structure: extension_root/pyBS.tab/3D Zone.panel/col2.stack/3D Zones from Areas.pushbutton/
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
from zone3d.spatial_adapter import AreaAdapter
from zone3d.selector_dialog import SpatialSelectorWindow, SpatialElementItem
from zone3d.zone_creator import create_zones_from_spatial_elements

# --- Area-Specific Classes ---

class AreaItem(SpatialElementItem):
    """Represents an area item in the filter dialog."""
    def __init__(self, area, doc, has_zone=False):
        adapter = AreaAdapter()
        super(AreaItem, self).__init__(area, doc, has_zone, adapter, "area")


class AreaSelectorWindow(SpatialSelectorWindow):
    """Custom WPF window for selecting areas with search functionality."""
    
    def get_xaml_filename(self):
        """Return XAML filename."""
        return "AreaSelector.xaml"
    
    def get_listview_name(self):
        """Return ListView control name."""
        return "areasListView"
    
    def get_element_attribute(self):
        """Return element attribute name."""
        return "area"


def show_area_filter_dialog(areas, doc):
    """Show a custom WPF dialog to filter and select areas.
    
    Args:
        areas: List of Area elements
        doc: Revit document
        
    Returns:
        List of selected Area elements, or None if cancelled
    """
    adapter = AreaAdapter()
    
    # Check for existing zones and create area items
    logger.debug("Checking for existing 3D zones...")
    area_items = []
    for area in areas:
        has_zone = adapter.check_existing_zone(area, doc)
        area_item = AreaItem(area, doc, has_zone)
        area_items.append(area_item)
    
    # Count existing zones
    existing_count = sum(1 for item in area_items if item.has_zone)
    logger.debug("Found {} areas with existing 3D zones".format(existing_count))
    
    # Show custom WPF selection dialog
    dialog = AreaSelectorWindow(area_items, pushbutton_dir, extension_dir)
    dialog.ShowDialog()
    
    # Return selected areas (or None if cancelled)
    return dialog.selected_elements


# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Get all Areas
    spatial_elements = FilteredElementCollector(doc)\
        .OfClass(SpatialElement)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Filter for Area instances only, and only include placed areas
    all_areas = [elem for elem in spatial_elements if isinstance(elem, Area)]
    placed_areas = [area for area in all_areas if area.Area > 0]
    
    unplaced_count = len(all_areas) - len(placed_areas)
    if unplaced_count > 0:
        logger.debug("Filtered out {} unplaced areas (Area = 0)".format(unplaced_count))
    
    if not placed_areas:
        if all_areas:
            forms.alert("Found {} Areas, but none are placed (all have Area = 0).\n\nPlease place areas in the model before running this tool.".format(len(all_areas)),
                       title="No Placed Areas", exitscript=True)
        else:
            forms.alert("No Areas found in the model.", title="No Areas", exitscript=True)
    
    logger.debug("Found {} placed Areas ({} total, {} unplaced)".format(
        len(placed_areas), len(all_areas), unplaced_count))
    
    # Create adapter
    adapter = AreaAdapter()
    
    # Create zones using shared orchestration function
    create_zones_from_spatial_elements(
        placed_areas,
        doc,
        adapter,
        extension_dir,
        pushbutton_dir,
        show_area_filter_dialog,
        "Areas"
    )
