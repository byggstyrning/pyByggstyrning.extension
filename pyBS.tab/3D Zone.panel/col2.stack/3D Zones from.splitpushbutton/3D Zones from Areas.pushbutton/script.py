# -*- coding: utf-8 -*-
"""Create 3D Zone Generic Model family instances from Area boundaries.

Creates Generic Model family instances using the 3DZone.rfa template,
replacing the extrusion profile with each area's boundary loops.
"""

__title__ = "Create 3D Zones from Areas"
__author__ = "Byggstyrning AB"
__doc__ = "Create Generic Model family instances from Area boundaries using 3DZone.rfa template"
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
extension_dir = op.dirname(tab_dir)  # Go up one more level to get extension root
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

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
    
    def __init__(self, element_items, pushbutton_dir, extension_dir, doc=None):
        """Initialize the area selector window.
        
        Args:
            element_items: List of AreaItem objects
            pushbutton_dir: Pushbutton directory (for XAML files)
            extension_dir: Extension root directory (for styles)
            doc: Revit document (optional, kept for backward compatibility)
        """
        # Call parent init
        super(AreaSelectorWindow, self).__init__(element_items, pushbutton_dir, extension_dir)
        
        # Bind DataGrid instead of ListView
        if hasattr(self, 'areasDataGrid'):
            self.areasDataGrid.ItemsSource = self.filtered_items
    
    def get_xaml_filename(self):
        """Return XAML filename."""
        return "AreaSelector.xaml"
    
    def get_listview_name(self):
        """Return DataGrid control name (for backward compatibility)."""
        return "areasDataGrid"
    
    def get_element_attribute(self):
        """Return element attribute name."""
        return "area"
    
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
                        search_text in (item.level_name or "").lower())
                ]
            
            # Update filtered items
            self.filtered_items = filtered
            
            # Update DataGrid
            if hasattr(self, 'areasDataGrid'):
                self.areasDataGrid.ItemsSource = self.filtered_items
        except Exception as e:
            logger.debug("Error applying filters: {}".format(e))
    
    def searchTextBox_TextChanged(self, sender, args):
        """Handle search text box text changed event."""
        self.apply_filters()
    
    def create_button_click(self, sender, args):
        """Handle Create button click - collect selected elements and close."""
        # Collect all selected items from DataGrid
        selected_items = []
        if hasattr(self, 'areasDataGrid'):
            for item in self.areasDataGrid.SelectedItems:
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
    success_count, fail_count, failed_elements, created_instance_ids = create_zones_from_spatial_elements(
        placed_areas,
        doc,
        adapter,
        extension_dir,
        pushbutton_dir,
        show_area_filter_dialog,
        "Areas"
    )
