# -*- coding: utf-8 -*-
"""Create 3D Zone Generic Model family instances from MEP Space boundaries.

Creates Generic Model family instances using the 3DZone.rfa template,
replacing the extrusion profile with each space's boundary loops.
"""

__title__ = "Create 3D Zones from Spaces"
__author__ = "Byggstyrning AB"
__doc__ = "Create Generic Model family instances from MEP Space boundaries using 3DZone.rfa template"


# Import standard libraries
import sys
import os.path as op

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Mechanical import Space

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Add the extension directory to the path
# Structure: extension_root/pyBS.tab/3D Zone.panel/Buttons.stack/3D Zones from.splitpushbutton/3D Zones from Spaces.pushbutton/
script_dir = script.get_script_path()
pushbutton_dir = script_dir
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Initialize logger
logger = script.get_logger()

# Import shared modules
from zone3d.spatial_adapter import SpaceAdapter
from zone3d.selector_dialog import SpatialSelectorWindow, SpatialElementItem
from zone3d.zone_creator import create_zones_from_spatial_elements

# --- Space-Specific Classes ---

class SpaceItem(SpatialElementItem):
    """Represents a space item in the filter dialog."""
    def __init__(self, space, doc, has_zone=False):
        adapter = SpaceAdapter()
        super(SpaceItem, self).__init__(space, doc, has_zone, adapter, "space")


class SpaceSelectorWindow(SpatialSelectorWindow):
    """Custom WPF window for selecting spaces with search functionality."""
    
    def __init__(self, element_items, pushbutton_dir, extension_dir):
        """Initialize the space selector window.
        
        Args:
            element_items: List of SpaceItem objects
            pushbutton_dir: Pushbutton directory (for XAML files)
            extension_dir: Extension root directory (for styles)
        """
        super(SpaceSelectorWindow, self).__init__(element_items, pushbutton_dir, extension_dir)
        
        if hasattr(self, 'spacesDataGrid'):
            self.spacesDataGrid.ItemsSource = self.filtered_items
    
    def get_xaml_filename(self):
        """Return XAML filename."""
        return "SpaceSelector.xaml"
    
    def get_listview_name(self):
        """Return item list control name."""
        return "spacesDataGrid"
    
    def get_element_attribute(self):
        """Return element attribute name."""
        return "space"
    
    def apply_filters(self):
        """Apply search text filter to items."""
        try:
            filtered = list(self.all_items)
            
            if hasattr(self, 'searchTextBox') and self.searchTextBox.Text:
                search_text = self.searchTextBox.Text.lower()
                filtered = [
                    item for item in filtered
                    if (search_text in (item.element_number or "").lower() or
                        search_text in (item.element_name or "").lower() or
                        search_text in (item.level_name or "").lower() or
                        search_text in (item.phase_name or "").lower())
                ]
            
            self.filtered_items = filtered
            
            if hasattr(self, 'spacesDataGrid'):
                self.spacesDataGrid.ItemsSource = self.filtered_items
        except Exception as e:
            logger.debug("Error applying filters: {}".format(e))
    
    def searchTextBox_TextChanged(self, sender, args):
        """Handle search text box text changed event."""
        self.apply_filters()


def show_space_filter_dialog(spaces, doc):
    """Show a custom WPF dialog to filter and select spaces.
    
    Args:
        spaces: List of Space elements
        doc: Revit document
        
    Returns:
        List of selected Space elements, or None if cancelled
    """
    adapter = SpaceAdapter()
    prefix = adapter.get_family_name_prefix()
    
    # Collect Generic Model instances once, keep only Space 3DZone families
    all_generic_instances = FilteredElementCollector(doc)\
        .OfClass(FamilyInstance)\
        .OfCategory(BuiltInCategory.OST_GenericModel)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    zone_instances_cache = []
    for instance in all_generic_instances:
        try:
            if instance.Symbol.Family.Name.startswith(prefix):
                zone_instances_cache.append(instance)
        except:
            pass
    
    logger.debug("Checking for existing 3D zones...")
    space_items = []
    for space in spaces:
        has_zone = adapter.check_existing_zone(space, doc, zone_instances_cache)
        space_items.append(SpaceItem(space, doc, has_zone))
    
    existing_count = sum(1 for item in space_items if item.has_zone)
    logger.debug("Found {} spaces with existing 3D zones".format(existing_count))
    
    dialog = SpaceSelectorWindow(space_items, pushbutton_dir, extension_dir)
    dialog.ShowDialog()
    
    return dialog.selected_elements


def collect_all_spaces(doc):
    """Return all Space elements in the document."""
    spatial_elements = FilteredElementCollector(doc)\
        .OfClass(SpatialElement)\
        .WhereElementIsNotElementType()\
        .ToElements()
    return [elem for elem in spatial_elements if isinstance(elem, Space)]


# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Check user selection
    selection = revit.get_selection()
    selected_ids = selection.element_ids if selection else []
    
    selected_spaces = []
    if selected_ids:
        selected_elements = [doc.GetElement(eid) for eid in selected_ids]
        selected_spaces = [elem for elem in selected_elements if isinstance(elem, Space)]
    
    # If Spaces are selected, use only those; otherwise get all placed Spaces
    if selected_spaces:
        logger.debug("Found {} Spaces in selection".format(len(selected_spaces)))
        placed_spaces = [space for space in selected_spaces if space.Area > 0]
        unplaced_count = len(selected_spaces) - len(placed_spaces)
        if unplaced_count > 0:
            logger.debug("Filtered out {} unplaced spaces from selection (Area = 0)".format(unplaced_count))
    else:
        placed_spaces = [space for space in collect_all_spaces(doc) if space.Area > 0]
    
    if not placed_spaces:
        if selected_spaces:
            forms.alert("Found {} Spaces in selection, but none are placed (all have Area = 0).\n\nPlease place spaces in the model before running this tool.".format(len(selected_spaces)),
                        title="No Placed Spaces", exitscript=True)
        else:
            all_spaces = collect_all_spaces(doc)
            if all_spaces:
                forms.alert("Found {} Spaces, but none are placed (all have Area = 0).\n\nPlease place spaces in the model before running this tool.".format(len(all_spaces)),
                            title="No Placed Spaces", exitscript=True)
            else:
                forms.alert("No Spaces found in the model.", title="No Spaces", exitscript=True)
    
    if selected_spaces:
        logger.debug("Using {} placed Spaces from selection".format(len(placed_spaces)))
    else:
        logger.debug("Found {} placed Spaces in model".format(len(placed_spaces)))
    
    adapter = SpaceAdapter()
    
    create_zones_from_spatial_elements(
        placed_spaces,
        doc,
        adapter,
        extension_dir,
        pushbutton_dir,
        show_space_filter_dialog,
        "Spaces"
    )
