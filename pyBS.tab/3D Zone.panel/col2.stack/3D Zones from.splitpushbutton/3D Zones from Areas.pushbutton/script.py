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


# --- Source Model Classes ---

class AreaSourceItem(object):
    """Represents a model source (active or linked) in the source dropdown."""
    
    def __init__(self, source_doc, area_count, link_instance=None):
        """Initialize area source item.
        
        Args:
            source_doc: Revit Document containing the areas
            area_count: Number of placed areas in the source
            link_instance: RevitLinkInstance element (None for active model)
        """
        self.source_doc = source_doc
        self.area_count = area_count
        self.link_instance = link_instance
        
        if link_instance is None:
            # Active model
            self.link_transform = None
            self.display_name = "Active model ({} areas)".format(area_count)
        else:
            # Linked model
            self.link_transform = link_instance.GetTransform()
            link_name = link_instance.Name
            # Remove file extension and instance info if present
            if ':' in link_name:
                link_name = link_name.split(':')[0].strip()
            self.display_name = "{} ({} areas)".format(link_name, area_count)
    
    def __str__(self):
        return self.display_name
    
    def __repr__(self):
        return self.display_name
    
    def ToString(self):
        """Explicit ToString for WPF binding in IronPython."""
        return self.display_name


def get_area_sources(doc):
    """Get all available area sources (active model + linked models with areas).
    
    Args:
        doc: Host Revit document
        
    Returns:
        list: List of AreaSourceItem objects
    """
    sources = []
    
    # Active model areas
    try:
        spatial_elements = FilteredElementCollector(doc)\
            .OfClass(SpatialElement)\
            .WhereElementIsNotElementType()\
            .ToElements()
        all_areas = [elem for elem in spatial_elements if isinstance(elem, Area)]
        placed_areas = [area for area in all_areas if area.Area > 0]
        # Always add active model (even with 0 areas)
        sources.append(AreaSourceItem(doc, len(placed_areas), link_instance=None))
    except Exception as e:
        logger.debug("Error getting active model areas: {}".format(e))
        sources.append(AreaSourceItem(doc, 0, link_instance=None))
    
    # Linked model areas
    try:
        link_instances = FilteredElementCollector(doc)\
            .OfClass(RevitLinkInstance)\
            .ToElements()
        
        for link in link_instances:
            try:
                link_doc = link.GetLinkDocument()
                if not link_doc:
                    continue
                
                # Get placed areas from linked document
                link_spatial = FilteredElementCollector(link_doc)\
                    .OfClass(SpatialElement)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
                link_areas = [elem for elem in link_spatial if isinstance(elem, Area)]
                placed_link_areas = [area for area in link_areas if area.Area > 0]
                
                if len(placed_link_areas) > 0:
                    sources.append(AreaSourceItem(link_doc, len(placed_link_areas), link_instance=link))
                    
            except Exception as e:
                logger.debug("Error accessing linked document: {}".format(e))
                continue
        
    except Exception as e:
        logger.debug("Error getting linked documents: {}".format(e))
    
    return sources


def get_areas_from_source(source_item):
    """Collect placed areas from the given source model.
    
    Args:
        source_item: AreaSourceItem with the source document
        
    Returns:
        list: List of placed Area elements
    """
    try:
        source_doc = source_item.source_doc
        spatial_elements = FilteredElementCollector(source_doc)\
            .OfClass(SpatialElement)\
            .WhereElementIsNotElementType()\
            .ToElements()
        all_areas = [elem for elem in spatial_elements if isinstance(elem, Area)]
        placed_areas = [area for area in all_areas if area.Area > 0]
        return placed_areas
    except Exception as e:
        logger.debug("Error getting areas from source: {}".format(e))
        return []


# --- Area-Specific Classes ---

class AreaItem(SpatialElementItem):
    """Represents an area item in the filter dialog."""
    def __init__(self, area, source_doc, host_doc, has_zone=False):
        """Initialize area item.
        
        Args:
            area: Area element
            source_doc: Document containing the area (for level/phase lookups)
            host_doc: Host document (for existing zone checks)
            has_zone: Whether element already has a 3D Zone in the host doc
        """
        adapter = AreaAdapter()
        # Pass source_doc to parent for level/phase lookups
        super(AreaItem, self).__init__(area, source_doc, has_zone, adapter, "area")
        
        # Add area type (specific to Areas) - use source_doc for lookup
        self.area_type = adapter.get_area_type(area, source_doc)
        logger.debug("AreaItem created: number='{}', name='{}', area_type='{}'".format(
            self.element_number, self.element_name, self.area_type))


class AreaSelectorWindow(SpatialSelectorWindow):
    """Custom WPF window for selecting areas with search and model selection."""
    
    def __init__(self, host_doc, pushbutton_dir, extension_dir):
        """Initialize the area selector window with model selection support.
        
        Args:
            host_doc: Host Revit document
            pushbutton_dir: Pushbutton directory (for XAML files)
            extension_dir: Extension root directory (for styles)
        """
        self.host_doc = host_doc
        self.source_doc = host_doc  # Default to active model
        self.link_transform = None
        self._area_sources = []
        
        # Call parent init with empty items (will be populated after model selection)
        super(AreaSelectorWindow, self).__init__([], pushbutton_dir, extension_dir)
        
        # Discover area sources and populate dropdown
        self._load_area_sources()
        
        # Set up source model ComboBox selection changed event
        if hasattr(self, 'sourceModelComboBox'):
            self.sourceModelComboBox.SelectionChanged += self.on_source_model_changed
        
        # Bind DataGrid
        if hasattr(self, 'areasDataGrid'):
            self.areasDataGrid.ItemsSource = self.filtered_items
    
    def _load_area_sources(self):
        """Discover and populate available area sources."""
        try:
            self._area_sources = get_area_sources(self.host_doc)
            
            if hasattr(self, 'sourceModelComboBox'):
                self.sourceModelComboBox.ItemsSource = self._area_sources
                if self._area_sources:
                    self.sourceModelComboBox.SelectedIndex = 0
                    # Load areas from the default source (active model)
                    self._load_areas_from_source(self._area_sources[0])
        except Exception as e:
            logger.error("Error loading area sources: {}".format(e))
    
    def _load_areas_from_source(self, source_item):
        """Load areas from the given source and refresh the DataGrid.
        
        Args:
            source_item: AreaSourceItem with the source model
        """
        try:
            # Update tracked source info
            self.source_doc = source_item.source_doc
            self.link_transform = source_item.link_transform
            
            # Get areas from source
            areas = get_areas_from_source(source_item)
            
            # Create area items with existing zone check (against host doc)
            adapter = AreaAdapter()
            area_items = []
            for area in areas:
                has_zone = adapter.check_existing_zone(area, self.host_doc)
                area_item = AreaItem(area, source_item.source_doc, self.host_doc, has_zone)
                area_items.append(area_item)
            
            # Update items
            self.all_items = area_items
            self.filtered_items = list(area_items)
            
            # Clear search box
            if hasattr(self, 'searchTextBox'):
                self.searchTextBox.Text = ""
            
            # Refresh DataGrid
            if hasattr(self, 'areasDataGrid'):
                self.areasDataGrid.ItemsSource = self.filtered_items
            
            logger.debug("Loaded {} areas from source '{}'".format(
                len(area_items), source_item.display_name))
            
        except Exception as e:
            logger.error("Error loading areas from source: {}".format(e))
    
    def on_source_model_changed(self, sender, args):
        """Handle source model dropdown selection change."""
        try:
            selected = self.sourceModelComboBox.SelectedItem
            if selected:
                self._load_areas_from_source(selected)
        except Exception as e:
            logger.debug("Error handling source model change: {}".format(e))
    
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
                        search_text in (item.level_name or "").lower() or
                        search_text in (getattr(item, 'area_type', '') or "").lower())
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


# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Show the area selector window (handles model selection + area selection)
    dialog = AreaSelectorWindow(doc, pushbutton_dir, extension_dir)
    dialog.ShowDialog()
    
    # Get selected areas and source info
    selected_areas = dialog.selected_elements
    source_doc = dialog.source_doc
    link_transform = dialog.link_transform
    
    # Check if user cancelled or selected no elements
    if not selected_areas:
        logger.debug("No areas selected, exiting.")
        script.exit()
    
    logger.debug("User selected {} areas for 3D Zone creation".format(len(selected_areas)))
    
    # Create adapter
    adapter = AreaAdapter()
    
    # Create zones using shared orchestration function
    # Pass None for show_filter_dialog_func since dialog was already shown
    success_count, fail_count, failed_elements, created_instance_ids = create_zones_from_spatial_elements(
        selected_areas,
        doc,
        adapter,
        extension_dir,
        pushbutton_dir,
        None,  # Dialog already shown - skip internal dialog
        "Areas",
        source_doc=source_doc,
        link_transform=link_transform
    )
