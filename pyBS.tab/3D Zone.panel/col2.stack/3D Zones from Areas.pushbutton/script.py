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
import os
import shutil
import json
import time
import tempfile


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
import os.path as op
# Initialize logger early to use script.get_script_path()
logger_temp = script.get_logger()
script_dir = script.get_script_path()  # Gets the directory containing script.py
# Calculate extension directory by going up from script directory
# Structure: extension_root/pyBS.tab/3D Zone.panel/col2.stack/3D Zones from Areas.pushbutton/
pushbutton_dir = script_dir  # 3D Zones from Areas.pushbutton (directory)
splitpushbutton_dir = op.dirname(pushbutton_dir)  # col2.stack
stack_dir = op.dirname(splitpushbutton_dir)  # 3D Zone.panel
panel_dir = op.dirname(stack_dir)  # pyBS.tab
tab_dir = op.dirname(panel_dir)  # pyByggstyrning.extension (this IS the extension root!)
extension_dir = tab_dir  # The extension root IS tab_dir!
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Initialize logger
logger = script.get_logger()

# Import MMI utilities
try:
    from mmi.core import get_mmi_parameter_name
except ImportError as e:
    logger.warning("Could not import MMI utilities: {}".format(e))
    def get_mmi_parameter_name(doc):
        return "MMI"

# Template family path
TEMPLATE_FAMILY_NAME = "3DZone.rfa"
# The template is in the extension root directory
template_family_path = op.join(extension_dir, TEMPLATE_FAMILY_NAME)

# --- Helper Functions ---

def get_template_family_path():
    """Get the path to the template family."""
    # Log for debugging
    logger.debug("Looking for template at: {}".format(template_family_path))
    logger.debug("Extension dir: {}".format(extension_dir))
    logger.debug("Template exists: {}".format(op.exists(template_family_path)))
    
    if not op.exists(template_family_path):
        logger.error("Template family not found at: {}".format(template_family_path))
        # Try alternative path - maybe it's in the extension root differently
        alt_path = op.join(op.dirname(extension_dir), TEMPLATE_FAMILY_NAME)
        logger.debug("Trying alternative path: {}".format(alt_path))
        if op.exists(alt_path):
            logger.debug("Found template at alternative path: {}".format(alt_path))
            return alt_path
        return None
    return template_family_path

def inspect_template_family(family_doc):
    """Inspect the template family to find extrusion and parameter names.
    
    Returns:
        dict: {
            'extrusion': Extrusion element,
            'extrusion_start_param': FamilyParameter for start,
            'extrusion_end_param': FamilyParameter for end,
            'material_param': FamilyParameter for material,
            'subcategory': Category for '3D Zone'
        }
    """
    result = {
        'extrusion': None,
        'extrusion_start_param': None,
        'extrusion_end_param': None,
        'top_offset_param': None,  # User-facing parameter that controls height
        'bottom_offset_param': None,  # User-facing parameter for bottom offset
        'material_param': None,
        'subcategory': None
    }
    
    try:
        # Find the single extrusion
        extrusions = FilteredElementCollector(family_doc).OfClass(Extrusion).ToElements()
        if extrusions:
            result['extrusion'] = extrusions[0]
            logger.debug("Found extrusion in template: ID {}".format(result['extrusion'].Id))
        else:
            logger.warning("No extrusion found in template family")
            return result
        
        # Get family manager for parameters
        fm = family_doc.FamilyManager
        
        # Find extrusion start/end parameters
        # Template may use ExtrusionStart/ExtrusionEnd or NVExtrusionStart/NVExtrusionEnd
        param_names_to_try = [
            "NVExtrusionStart", "NVExtrusionEnd",  # Template family parameters (NV prefix - prioritized)
            "ExtrusionStart", "ExtrusionEnd",  # Alternative template family parameters
            "Extrusion Start", "Extrusion End",
            "Start", "End"
        ]
        
        for param_name in param_names_to_try:
            try:
                param = fm.get_Parameter(param_name)
                if param:
                    if "Start" in param_name or "start" in param_name.lower():
                        if not result['extrusion_start_param']:
                            result['extrusion_start_param'] = param
                            logger.debug("Found extrusion start parameter: {}".format(param_name))
                    elif "End" in param_name or "end" in param_name.lower():
                        if not result['extrusion_end_param']:
                            result['extrusion_end_param'] = param
                            logger.debug("Found extrusion end parameter: {}".format(param_name))
            except Exception as e:
                pass
        
        # Find Top Offset and Bottom Offset parameters (user-facing parameters)
        # These are family type parameters, not instance parameters
        top_offset_names = ["Top Offset", "TopOffset", "Top Offset (default)", "Top_Offset"]
        for param_name in top_offset_names:
            try:
                param = fm.get_Parameter(param_name)
                if param:
                    result['top_offset_param'] = param
                    logger.debug("Found Top Offset parameter: {}".format(param_name))
                    break
            except Exception as e:
                pass
        
        # Find Bottom Offset parameter
        bottom_offset_names = ["Bottom Offset", "BottomOffset", "Bottom Offset (default)", "Bottom_Offset"]
        for param_name in bottom_offset_names:
            try:
                param = fm.get_Parameter(param_name)
                if param:
                    result['bottom_offset_param'] = param
                    logger.debug("Found Bottom Offset parameter: {}".format(param_name))
                    break
            except Exception as e:
                pass
        
        # Find material parameter
        material_param_names = ["Material", "Material Parameter", "MaterialParam"]
        for param_name in material_param_names:
            try:
                param = fm.get_Parameter(param_name)
                if param:
                    result['material_param'] = param
                    logger.debug("Found material parameter: {}".format(param_name))
                    break
            except:
                pass
        
        # Find subcategory "3D Zone"
        try:
            gen_model_cat = Category.GetCategory(family_doc, BuiltInCategory.OST_GenericModel)
            if gen_model_cat:
                subcats = gen_model_cat.SubCategories
                for subcat in subcats:
                    if subcat.Name == "3D Zone":
                        result['subcategory'] = subcat
                        logger.debug("Found subcategory: 3D Zone")
                        break
        except Exception as e:
            logger.debug("Error finding subcategory: {}".format(e))
        
    except Exception as e:
        logger.error("Error inspecting template family: {}".format(e))
        import traceback
        logger.error(traceback.format_exc())
    
    return result

def check_area_has_3d_zone(area, doc):
    """Check if an area already has a 3D Zone instance.
    
    Args:
        area: Area element
        doc: Revit document
        
    Returns:
        bool: True if area has a 3D Zone instance, False otherwise
    """
    try:
        # Get area number and name
        # Areas don't have BuiltInParameter constants, use LookupParameter instead
        # Try multiple parameter name variations
        area_number_param = None
        area_number_str = None
        param_names_to_try = ["Number", "Area Number", "AREA_NUMBER", "AREA_NUM"]
        for param_name in param_names_to_try:
            area_number_param = area.LookupParameter(param_name)
            if area_number_param and area_number_param.HasValue:
                area_number_str = area_number_param.AsString()
                break
        
        area_name_param = None
        area_name_str = None
        name_param_names_to_try = ["Name", "Area Name", "AREA_NAME"]
        for param_name in name_param_names_to_try:
            area_name_param = area.LookupParameter(param_name)
            if area_name_param and area_name_param.HasValue:
                area_name_str = area_name_param.AsString()
                break
        
        if not area_number_str:
            return False
        
        # Find all Generic Model instances
        generic_instances = FilteredElementCollector(doc)\
            .OfClass(FamilyInstance)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        # Check if any instance matches this area by family name pattern
        # Family names are created as "3DZone_Area-{area_number}_{area_id}"
        # We'll check if any instance's family name matches this pattern
        area_id_str = str(area.Id.IntegerValue)
        expected_family_name_pattern = "3DZone_Area-{}_".format(area_number_str.replace(" ", "-"))
        
        found_family_names = []
        for instance in generic_instances:
            if instance.Category and instance.Category.Id == Category.GetCategory(doc, BuiltInCategory.OST_GenericModel).Id:
                # Get the family name from the instance's symbol
                try:
                    symbol = instance.Symbol
                    if symbol:
                        family = symbol.Family
                        if family:
                            family_name = family.Name
                            found_family_names.append(family_name)
                            # Check if family name starts with the expected pattern
                            if family_name.startswith(expected_family_name_pattern):
                                # Also verify it contains the area ID to be more precise
                                if area_id_str in family_name:
                                    return True
                except Exception as e:
                    pass
        
        return False
    except Exception as e:
        logger.debug("Error checking for existing 3D zone: {}".format(e))
        return False

def extract_area_boundary_loops(area, doc):
    """Extract boundary loops from an area.
    
    Returns:
        tuple: (loops: list of CurveLoop, insertion_point: XYZ, height: float)
    """
    loops = []
    insertion_point = None
    height = 0.0
    
    try:
        # Get boundary segments
        boundary_options = SpatialElementBoundaryOptions()
        boundary_segments = area.GetBoundarySegments(boundary_options)
        
        if not boundary_segments or len(boundary_segments) == 0:
            logger.warning("Area {} has no boundary segments".format(area.Id))
            return loops, insertion_point, height
        
        # Process each boundary loop
        all_points = []
        for segment_group in boundary_segments:
            curve_loop = CurveLoop()
            loop_points = []
            
            for segment in segment_group:
                curve = segment.GetCurve()
                if curve:
                    try:
                        start_pt = curve.GetEndPoint(0)
                        end_pt = curve.GetEndPoint(1)
                        loop_points.append(start_pt)
                        loop_points.append(end_pt)
                        all_points.append(start_pt)
                        all_points.append(end_pt)
                        
                        # Add curve to loop
                        curve_loop.Append(curve)
                    except Exception as e:
                        logger.debug("Error processing curve segment: {}".format(e))
                        continue
            
            # Check if loop has curves (CurveLoop is iterable in IronPython)
            try:
                # Try to iterate and check if loop has curves
                curve_list = list(curve_loop)
                if len(curve_list) > 0:
                    loops.append(curve_loop)
            except:
                # Loop is empty or error accessing it
                pass
        
        # Calculate insertion point (centroid of footprint)
        if all_points:
            min_x = min(pt.X for pt in all_points)
            max_x = max(pt.X for pt in all_points)
            min_y = min(pt.Y for pt in all_points)
            max_y = max(pt.Y for pt in all_points)
            min_z = min(pt.Z for pt in all_points)
            
            insertion_point = XYZ(
                (min_x + max_x) / 2.0,
                (min_y + max_y) / 2.0,
                min_z
            )
        
        # Calculate height based on level's "Storey Above" parameter
        # Get area level
        level_id = area.LevelId if hasattr(area, 'LevelId') and area.LevelId else None
        height = None  # None means use family default
        if level_id:
            level = doc.GetElement(level_id)
            if level:
                try:
                    # Get "Storey Above" parameter from level
                    # Try BuiltInParameter first
                    storey_above_param = None
                    try:
                        storey_above_param = level.get_Parameter(BuiltInParameter.LEVEL_STOREY_ABOVE)
                    except:
                        # Fallback: try LookupParameter
                        try:
                            storey_above_param = level.LookupParameter("Storey Above")
                        except:
                            pass
                    
                    if storey_above_param and storey_above_param.HasValue:
                        storey_above_id = storey_above_param.AsElementId()
                        
                        # Check if it's set to "default" (InvalidElementId or None)
                        if storey_above_id and storey_above_id != ElementId.InvalidElementId:
                            # Get the level above
                            level_above = doc.GetElement(storey_above_id)
                            if level_above:
                                base_z = level.Elevation
                                top_z = level_above.Elevation
                                height = top_z - base_z
                            else:
                                logger.debug("Area {}: Storey Above level not found, using family default".format(area.Id))
                        else:
                            # Storey Above is set to "default" - use family default height
                            logger.debug("Area {}: Storey Above is set to 'default', using family default height".format(area.Id))
                    else:
                        # Storey Above parameter not found or not set - fallback to next level
                        logger.debug("Area {}: Storey Above parameter not found, falling back to next level".format(area.Id))
                        base_z = level.Elevation
                        # Find level above
                        all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                        levels_sorted = sorted(all_levels, key=lambda l: l.Elevation)
                        top_z = base_z + 10.0  # Default 10 feet
                        for lvl in levels_sorted:
                            if lvl.Elevation > base_z:
                                top_z = lvl.Elevation
                                break
                        height = top_z - base_z
                except Exception as storey_error:
                    logger.debug("Error getting Storey Above parameter: {}, falling back to next level".format(storey_error))
                    # Fallback to next level calculation
                    base_z = level.Elevation
                    all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                    levels_sorted = sorted(all_levels, key=lambda l: l.Elevation)
                    top_z = base_z + 10.0
                    for lvl in levels_sorted:
                        if lvl.Elevation > base_z:
                            top_z = lvl.Elevation
                            break
                    height = top_z - base_z
        else:
            # Fallback if no level found
            logger.debug("Area {} has no level, using family default height".format(area.Id))
            height = None  # Use family default
        
    except Exception as e:
        logger.error("Error extracting area boundary loops: {}".format(e))
        import traceback
        logger.error(traceback.format_exc())
    
    return loops, insertion_point, height

# --- Area Filter Dialog ---

class AreaItem(object):
    """Represents an area item in the filter dialog."""
    def __init__(self, area, doc, has_zone=False):
        self.area = area
        self.has_zone = has_zone
        
        # Get area display info
        # Areas don't have BuiltInParameter constants, use LookupParameter instead
        area_number_param = area.LookupParameter("Number")
        if not area_number_param:
            area_number_param = area.LookupParameter("Area Number")
        self.area_number = area_number_param.AsString() if area_number_param and area_number_param.HasValue else "?"
        
        area_name_param = area.LookupParameter("Name")
        if not area_name_param:
            area_name_param = area.LookupParameter("Area Name")
        self.area_name = area_name_param.AsString() if area_name_param and area_name_param.HasValue else "Unnamed"
        
        # Get level name from area's LevelId property
        level_name = "?"
        level_id = area.LevelId if hasattr(area, 'LevelId') and area.LevelId else None
        if level_id:
            level_elem = doc.GetElement(level_id)
            if level_elem:
                level_name = level_elem.Name
        
        # Create display text with icon: (*) if zone exists, empty if not
        icon = "(*)" if has_zone else "   "
        self.display_text = "{} {} - {} ({})".format(icon, self.area_number, self.area_name, level_name)
    
    def __str__(self):
        """String representation for pyRevit forms."""
        return self.display_text
    
    def __repr__(self):
        """Representation for debugging."""
        return self.display_text

class AreaSelectorWindow(forms.WPFWindow):
    """Custom WPF window for selecting areas with search functionality."""
    
    def __init__(self, area_items):
        """Initialize the area selector window.
        
        Args:
            area_items: List of AreaItem objects
        """
        # Load XAML file
        xaml_path = op.join(pushbutton_dir, 'AreaSelector.xaml')
        forms.WPFWindow.__init__(self, xaml_path)
        
        # Load common styles programmatically
        self.load_styles()
        
        # Store all area items and filtered items
        self.all_area_items = area_items
        self.filtered_area_items = list(area_items)
        
        # Store selected areas (will be populated when Create is clicked)
        self.selected_areas = None
        
        # Bind collection to ListView
        self.areasListView.ItemsSource = self.filtered_area_items
        
        # Set up event handlers
        self.createButton.Click += self.create_button_click
        self.cancelButton.Click += self.cancel_button_click
    
    def load_styles(self):
        """Load the common styles ResourceDictionary."""
        try:
            styles_path = op.join(extension_dir, 'lib', 'styles', 'CommonStyles.xaml')
            
            if op.exists(styles_path):
                from System.Windows.Markup import XamlReader
                from System.IO import File
                
                # Read XAML content
                xaml_content = File.ReadAllText(styles_path)
                
                # Parse as ResourceDictionary
                styles_dict = XamlReader.Parse(xaml_content)
                
                # Merge into window resources
                if self.Resources is None:
                    from System.Windows import ResourceDictionary
                    self.Resources = ResourceDictionary()
                
                # Merge styles into existing resources
                if hasattr(styles_dict, 'MergedDictionaries'):
                    for merged_dict in styles_dict.MergedDictionaries:
                        self.Resources.MergedDictionaries.Add(merged_dict)
                
                # Copy individual resources
                for key in styles_dict.Keys:
                    self.Resources[key] = styles_dict[key]
        except Exception as e:
            logger.debug("Could not load styles: {}".format(e))
    
    def searchTextBox_TextChanged(self, sender, args):
        """Handle search text box text changed event."""
        try:
            search_text = sender.Text.lower() if sender.Text else ""
            
            # Filter area items based on search text
            if search_text:
                self.filtered_area_items = [
                    item for item in self.all_area_items
                    if search_text in item.display_text.lower()
                ]
            else:
                self.filtered_area_items = list(self.all_area_items)
            
            # Update ListView
            self.areasListView.ItemsSource = self.filtered_area_items
        except Exception as e:
            logger.debug("Error filtering areas: {}".format(e))
    
    def create_button_click(self, sender, args):
        """Handle Create button click - collect selected areas and close."""
        # Collect all selected area items
        selected_items = []
        for item in self.areasListView.SelectedItems:
            selected_items.append(item)
        
        # Extract Area elements from selected items
        if selected_items:
            self.selected_areas = [item.area for item in selected_items]
        else:
            self.selected_areas = []
        
        # Close window
        self.Close()
    
    def cancel_button_click(self, sender, args):
        """Handle Cancel button click - close without selection."""
        self.selected_areas = None
        self.Close()

def show_area_filter_dialog(areas, doc):
    """Show a custom WPF dialog to filter and select areas.
    
    Args:
        areas: List of Area elements
        doc: Revit document
        
    Returns:
        List of selected Area elements, or None if cancelled
    """
    # Check for existing zones and create area items
    logger.debug("Checking for existing 3D zones...")
    area_items = []
    for area in areas:
        has_zone = check_area_has_3d_zone(area, doc)
        area_item = AreaItem(area, doc, has_zone)
        area_items.append(area_item)
    
    # Count existing zones
    existing_count = sum(1 for item in area_items if item.has_zone)
    logger.debug("Found {} areas with existing 3D zones".format(existing_count))
    
    # Show custom WPF selection dialog
    dialog = AreaSelectorWindow(area_items)
    dialog.ShowDialog()
    
    # Return selected areas (or None if cancelled)
    return dialog.selected_areas

# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    app = doc.Application
    
    # Check if document supports families (must be a project document)
    if doc.IsFamilyDocument:
        forms.alert("This tool can only be used in project documents, not family documents.", 
                   title="Invalid Document Type", exitscript=True)
    
    # Verify template family exists
    template_path = get_template_family_path()
    if not template_path:
        # Log debug info
        logger.error("Extension dir: {}".format(extension_dir))
        logger.error("Template path: {}".format(template_family_path))
        logger.error("Template exists: {}".format(op.exists(template_family_path)))
        forms.alert("Template family '{}' not found at:\n{}\n\nExtension dir: {}".format(
            TEMPLATE_FAMILY_NAME, template_family_path, extension_dir), 
            title="Template Not Found", exitscript=True)
    
    # Create temporary directory for RFA files (will be cleaned up after loading)
    temp_dir = tempfile.mkdtemp(prefix="pyBS_3DZone_")
    logger.debug("Created temporary directory: {}".format(temp_dir))
    
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
    
    # Show area filter dialog
    areas_list = show_area_filter_dialog(placed_areas, doc)
    
    # Check if user cancelled or selected no areas
    if not areas_list:
        logger.debug("No areas selected, exiting.")
        script.exit()
    
    logger.debug("User selected {} areas for 3D Zone creation".format(len(areas_list)))
    
    # Check if template family exists in project first
    template_family_name_in_project = "3DZone"  # Name without .rfa extension
    project_template_family = None
    families = FilteredElementCollector(doc).OfClass(Family).ToElements()
    for fam in families:
        if fam.Name == template_family_name_in_project:
            project_template_family = fam
            logger.debug("Found template family '{}' in project".format(template_family_name_in_project))
            break
    
    # If template family doesn't exist in project, load it
    if not project_template_family:
        logger.debug("Template family '{}' not found in project, loading from file...".format(template_family_name_in_project))
        try:
            from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource
            
            class FamilyLoadOptions(IFamilyLoadOptions):
                def OnFamilyFound(self, familyInUse, overwriteParameterValues):
                    overwriteParameterValues[0] = True
                    return True
                
                def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
                    source[0] = FamilySource.Family
                    overwriteParameterValues[0] = True
                    return True
            
            load_options = FamilyLoadOptions()
            
            with revit.Transaction("Load 3DZone Template Family"):
                load_result = doc.LoadFamily(template_path, load_options)
                
                # Handle LoadFamily return value - may be bool or tuple (bool, Family)
                if isinstance(load_result, tuple):
                    project_template_family = load_result[1] if len(load_result) > 1 else None
                elif isinstance(load_result, bool):
                    if load_result:
                        # Find the newly loaded family by its name
                        families = FilteredElementCollector(doc).OfClass(Family).ToElements()
                        project_template_family = next((f for f in families if f.Name == template_family_name_in_project), None)
                    else:
                        project_template_family = None
                else:
                    project_template_family = load_result  # Direct Family object
                
                if project_template_family:
                    logger.debug("Successfully loaded template family '{}' into project".format(template_family_name_in_project))
                else:
                    logger.error("Failed to load template family from: {}".format(template_path))
                    forms.alert("Failed to load template family '{}' from:\n{}\n\nPlease ensure the file exists and try again.".format(
                        template_family_name_in_project, template_path),
                        title="Failed to Load Template", exitscript=True)
        except Exception as load_error:
            logger.error("Error loading template family: {}".format(load_error))
            import traceback
            logger.error(traceback.format_exc())
            forms.alert("Error loading template family:\n{}\n\nCheck logs for details.".format(str(load_error)),
                       title="Error Loading Template", exitscript=True)
    
    # Export the project family to a temporary file for use as template
    template_family_doc = None
    template_path_to_use = template_path
    
    if project_template_family:
        # Export the project family to a temporary file
        try:
            logger.debug("Exporting template family from project...")
            temp_template_path = op.join(temp_dir, "temp_3DZone_template.rfa")
            
            # Open the family document for editing
            # EditFamily opens the family in edit mode and returns the family document
            template_family_doc = doc.EditFamily(project_template_family)
            if template_family_doc:
                # Save it to temporary location
                save_options = SaveAsOptions()
                save_options.OverwriteExistingFile = True
                template_family_doc.SaveAs(temp_template_path, save_options)
                template_family_doc.Close(False)
                template_family_doc = None
                template_path_to_use = temp_template_path
                logger.debug("Exported template family to: {}".format(temp_template_path))
            else:
                logger.warning("Could not edit project family, falling back to file template")
                project_template_family = None
        except Exception as export_error:
            logger.warning("Error exporting project family: {}, falling back to file template".format(export_error))
            import traceback
            logger.debug(traceback.format_exc())
            project_template_family = None
            if template_family_doc:
                try:
                    template_family_doc.Close(False)
                except:
                    pass
                template_family_doc = None
    
    # Open template family document (either from project export or file)
    logger.debug("Opening template family for inspection...")
    try:
        template_family_doc = app.OpenDocumentFile(template_path_to_use)
        if not template_family_doc:
            raise Exception("Failed to open template family document")
        
        # Inspect template
        template_info = inspect_template_family(template_family_doc)
        
        if not template_info['extrusion']:
            template_family_doc.Close(False)
            forms.alert("Template family does not contain an extrusion element.", 
                       title="Invalid Template", exitscript=True)
        
        logger.debug("Template inspection complete")
        
    except Exception as e:
        if template_family_doc:
            template_family_doc.Close(False)
        logger.error("Error opening template family: {}".format(e))
        import traceback
        logger.error(traceback.format_exc())
        forms.alert("Error opening template family: {}\n\nCheck logs for details.".format(str(e)), 
                   title="Error", exitscript=True)
    
    # Process each area
    success_count = 0
    fail_count = 0
    created_families = []
    failed_areas = []  # List of dicts: {area, area_number_str, area_name_str, reason}
    
    logger.debug("Starting to process {} areas...".format(len(areas_list)))
    
    # First pass: Process all family documents (separate transactions per family doc)
    # Store area data for second pass (loading/placing in project)
    area_family_data = []  # List of dicts: {area, output_family_path, insertion_point, area_number_str, area_name_str}
    
    # Process areas with progress bar
    # Combined progress bar for both phases (creating families: 0-50%, loading/placing: 50-100%)
    total_areas = len(areas_list)
    with forms.ProgressBar(title="Creating 3D Zone Families and Placing Instances ({} areas)".format(total_areas)) as pb:
        # Phase 1: Process areas and create family documents (0-50% of progress)
        for area_idx, area in enumerate(areas_list):
            try:
                area_id = area.Id.IntegerValue
                
                # Get area name for file naming
                # Areas don't have BuiltInParameter constants, use LookupParameter instead
                area_name_param = area.LookupParameter("Name")
                if not area_name_param:
                    area_name_param = area.LookupParameter("Area Name")
                area_name_str = area_name_param.AsString() if area_name_param and area_name_param.HasValue else "Unnamed"
                
                area_number_param = area.LookupParameter("Number")
                if not area_number_param:
                    area_number_param = area.LookupParameter("Area Number")
                area_number_str = area_number_param.AsString() if area_number_param and area_number_param.HasValue else str(area_id)
                
                logger.debug("Processing area {} ({}/{}): {} - {}".format(
                    area_number_str, area_idx + 1, len(areas_list), area_name_str, area_id))
                
                # Update progress bar: Phase 1 is 0-50% (area_idx+1 out of total_areas = 50% max)
                # Progress = (area_idx + 1) / total_areas * 0.5 * 100
                phase1_progress = int((area_idx + 1) / float(total_areas) * 50)
                pb.update_progress(phase1_progress, 100)
                
                # Extract area boundary loops
                loops, insertion_point, height = extract_area_boundary_loops(area, doc)
                
                # Check for valid boundary data
                # height can be None (use family default) or > 0
                if not loops or not insertion_point:
                    logger.warning("Area {} (ID: {}) - invalid boundary data, skipping".format(
                        area_number_str, area_id))
                    fail_count += 1
                    failed_areas.append({
                        'area': area,
                        'area_number_str': area_number_str,
                        'area_name_str': area_name_str,
                        'reason': 'Invalid boundary data (no loops or no insertion point)'
                    })
                    continue
                
                # If height is explicitly set but <= 0, that's invalid
                if height is not None and height <= 0:
                    logger.warning("Area {} (ID: {}) - invalid height ({}), skipping".format(
                        area_number_str, area_id, height))
                    fail_count += 1
                    failed_areas.append({
                        'area': area,
                        'area_number_str': area_number_str,
                        'area_name_str': area_name_str,
                        'reason': 'Invalid height (height <= 0)'
                    })
                    continue
                
                # Log height (None means use family default)
                height_str = "None (use family default)" if height is None else str(height)
                logger.debug("Area {}: {} loops, height: {}, insertion point: ({:.2f}, {:.2f}, {:.2f})".format(
                    area_number_str, len(loops), height_str, insertion_point.X, insertion_point.Y, insertion_point.Z))
                
                # Create output family name and path (using temporary file)
                output_family_name = "3DZone_Area-{}_{}.rfa".format(area_number_str.replace(" ", "-"), area_id)
                output_family_path = op.join(temp_dir, output_family_name)
                family_name_without_ext = op.splitext(output_family_name)[0]  # Name without .rfa
                
                # Check if family already exists in project
                existing_family = None
                families = FilteredElementCollector(doc).OfClass(Family).ToElements()
                for fam in families:
                    if fam.Name == family_name_without_ext:
                        existing_family = fam
                        logger.debug("Family '{}' already exists in project, will reuse it".format(family_name_without_ext))
                        break
                
                # Only create/modify family document if it doesn't exist
                if not existing_family:
                    # Copy template to output location (use the template path we determined earlier)
                    try:
                        shutil.copy2(template_path_to_use, output_family_path)
                        logger.debug("Copied template to: {}".format(output_family_path))
                    except Exception as copy_error:
                        logger.error("Error copying template for area {}: {}".format(area_number_str, copy_error))
                        fail_count += 1
                        continue
                    
                    # Open the copied family document
                    family_doc = None
                    try:
                        family_doc = app.OpenDocumentFile(output_family_path)
                        if not family_doc:
                            raise Exception("Failed to open family document")
                    
                        # Get the extrusion (should be the same as template)
                        extrusions = FilteredElementCollector(family_doc).OfClass(Extrusion).ToElements()
                        if not extrusions:
                            raise Exception("No extrusion found in copied family")
                        
                        extrusion = extrusions[0]
                        logger.debug("Found extrusion in copied family: ID {}".format(extrusion.Id))
                        
                        # Get family manager
                        fm = family_doc.FamilyManager
                    
                        # Translate loops to be relative to insertion point (center near origin)
                        # Also translate to Z=0 to ensure sketch plane is perfectly horizontal
                        # The insertion point will be used when placing the instance
                        translated_loops = []
                        for loop_idx, loop in enumerate(loops):
                            translated_loop = CurveLoop()
                            loop_curves = []
                            for curve in loop:
                                try:
                                    # Validate curve before translation
                                    if curve is None:
                                        logger.debug("Area {}: Found None curve in loop {}, skipping".format(area_number_str, loop_idx))
                                        continue
                                    
                                    # Check if curve is valid (has valid endpoints)
                                    try:
                                        start_pt = curve.GetEndPoint(0)
                                        end_pt = curve.GetEndPoint(1)
                                        
                                        # Check for degenerate curves (zero length)
                                        if start_pt.DistanceTo(end_pt) < 0.001:  # Less than 1mm
                                            logger.debug("Area {}: Degenerate curve detected in loop {} (length < 1mm), skipping".format(area_number_str, loop_idx))
                                            continue
                                    except Exception as curve_check_error:
                                        logger.debug("Area {}: Error checking curve in loop {}: {}, skipping".format(area_number_str, loop_idx, curve_check_error))
                                        continue
                                    
                                    # Translate curve to be relative to insertion point AND set Z=0
                                    first_pt = curve.GetEndPoint(0)
                                    # Create translation that moves to origin and sets Z=0
                                    translation = Transform.CreateTranslation(XYZ(-insertion_point.X, -insertion_point.Y, -first_pt.Z))
                                    translated_curve = curve.CreateTransformed(translation)
                                    
                                    # Validate translated curve
                                    if translated_curve is None:
                                        logger.debug("Area {}: Translated curve is None in loop {}, skipping".format(area_number_str, loop_idx))
                                        continue
                                    
                                    translated_loop.Append(translated_curve)
                                    loop_curves.append(translated_curve)
                                except Exception as curve_error:
                                    logger.debug("Area {}: Error processing curve in loop {}: {}, skipping".format(area_number_str, loop_idx, curve_error))
                                    continue
                            
                            # Validate loop has at least 3 curves (minimum for a valid closed loop)
                            if len(loop_curves) < 3:
                                logger.debug("Area {}: Loop {} has only {} curves (minimum 3 required), skipping".format(area_number_str, loop_idx, len(loop_curves)))
                                continue
                            
                            # Check if loop is closed (first point should match last point)
                            try:
                                if translated_loop.IsOpen():
                                    logger.debug("Area {}: Loop {} is open (not closed), attempting to close".format(area_number_str, loop_idx))
                                    # Try to close the loop by adding a line from last to first point
                                    first_curve = loop_curves[0]
                                    last_curve = loop_curves[-1]
                                    first_pt = first_curve.GetEndPoint(0)
                                    last_pt = last_curve.GetEndPoint(1)
                                    
                                    # Only close if gap is small (< 1mm)
                                    gap = first_pt.DistanceTo(last_pt)
                                    if gap > 0.001:
                                        logger.debug("Area {}: Loop {} gap is {} (too large to auto-close)".format(area_number_str, loop_idx, gap))
                                        continue
                            except Exception as loop_check_error:
                                logger.debug("Area {}: Could not check if loop {} is closed: {}".format(area_number_str, loop_idx, loop_check_error))
                            
                            translated_loops.append(translated_loop)
                        
                        # Validate we have at least one valid loop
                        if not translated_loops:
                            logger.warning("Area {}: No valid loops after translation (likely open loop or invalid geometry), skipping".format(area_number_str))
                            fail_count += 1
                            failed_areas.append({
                                'area': area,
                                'area_number_str': area_number_str,
                                'area_name_str': area_name_str,
                                'reason': 'No valid loops after translation (open loop or invalid geometry)'
                            })
                            # Close family document if it was opened
                            if family_doc:
                                try:
                                    family_doc.Close(False)
                                except:
                                    pass
                                family_doc = None
                            continue
                        
                        # Attempt A: Try to edit the existing extrusion profile
                        profile_edited = False
                        try:
                            # Get the extrusion's sketch
                            sketch_id = extrusion.SketchId
                            if sketch_id and sketch_id != ElementId.InvalidElementId:
                                sketch = family_doc.GetElement(sketch_id)
                                if sketch:
                                    # Try to edit the sketch
                                    # Note: This may not work for family extrusions, but we'll try
                                    logger.debug("Attempting to edit extrusion sketch...")
                                    
                                    # For now, we'll skip direct sketch editing as it's complex
                                    # and go straight to recreation approach
                                    profile_edited = False
                        except Exception as edit_error:
                            logger.debug("Could not edit extrusion sketch (expected): {}".format(edit_error))
                            profile_edited = False
                        
                        # Fallback B: Delete and recreate extrusion
                        if not profile_edited:
                            logger.debug("Recreating extrusion...")
                            
                            # Create transaction in family document
                            # Suppress "off axis" warnings using FailureHandlingOptions
                            from Autodesk.Revit.DB import FailureProcessingResult, IFailuresPreprocessor
                            
                            class OffAxisFailurePreprocessor(IFailuresPreprocessor):
                                def PreprocessFailures(self, failuresAccessor):
                                    # Delete all warnings to suppress "off axis" warnings
                                    failures = failuresAccessor.GetFailureMessages()
                                    for failure in failures:
                                        # Delete warnings (not errors)
                                        if failure.GetSeverity().ToString() == "Warning":
                                            failuresAccessor.DeleteWarning(failure)
                                    return FailureProcessingResult.Continue
                            
                            t = Transaction(family_doc, "Recreate Extrusion")
                            t.Start()
                            # Get existing failure handling options and modify them (must be after Start())
                            failure_options = t.GetFailureHandlingOptions()
                            failure_options.SetFailuresPreprocessor(OffAxisFailurePreprocessor())
                            t.SetFailureHandlingOptions(failure_options)
                            try:
                                # Delete old extrusion
                                family_doc.Delete(extrusion.Id)
                                
                                # Create sketch plane at Z=0 (or appropriate level)
                                # Use the first loop's plane or create a horizontal plane
                                first_loop = translated_loops[0]
                                # Get a point from the first curve to determine Z
                                # In IronPython, CurveLoop can be accessed with indexing [0]
                                try:
                                    first_curve = first_loop[0]
                                except (IndexError, TypeError):
                                    # Try iterating if indexing doesn't work
                                    curve_list = list(first_loop)
                                    if not curve_list:
                                        raise Exception("First loop has no curves")
                                    first_curve = curve_list[0]
                                first_point = first_curve.GetEndPoint(0)
                                
                                # Create sketch plane at Z=0 to avoid "off axis" warnings
                                # Translate all curves to Z=0 first, then create horizontal plane
                                origin = XYZ(0, 0, 0)
                                normal = XYZ.BasisZ
                                plane = Plane.CreateByNormalAndOrigin(normal, origin)
                                sketch_plane = SketchPlane.Create(family_doc, plane)
                                
                                # Create CurveArrArray from loops
                                # Curves should already be at Z=0 from translation, but we'll suppress any "off axis" warnings
                                curve_arr_array = CurveArrArray()
                                total_curves = 0
                                for loop_idx, loop in enumerate(translated_loops):
                                    curve_arr = CurveArray()
                                    loop_curve_count = 0
                                    for curve in loop:
                                        curve_arr.Append(curve)
                                        loop_curve_count += 1
                                        total_curves += 1
                                    curve_arr_array.Append(curve_arr)
                                
                                # Create new extrusion using FamilyCreate (not FamilyManager)
                                # Extrusion start = 0, end = height (relative to sketch plane)
                                # Use FamilyCreate.NewExtrusion (not FamilyManager)
                                # If height is None (use family default), use a temporary height for creation
                                # The family's default Top Offset will control the actual height
                                family_create = family_doc.FamilyCreate
                                
                                # Use height if provided, otherwise use temporary default (family default will override)
                                extrusion_height = height if height is not None else 10.0  # Temporary default for NewExtrusion
                                
                                new_extrusion = family_create.NewExtrusion(True, curve_arr_array, sketch_plane, extrusion_height)
                                
                                if new_extrusion:
                                    logger.debug("Created new extrusion: ID {}".format(new_extrusion.Id))
                                    
                                    # Set subcategory if found
                                    if template_info['subcategory']:
                                        try:
                                            subcat_param = new_extrusion.get_Parameter(BuiltInParameter.FAMILY_ELEM_SUBCATEGORY)
                                            if subcat_param and not subcat_param.IsReadOnly:
                                                subcat_param.Set(template_info['subcategory'].Id)
                                                logger.debug("Set subcategory to: 3D Zone")
                                        except Exception as subcat_error:
                                            logger.debug("Could not set subcategory: {}".format(subcat_error))
                                    
                                    # FIX: Set FAMILY PARAMETER values FIRST, then associate
                                    # When associating a family parameter to an instance parameter in Revit,
                                    # it copies the CURRENT VALUE of the family parameter INTO the instance parameter.
                                    # So we must set the family parameter value FIRST, regenerate to commit it,
                                    # THEN associate (which will copy the correct value into the element parameter).
                                    
                                    # Set family parameter values FIRST (before association)
                                    if template_info['extrusion_start_param']:
                                        try:
                                            fm.Set(template_info['extrusion_start_param'], 0.0)
                                        except Exception as pre_set_error:
                                            logger.debug("Could not pre-set start family parameter: {}".format(pre_set_error))
                                    
                                    if template_info['extrusion_end_param']:
                                        try:
                                            fm.Set(template_info['extrusion_end_param'], height)
                                        except Exception as pre_set_error:
                                            logger.debug("Could not pre-set end family parameter: {}".format(pre_set_error))
                                    
                                    # Regenerate to commit family parameter values
                                    family_doc.Regenerate()
                                    
                                    # Verify family parameter values before association
                                    if template_info['extrusion_end_param']:
                                        try:
                                            # Get current family parameter value via FamilyType
                                            family_type_param = fm.CurrentType.get_Parameter(template_info['extrusion_end_param'].Definition)
                                            if family_type_param:
                                                current_family_value = family_type_param.AsDouble()
                                        except Exception as verify_error:
                                            logger.debug("Could not verify family parameter value: {}".format(verify_error))
                                    
                                    # Also check element parameter value BEFORE association
                                    if template_info['extrusion_end_param']:
                                        try:
                                            end_param_before = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                                        except Exception as elem_before_error:
                                            logger.debug("Could not get element parameter before association: {}".format(elem_before_error))
                                    
                                    # FIX: DO NOT associate parameters - this makes them read-only in the project
                                    # Instead, set the values directly on the element in the family editor
                                    # Then we can set them on instances in the project without read-only issues
                                    
                                    # Set element parameters directly (they're writable before association)
                                    if template_info['extrusion_start_param']:
                                        try:
                                            start_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_START_PARAM)
                                            if start_param and not start_param.IsReadOnly:
                                                start_param.Set(0.0)
                                                logger.debug("Set extrusion start parameter directly on element: 0.0")
                                        except Exception as set_error:
                                            logger.debug("Could not set start parameter directly: {}".format(set_error))
                                    
                                    if template_info['extrusion_end_param']:
                                        try:
                                            end_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                                            if end_param and not end_param.IsReadOnly:
                                                end_param.Set(height)
                                                logger.debug("Set extrusion end parameter directly on element: {}".format(height))
                                        except Exception as set_error:
                                            logger.debug("Could not set end parameter directly: {}".format(set_error))
                                    
                                    # Regenerate after setting element parameters
                                    family_doc.Regenerate()
                                    
                                    # FIX: Set Top Offset and Bottom Offset (user-facing parameters) BEFORE association
                                    # These are the parameters that actually control the height
                                    # The formulas (ExtrusionEnd = Top Offset) will automatically update ExtrusionEnd
                                    # If Top Offset doesn't exist, set ExtrusionEnd family parameter BEFORE association
                                    # so the element gets the correct value when associated
                                    height_set_in_family = False
                                    if height is not None:
                                        # Set Bottom Offset first
                                        if template_info['bottom_offset_param']:
                                            try:
                                                fm.Set(template_info['bottom_offset_param'], 0.0)
                                                logger.debug("Set bottom offset parameter: 0.0")
                                            except Exception as bottom_error:
                                                logger.debug("Could not set bottom offset parameter: {}".format(bottom_error))
                                        
                                        # Set Top Offset BEFORE association
                                        # ExtrusionEnd is ONLY for association, NOT for setting height value
                                        if template_info['top_offset_param']:
                                            try:
                                                fm.Set(template_info['top_offset_param'], height)
                                                logger.debug("Set top offset parameter: {}".format(height))
                                                height_set_in_family = True
                                            except Exception as top_error:
                                                logger.debug("Could not set top offset parameter: {}".format(top_error))
                                        else:
                                            # Top Offset not found - cannot set height
                                            # ExtrusionEnd is only for association, not for setting height value
                                            logger.debug("Top Offset parameter not found - cannot set height, will use family default")
                                            height_set_in_family = False
                                    
                                    # Regenerate after setting family parameters
                                    family_doc.Regenerate()
                                    
                                    # FIX: Associate ExtrusionEnd/ExtrusionStart to the extrusion element
                                    # This allows the formulas (ExtrusionEnd = Top Offset) to work
                                    # Association happens AFTER setting family parameters so element gets correct value
                                    if template_info['extrusion_start_param']:
                                        try:
                                            start_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_START_PARAM)
                                            if start_param:
                                                fm.AssociateElementParameterToFamilyParameter(start_param, template_info['extrusion_start_param'])
                                                logger.debug("Associated extrusion start parameter")
                                        except Exception as assoc_error:
                                            logger.debug("Could not associate start parameter: {}".format(assoc_error))
                                    
                                    if template_info['extrusion_end_param']:
                                        try:
                                            end_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                                            if end_param:
                                                fm.AssociateElementParameterToFamilyParameter(end_param, template_info['extrusion_end_param'])
                                                logger.debug("Associated extrusion end parameter")
                                        except Exception as assoc_error:
                                            logger.debug("Could not associate end parameter: {}".format(assoc_error))
                                    
                                    # Regenerate after association
                                    family_doc.Regenerate()
                                    
                                    # Handle case when height is None (use family default)
                                    if height is None:
                                        # Height is None - use family default, don't set Top Offset
                                        logger.debug("Area {}: Height is None, using family default Top Offset".format(area_number_str))
                                    
                                    # Regenerate after setting offset parameters
                                    family_doc.Regenerate()
                                    
                                    if template_info['material_param']:
                                        try:
                                            mat_param = new_extrusion.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                                            if mat_param:
                                                fm.AssociateElementParameterToFamilyParameter(mat_param, template_info['material_param'])
                                                logger.debug("Associated material parameter")
                                        except Exception as assoc_error:
                                            logger.debug("Could not associate material parameter: {}".format(assoc_error))
                                    
                                    # Set extrusion start/end values (verify they're still correct after association)
                                    
                                    if template_info['extrusion_start_param']:
                                        try:
                                            fm.Set(template_info['extrusion_start_param'], 0.0)
                                        except Exception as start_error:
                                            pass
                                    
                                    if template_info['extrusion_end_param']:
                                        # FIX: Set parameter value using multiple methods for reliability
                                        param_set_success = False
                                        
                                        try:
                                            
                                            # Method 1: Try setting via FamilyManager (family parameter)
                                            try:
                                                fm.Set(template_info['extrusion_end_param'], height)
                                                # Regenerate to ensure parameter value propagates to element
                                                family_doc.Regenerate()
                                                # If fm.Set() doesn't throw, assume it succeeded
                                                param_set_success = True
                                            except Exception as fm_error:
                                                pass
                                            
                                            # Method 2: Always set directly on element parameter as well (ensures value is correct)
                                            # Even if family parameter setting succeeded, set element parameter directly to be sure
                                            try:
                                                element_end_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                                                if element_end_param and not element_end_param.IsReadOnly:
                                                    element_end_param.Set(height)
                                                    # Regenerate after setting element parameter
                                                    family_doc.Regenerate()
                                                    param_set_success = True
                                            except Exception as elem_error:
                                                if not param_set_success:
                                                    # Only mark as failed if family parameter also failed
                                                    pass
                                            
                                            # Final verification
                                            if param_set_success:
                                                # Check element parameter (more reliable than family parameter)
                                                element_param_value = None
                                                try:
                                                    elem_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                                                    if elem_param:
                                                        element_param_value = elem_param.AsDouble()
                                                except:
                                                    pass
                                            else:
                                                logger.warning("Failed to set extrusion end parameter for area {} - value may default to template value".format(area_number_str))
                                            
                                        except Exception as end_error:
                                            logger.warning("Error setting extrusion end parameter for area {}: {}".format(area_number_str, end_error))
                                else:
                                    logger.debug("Area {}: Failed to create new extrusion (invalid geometry), skipping".format(area_number_str))
                                    t.RollBack()
                                    fail_count += 1
                                    failed_areas.append({
                                        'area': area,
                                        'area_number_str': area_number_str,
                                        'area_name_str': area_name_str,
                                        'reason': 'Failed to create new extrusion (invalid geometry)'
                                    })
                                    if family_doc:
                                        try:
                                            family_doc.Close(False)
                                        except:
                                            pass
                                        family_doc = None
                                    continue
                                
                                t.Commit()
                            except Exception as recreate_error:
                                t.RollBack()
                                logger.debug("Area {}: Error creating extrusion: {}, skipping".format(area_number_str, recreate_error))
                                fail_count += 1
                                failed_areas.append({
                                    'area': area,
                                    'area_number_str': area_number_str,
                                    'area_name_str': area_name_str,
                                    'reason': 'Error creating extrusion: {}'.format(str(recreate_error))
                                })
                                if family_doc:
                                    try:
                                        family_doc.Close(False)
                                    except:
                                        pass
                                    family_doc = None
                                continue
                    
                        # Save the family document
                        save_options = SaveAsOptions()
                        save_options.OverwriteExistingFile = True
                        family_doc.SaveAs(output_family_path, save_options)
                        logger.debug("Saved family: {}".format(output_family_name))
                        
                    except Exception as family_error:
                        logger.error("Error processing family document for area {}: {}".format(
                            area_number_str, family_error))
                        import traceback
                        logger.error(traceback.format_exc())
                        fail_count += 1
                        failed_areas.append({
                            'area': area,
                            'area_number_str': area_number_str,
                            'area_name_str': area_name_str,
                            'reason': 'Error processing family document: {}'.format(str(family_error))
                        })
                        if family_doc:
                            try:
                                family_doc.Close(False)
                            except:
                                pass
                        continue
                    finally:
                        # Keep family doc open for phase 2 (will be closed after loading)
                        # Don't close it here - we'll reuse it to avoid reopening from disk
                        pass
                    
                    # Store area data for second pass (loading/placing in project)
                    # Track if height was set in family editor (either Top Offset or ExtrusionEnd fallback)
                    area_family_data.append({
                        'area': area,
                        'output_family_path': output_family_path,
                        'insertion_point': insertion_point,
                        'area_number_str': area_number_str,
                        'area_name_str': area_name_str,
                        'output_family_name': output_family_name,
                        'family_name_without_ext': family_name_without_ext,
                        'existing_family': existing_family,  # None if needs to be loaded, Family object if already exists
                        'family_doc': family_doc,  # Keep reference to open family doc (will close after loading)
                        'height': height,  # Store height to set on instance after creation
                        'height_set_in_family': height_set_in_family,  # Track if height was set in family editor
                        'template_info': template_info  # Store template info for later use
                    })
                else:
                    # Family already exists, skip family document processing
                    logger.debug("Skipping family document processing for area {} - family already exists".format(area_number_str))
                    
                    # Store area data for second pass (loading/placing in project)
                    # For existing families, height_set_in_family is False (family wasn't modified)
                    area_family_data.append({
                        'area': area,
                        'output_family_path': None,  # No file path since family already exists
                        'insertion_point': insertion_point,
                        'area_number_str': area_number_str,
                        'area_name_str': area_name_str,
                        'output_family_name': output_family_name,
                        'family_name_without_ext': family_name_without_ext,
                        'existing_family': existing_family,  # Use existing family
                        'height': height,  # Store height to set on instance after creation
                        'height_set_in_family': False,  # Family wasn't modified, so height wasn't set
                        'template_info': template_info  # Store template info for later use
                    })
                
            except Exception as e:
                logger.error("Error processing area {} (ID: {}): {}".format(
                    area_number_str if 'area_number_str' in locals() else "Unknown",
                    area_id if 'area_id' in locals() else "Unknown",
                    str(e)))
                import traceback
                logger.error(traceback.format_exc())
                fail_count += 1
                # Try to add to failed_areas if we have area info
                try:
                    if 'area' in locals():
                        failed_areas.append({
                            'area': area,
                            'area_number_str': area_number_str if 'area_number_str' in locals() else "Unknown",
                            'area_name_str': area_name_str if 'area_name_str' in locals() else "Unknown",
                            'reason': 'Error processing area: {}'.format(str(e))
                        })
                except:
                    pass  # If we can't add to failed_areas, just continue
                
                # Make sure family doc is closed
                if 'family_doc' in locals() and family_doc:
                    try:
                        family_doc.Close(False)
                    except:
                        pass
    
    # Second pass: Split into (A) Load families with NO transaction, (B) Place instances in one transaction
    # Combined progress bar for phase 2: 50-100% (loading: 50-75%, placing: 75-100%)
    logger.debug("Loading families and placing instances in project...")
    
    if area_family_data:
        # Create FamilyLoadOptions (required parameter)
        from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource
        
        class FamilyLoadOptions(IFamilyLoadOptions):
            def OnFamilyFound(self, familyInUse, overwriteParameterValues):
                overwriteParameterValues[0] = True
                return True
            
            def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
                source[0] = FamilySource.Family
                overwriteParameterValues[0] = True
                return True
        
        load_options = FamilyLoadOptions()
        
        # Combined progress bar for phase 2 (50-100% of total)
        total_areas = len(area_family_data)
        with forms.ProgressBar(title="Loading Families and Placing Instances ({} areas)".format(total_areas)) as pb:
            # A) Load all families with NO transaction open (required for familyDoc.LoadFamily)
            # Progress: 50-75% (first half of phase 2)
            for area_idx, area_data in enumerate(area_family_data):
                area = area_data['area']
                output_family_path = area_data['output_family_path']
                area_number_str = area_data['area_number_str']
                output_family_name = area_data['output_family_name']
                existing_family = area_data.get('existing_family')
                
                # Update progress bar: Loading phase is 50-75% (first half of phase 2)
                # Progress = 50 + (area_idx + 1) / total_areas * 25
                loading_progress = int(50 + (area_idx + 1) / float(total_areas) * 25)
                pb.update_progress(loading_progress, 100)
                
                if existing_family:
                    area_data['loaded_family'] = existing_family
                    continue
                
                load_result = None
                load_family_doc = area_data.get('family_doc')  # Reuse family doc from phase 1 if available
                try:
                    if load_family_doc:
                        # Family doc already open from phase 1 - use it directly (no reopen needed!)
                        load_result = load_family_doc.LoadFamily(doc, load_options)
                    else:
                        # Fallback: reopen from disk if doc not available
                        load_family_doc = app.OpenDocumentFile(output_family_path)
                        load_result = load_family_doc.LoadFamily(doc, load_options)
                except Exception as e_load_from_doc:
                    # Safety fallback (legacy slow path)
                    load_result = doc.LoadFamily(output_family_path, load_options)
                finally:
                    # Close family doc after loading (whether reused or reopened)
                    if load_family_doc:
                        try:
                            load_family_doc.Close(False)
                        except:
                            pass
                
                # Handle return value - may be bool or tuple (bool, Family)
                loaded_family = None
                try:
                    if isinstance(load_result, tuple):
                        loaded_family = load_result[1] if len(load_result) > 1 else None
                    elif isinstance(load_result, bool):
                        if load_result:
                            family_name = op.splitext(op.basename(output_family_path))[0]
                            families = FilteredElementCollector(doc).OfClass(Family).ToElements()
                            loaded_family = next((f for f in families if f.Name == family_name), None)
                        else:
                            loaded_family = None
                    else:
                        loaded_family = load_result
                except:
                    loaded_family = None
                
                area_data['loaded_family'] = loaded_family
            
            # B) Place all instances in a single transaction
            # Progress: 75-100% (second half of phase 2)
            t = Transaction(doc, "Place 3D Zone Instances")
            t.Start()
            try:
                for area_idx, area_data in enumerate(area_family_data):
                    area = area_data['area']
                    insertion_point = area_data['insertion_point']
                    area_number_str = area_data['area_number_str']
                    area_name_str = area_data['area_name_str']
                    output_family_name = area_data['output_family_name']
                    loaded_family = area_data.get('loaded_family') or area_data.get('existing_family')
                    
                    # Update progress bar: Placing phase is 75-100% (second half of phase 2)
                    # Progress = 75 + (area_idx + 1) / total_areas * 25
                    placing_progress = int(75 + (area_idx + 1) / float(total_areas) * 25)
                    pb.update_progress(placing_progress, 100)
                    
                    try:
                        if not loaded_family:
                            logger.warning("Failed to load family: {}".format(output_family_name))
                            fail_count += 1
                            continue
                        
                        logger.debug("Loaded family: {}".format(loaded_family.Name))
                        
                        symbol_ids_set = loaded_family.GetFamilySymbolIds()
                        symbol_ids = list(symbol_ids_set) if symbol_ids_set else []
                        if not symbol_ids or len(symbol_ids) == 0:
                            logger.warning("No symbols found in loaded family")
                            fail_count += 1
                            continue
                        
                        symbol = doc.GetElement(symbol_ids[0])
                        
                        if symbol and not symbol.IsActive:
                            symbol.Activate()
                            doc.Regenerate()
                        
                        if not symbol:
                            logger.warning("No active symbol found for family {}".format(loaded_family.Name))
                            fail_count += 1
                            continue
                        
                        # SKIPPED: Phase 2 - Setting parameters on symbol before instance creation
                        # Height is already set in family editor (Phase 1), instances will inherit correct value
                        
                        # Place instance at insertion point
                        if symbol:
                            # Get area level
                            level_id = area.LevelId if hasattr(area, 'LevelId') and area.LevelId else None
                            level = doc.GetElement(level_id) if level_id else None
                            
                            # Create insertion point with Z=0 relative to level
                            # When placing with a level, Z coordinate is relative to level elevation
                            placement_point = XYZ(
                                insertion_point.X,
                                insertion_point.Y,
                                0.0  # Elevation from level = 0
                            )
                            
                            # Place instance
                            instance = doc.Create.NewFamilyInstance(
                                placement_point,
                                symbol,
                                level,
                                StructuralType.NonStructural
                            )
                            
                            if instance:
                                logger.debug("Placed instance for area {}: {}".format(
                                    area_number_str, instance.Id))
                                
                                # Copy area properties to instance
                                try:
                                    # Copy Name
                                    # Areas don't have BuiltInParameter constants, use LookupParameter instead
                                    area_name = area.LookupParameter("Name")
                                    if not area_name:
                                        area_name = area.LookupParameter("Area Name")
                                    if area_name and area_name.HasValue:
                                        inst_name = instance.LookupParameter("Name")
                                        if inst_name and not inst_name.IsReadOnly:
                                            inst_name.Set(area_name.AsString())
                                    
                                    # Copy Number
                                    area_number = area.LookupParameter("Number")
                                    if not area_number:
                                        area_number = area.LookupParameter("Area Number")
                                    if area_number and area_number.HasValue:
                                        inst_number = instance.LookupParameter("Number")
                                        if inst_number and not inst_number.IsReadOnly:
                                            inst_number.Set(area_number.AsString())
                                    
                                    # Copy MMI
                                    configured_mmi_param = get_mmi_parameter_name(doc)
                                    mmi_copied = False
                                    if configured_mmi_param:
                                        source_mmi = area.LookupParameter(configured_mmi_param)
                                        if source_mmi and source_mmi.HasValue:
                                            target_mmi = instance.LookupParameter(configured_mmi_param)
                                            if target_mmi and not target_mmi.IsReadOnly:
                                                if source_mmi.StorageType == StorageType.String:
                                                    value = source_mmi.AsString()
                                                    if value:
                                                        target_mmi.Set(value)
                                                        mmi_copied = True
                                    
                                    if not mmi_copied:
                                        source_mmi = area.LookupParameter("MMI")
                                        if source_mmi and source_mmi.HasValue:
                                            target_mmi = instance.LookupParameter("MMI")
                                            if target_mmi and not target_mmi.IsReadOnly:
                                                if source_mmi.StorageType == StorageType.String:
                                                    value = source_mmi.AsString()
                                                    if value:
                                                        target_mmi.Set(value)
                                                        
                                except Exception as prop_error:
                                    logger.debug("Error copying properties: {}".format(prop_error))
                                
                                created_families.append(output_family_name)
                                success_count += 1
                            else:
                                logger.warning("Failed to place instance for area {}".format(area_number_str))
                                fail_count += 1
                                failed_areas.append({
                                    'area': area,
                                    'area_number_str': area_number_str,
                                    'area_name_str': area_name_str,
                                    'reason': 'Failed to place instance'
                                })
                        else:
                            logger.warning("No active symbol found for family {}".format(loaded_family.Name))
                            fail_count += 1
                            
                    except Exception as area_load_error:
                        logger.error("Error loading/placing family for area {}: {}".format(
                            area_number_str, area_load_error))
                        import traceback
                        logger.error(traceback.format_exc())
                        fail_count += 1
                        failed_areas.append({
                            'area': area,
                            'area_number_str': area_number_str,
                            'area_name_str': area_name_str,
                            'reason': 'Error loading/placing family: {}'.format(str(area_load_error))
                        })
                
                t.Commit()
            except Exception as tx_error:
                t.RollBack()
                logger.error("Transaction failed, rolled back: {}".format(tx_error))
                import traceback
                logger.error(traceback.format_exc())
                fail_count += len(area_family_data) - success_count
    
    # Close template family doc
    if template_family_doc:
        template_family_doc.Close(False)
    
    # Clean up temporary files
    try:
        if 'temp_dir' in locals() and op.exists(temp_dir):
            shutil.rmtree(temp_dir)
            logger.debug("Cleaned up temporary directory: {}".format(temp_dir))
    except Exception as cleanup_error:
        logger.warning("Error cleaning up temporary directory: {}".format(cleanup_error))
    
    # Log results
    logger.debug("Created {} 3D Zone families from {} Areas ({} failed)".format(
        success_count, len(areas_list), fail_count))
    
    # Show failed areas with Linkify if any
    if failed_areas:
        output = script.get_output()
        output.print_md("## ⚠️ Failed Areas - Invalid Geometry")
        output.print_md("**{} area(s) could not be processed due to invalid geometry.**".format(len(failed_areas)))
        output.print_md("Click area ID to zoom to area in Revit")
        output.print_md("---")
        
        # Group by reason for better overview
        by_reason = {}
        for item in failed_areas:
            reason = item['reason']
            if reason not in by_reason:
                by_reason[reason] = []
            by_reason[reason].append(item)
        
        # Print by reason
        for reason, items in sorted(by_reason.items()):
            output.print_md("### {} ({} areas)".format(reason, len(items)))
            output.print_md("")
            
            # Print each area with clickable link
            for item in items:
                area = item['area']
                area_number = item['area_number_str']
                area_name = item['area_name_str']
                # Use linkify method from output window to create clickable link
                area_link = output.linkify(area.Id)
                output.print_md("- **Area {} - {}**: {}".format(area_number, area_name, area_link))
            
            output.print_md("")  # Empty line between reasons
        
        output.print_md("---")
        output.print_md("**Tip:** Fix the area boundaries in Revit and run the script again.")
