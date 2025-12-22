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
import os
import shutil
import json
import time
import tempfile

# #region agent log
# Debug logging setup for performance monitoring
DEBUG_LOG_PATH = r"c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log"
def debug_log(location, message, data=None):
    try:
        import json as json_mod
        log_entry = {
            "timestamp": int(time.time() * 1000),
            "location": location,
            "message": message,
            "data": data or {},
            "sessionId": "perf-session"
        }
        with open(DEBUG_LOG_PATH, "a") as f:
            f.write(json_mod.dumps(log_entry) + "\n")
    except:
        pass
# #endregion

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
# Structure: extension_root/pyBS.tab/3D Zone.panel/col2.stack/Create 3D Zones.pushbutton/
pushbutton_dir = script_dir  # Create 3D Zones.pushbutton (directory)
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

def sanitize_family_name(name):
    """Sanitize a string to be safe for use as a Revit family name.
    
    Removes or replaces characters that are invalid in family names:
    - Periods (.) - causes issues with family loading
    - Slashes (/, \\) - path separators
    - Special chars (*, ?, ", <, >, |, :) - Windows filename restrictions
    - Leading/trailing spaces
    
    Args:
        name: String to sanitize
        
    Returns:
        Sanitized string safe for family naming
    """
    if not name:
        return "Unknown"
    
    # Replace problematic characters
    replacements = {
        '.': '_',
        '/': '-',
        '\\': '-',
        '*': '',
        '?': '',
        '"': '',
        '<': '',
        '>': '',
        '|': '',
        ':': '-',
        ' ': '-',
    }
    
    result = name
    for char, replacement in replacements.items():
        result = result.replace(char, replacement)
    
    # Remove any double dashes or underscores
    while '--' in result:
        result = result.replace('--', '-')
    while '__' in result:
        result = result.replace('__', '_')
    
    # Strip leading/trailing dashes and underscores
    result = result.strip('-_')
    
    # Ensure we have something left
    if not result:
        return "Unknown"
    
    return result

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
        # Template uses ExtrusionStart and ExtrusionEnd
        param_names_to_try = [
            "ExtrusionStart", "ExtrusionEnd",  # Template family parameters (prioritized)
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

def check_room_has_3d_zone(room, doc, zone_instances_cache=None):
    """Check if a room already has a 3D Zone instance.
    
    Simplified version: Uses room ID (unique) instead of room number pattern matching.
    Uses pre-filtered cache of only 3DZone instances for maximum performance.
    
    Args:
        room: Room element
        doc: Revit document (unused when cache provided)
        zone_instances_cache: Pre-filtered list of 3DZone FamilyInstance elements.
        
    Returns:
        bool: True if room has a 3D Zone instance, False otherwise
    """
    room_id_str = str(room.Id.IntegerValue)
    
    # Iterate through pre-filtered 3DZone instances only
    for instance in zone_instances_cache or []:
        try:
            family_name = instance.Symbol.Family.Name
            # Check if this 3DZone family matches this room (room ID is unique)
            if room_id_str in family_name:
                return True
        except:
            pass
    
    return False

def extract_room_boundary_loops(room, doc, levels_cache=None):
    """Extract boundary loops from a room.
    
    Args:
        room: Room element
        doc: Revit document
        levels_cache: Optional pre-collected and sorted list of Level elements.
                     If None, will collect them (slower).
    
    Returns:
        tuple: (loops: list of CurveLoop, insertion_point: XYZ, height: float)
    """
    loops = []
    insertion_point = None
    height = 0.0
    
    try:
        # Get boundary segments
        boundary_options = SpatialElementBoundaryOptions()
        boundary_segments = room.GetBoundarySegments(boundary_options)
        
        if not boundary_segments or len(boundary_segments) == 0:
            logger.warning("Room {} has no boundary segments".format(room.Id))
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
        
        # Calculate height
        # Try UnboundedHeight first
        try:
            if hasattr(room, 'UnboundedHeight') and room.UnboundedHeight > 0:
                height = room.UnboundedHeight
            else:
                # Fallback: level to level
                level_id = room.LevelId if hasattr(room, 'LevelId') and room.LevelId else None
                if level_id:
                    level = doc.GetElement(level_id)
                    if level:
                        base_z = level.Elevation
                        
                        # Use cached levels if provided, otherwise collect them
                        if levels_cache is None:
                            # Find level above
                            all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                            levels_sorted = sorted(all_levels, key=lambda l: l.Elevation)
                        else:
                            levels_sorted = levels_cache
                        
                        top_z = base_z + 10.0  # Default 10 feet
                        for lvl in levels_sorted:
                            if lvl.Elevation > base_z:
                                top_z = lvl.Elevation
                                break
                        height = top_z - base_z
        except Exception as e:
            logger.debug("Error calculating height: {}".format(e))
            height = 10.0  # Default fallback
        
    except Exception as e:
        logger.error("Error extracting room boundary loops: {}".format(e))
        import traceback
        logger.error(traceback.format_exc())
    
    return loops, insertion_point, height

# --- Room Filter Dialog ---

class RoomItem(object):
    """Represents a room item in the filter dialog."""
    def __init__(self, room, doc, has_zone=False):
        self.room = room
        self.has_zone = has_zone
        
        # Get room display info
        room_number_param = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        self.room_number = room_number_param.AsString() if room_number_param and room_number_param.HasValue else "?"
        
        room_name_param = room.get_Parameter(BuiltInParameter.ROOM_NAME)
        self.room_name = room_name_param.AsString() if room_name_param and room_name_param.HasValue else "Unnamed"
        
        # Get level name from room's LevelId property
        level_name = "?"
        level_id = room.LevelId if hasattr(room, 'LevelId') and room.LevelId else None
        if level_id:
            level_elem = doc.GetElement(level_id)
            if level_elem:
                level_name = level_elem.Name
        
        # Create display text with icon: (*) if zone exists, empty if not
        icon = "(*)" if has_zone else "   "
        self.display_text = "{} {} - {} ({})".format(icon, self.room_number, self.room_name, level_name)
    
    def __str__(self):
        """String representation for pyRevit forms."""
        return self.display_text
    
    def __repr__(self):
        """Representation for debugging."""
        return self.display_text

def show_room_filter_dialog(rooms, doc):
    """Show a pyRevit native forms dialog to filter and select rooms.
    
    Args:
        rooms: List of Room elements
        doc: Revit document
        
    Returns:
        List of selected Room elements, or None if cancelled
    """
    # #region agent log
    cache_start = time.time()
    debug_log("show_room_filter_dialog:start", "Starting dialog", {"room_count": len(rooms)})
    # #endregion
    
    # OPTIMIZATION: Collect Generic Model instances ONCE and filter to only 3DZone families
    # This avoids collecting instances multiple times AND reduces iteration overhead
    # #region agent log
    collect_start = time.time()
    # #endregion
    
    all_generic_instances = FilteredElementCollector(doc)\
        .OfClass(FamilyInstance)\
        .OfCategory(BuiltInCategory.OST_GenericModel)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # #region agent log
    collect_time = time.time() - collect_start
    debug_log("show_room_filter_dialog:after_collect", "Collected all Generic Model instances", {"instance_count": len(all_generic_instances), "collect_time_ms": collect_time * 1000})
    # #endregion
    
    # Filter to only 3DZone families (significant optimization - reduces iteration from ~1555 to ~21 instances)
    # #region agent log
    filter_start = time.time()
    # #endregion
    
    zone_instances_cache = []
    for instance in all_generic_instances:
        try:
            family_name = instance.Symbol.Family.Name
            if family_name.startswith("3DZone_Room-"):
                zone_instances_cache.append(instance)
        except:
            pass
    
    # #region agent log
    filter_time = time.time() - filter_start
    cache_time = time.time() - cache_start
    debug_log("show_room_filter_dialog:after_filter", "Filtered to 3DZone instances", {"zone_instance_count": len(zone_instances_cache), "filter_time_ms": filter_time * 1000, "total_cache_time_ms": cache_time * 1000})
    # #endregion
    
    # Check for existing zones and create room items
    logger.debug("Checking for existing 3D zones...")
    
    # #region agent log
    check_start = time.time()
    # #endregion
    
    room_items = []
    for room in rooms:
        has_zone = check_room_has_3d_zone(room, doc, zone_instances_cache)
        room_item = RoomItem(room, doc, has_zone)
        room_items.append(room_item)
    
    # #region agent log
    check_time = time.time() - check_start
    total_time = time.time() - cache_start
    debug_log("show_room_filter_dialog:after_checks", "Finished checking rooms", {"room_count": len(rooms), "check_time_ms": check_time * 1000, "total_time_ms": total_time * 1000, "avg_per_room_ms": (check_time * 1000) / len(rooms) if rooms else 0})
    # #endregion
    
    # Count existing zones
    existing_count = sum(1 for item in room_items if item.has_zone)
    logger.debug("Found {} rooms with existing 3D zones".format(existing_count))
    
    # Show selection dialog with search/filter capability
    # forms.SelectFromList supports multiple selection and built-in search
    selected_items = forms.SelectFromList.show(
        room_items,
        title="Select Rooms for 3D Zone Creation",
        multiselect=True,
        button_name='Create Zones',
        width=500,
        height=600
    )
    
    # Return selected rooms (or None if cancelled)
    if selected_items:
        selected_rooms = [item.room for item in selected_items]
        return selected_rooms
    else:
        return None

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
    
    # OPTIMIZATION: Cache families and levels collections once at the start
    # This avoids repeated FilteredElementCollector calls throughout execution
    # Cache families collection
    families_cache = FilteredElementCollector(doc).OfClass(Family).ToElements()
    
    # Cache levels collection (sorted by elevation for height calculations)
    levels_cache = FilteredElementCollector(doc).OfClass(Level).ToElements()
    levels_cache_sorted = sorted(levels_cache, key=lambda l: l.Elevation)
    
    # Show room filter dialog
    rooms_list = show_room_filter_dialog(placed_rooms, doc)
    
    # Check if user cancelled or selected no rooms
    if not rooms_list:
        logger.debug("No rooms selected, exiting.")
        script.exit()
    
    logger.debug("User selected {} rooms for 3D Zone creation".format(len(rooms_list)))
    
    # Check if template family exists in project first
    template_family_name_in_project = "3DZone"  # Name without .rfa extension
    project_template_family = None
    
    # Use cached families collection
    for fam in families_cache:
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
                        # Refresh families cache after loading
                        families_cache = FilteredElementCollector(doc).OfClass(Family).ToElements()
                        project_template_family = next((f for f in families_cache if f.Name == template_family_name_in_project), None)
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
    
    # Process each room
    success_count = 0
    fail_count = 0
    created_families = []
    failed_rooms = []  # List of dicts: {room, room_number_str, room_name_str, reason}
    
    logger.debug("Starting to process {} rooms...".format(len(rooms_list)))
    
    # First pass: Process all family documents (separate transactions per family doc)
    # Store room data for second pass (loading/placing in project)
    room_family_data = []  # List of dicts: {room, output_family_path, insertion_point, room_number_str, room_name_str}
    
    # Combined progress bar for both phases (creating families: 0-50%, loading/placing: 50-100%)
    total_rooms = len(rooms_list)
    with forms.ProgressBar(title="Creating 3D Zone Families and Placing Instances ({} rooms)".format(total_rooms)) as pb:
        # Phase 1: Process rooms and create family documents (0-50% of progress)
        for room_idx, room in enumerate(rooms_list):
            try:
                room_id = room.Id.IntegerValue
                
                # Get room name for file naming
                room_name_param = room.get_Parameter(BuiltInParameter.ROOM_NAME)
                room_name_str = room_name_param.AsString() if room_name_param and room_name_param.HasValue else "Unnamed"
                room_number_param = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                room_number_str = room_number_param.AsString() if room_number_param and room_number_param.HasValue else str(room_id)
                
                logger.debug("Processing room {} ({}/{}): {} - {}".format(
                    room_number_str, room_idx + 1, len(rooms_list), room_name_str, room_id))
                
                # Update progress bar: Phase 1 is 0-50% (room_idx+1 out of total_rooms = 50% max)
                # Progress = (room_idx + 1) / total_rooms * 0.5 * 100
                phase1_progress = int((room_idx + 1) / float(total_rooms) * 50)
                pb.update_progress(phase1_progress, 100)
                
                # Extract room boundary loops (pass cached levels for performance)
                loops, insertion_point, height = extract_room_boundary_loops(room, doc, levels_cache_sorted)
                
                if not loops or not insertion_point or height <= 0:
                    logger.debug("Room {} (ID: {}) - invalid boundary data, skipping".format(
                        room_number_str, room_id))
                    fail_count += 1
                    failed_rooms.append({
                        'room': room,
                        'room_number_str': room_number_str,
                        'room_name_str': room_name_str,
                        'reason': 'Invalid boundary data (no loops, invalid insertion point, or height <= 0)'
                    })
                    continue
                
                logger.debug("Room {}: {} loops, height: {}, insertion point: ({:.2f}, {:.2f}, {:.2f})".format(
                    room_number_str, len(loops), height, insertion_point.X, insertion_point.Y, insertion_point.Z))
                
                # Create output family name and path (using temporary file)
                # Sanitize room number to be safe for family names (no periods, slashes, etc.)
                safe_room_number = sanitize_family_name(room_number_str)
                output_family_name = "3DZone_Room-{}_{}.rfa".format(safe_room_number, room_id)
                output_family_path = op.join(temp_dir, output_family_name)
                family_name_without_ext = op.splitext(output_family_name)[0]  # Name without .rfa
                
                # Check if family already exists in project
                existing_family = None
                
                # Use cached families collection (refresh if new families were loaded)
                # Note: families_cache may need refresh after loading new families, but for checking existing
                # families during processing, the cache should be sufficient
                for fam in families_cache:
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
                        logger.error("Error copying template for room {}: {}".format(room_number_str, copy_error))
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
                                        logger.debug("Room {}: Found None curve in loop {}, skipping".format(room_number_str, loop_idx))
                                        continue
                                    
                                    # Check if curve is valid (has valid endpoints)
                                    try:
                                        start_pt = curve.GetEndPoint(0)
                                        end_pt = curve.GetEndPoint(1)
                                        
                                        # Check for degenerate curves (zero length)
                                        if start_pt.DistanceTo(end_pt) < 0.001:  # Less than 1mm
                                            logger.debug("Room {}: Degenerate curve detected in loop {} (length < 1mm), skipping".format(room_number_str, loop_idx))
                                            continue
                                    except Exception as curve_check_error:
                                        logger.debug("Room {}: Error checking curve in loop {}: {}, skipping".format(room_number_str, loop_idx, curve_check_error))
                                        continue
                                    
                                    # Translate curve to be relative to insertion point AND set Z=0
                                    first_pt = curve.GetEndPoint(0)
                                    # Create translation that moves to origin and sets Z=0
                                    translation = Transform.CreateTranslation(XYZ(-insertion_point.X, -insertion_point.Y, -first_pt.Z))
                                    translated_curve = curve.CreateTransformed(translation)
                                    
                                    # Validate translated curve
                                    if translated_curve is None:
                                        logger.debug("Room {}: Translated curve is None in loop {}, skipping".format(room_number_str, loop_idx))
                                        continue
                                    
                                    translated_loop.Append(translated_curve)
                                    loop_curves.append(translated_curve)
                                except Exception as curve_error:
                                    logger.debug("Room {}: Error processing curve in loop {}: {}, skipping".format(room_number_str, loop_idx, curve_error))
                                    continue
                            
                            # Validate loop has at least 3 curves (minimum for a valid closed loop)
                            if len(loop_curves) < 3:
                                logger.debug("Room {}: Loop {} has only {} curves (minimum 3 required), skipping".format(room_number_str, loop_idx, len(loop_curves)))
                                continue
                            
                            # Check if loop is closed (first point should match last point)
                            try:
                                if translated_loop.IsOpen():
                                    logger.debug("Room {}: Loop {} is open (not closed), attempting to close".format(room_number_str, loop_idx))
                                    # Try to close the loop by adding a line from last to first point
                                    first_curve = loop_curves[0]
                                    last_curve = loop_curves[-1]
                                    first_pt = first_curve.GetEndPoint(0)
                                    last_pt = last_curve.GetEndPoint(1)
                                    
                                    # Only close if gap is small (< 1mm)
                                    gap = first_pt.DistanceTo(last_pt)
                                    if gap > 0.001:
                                        logger.debug("Room {}: Loop {} gap is {} (too large to auto-close)".format(room_number_str, loop_idx, gap))
                                        continue
                            except Exception as loop_check_error:
                                logger.debug("Room {}: Could not check if loop {} is closed: {}".format(room_number_str, loop_idx, loop_check_error))
                            
                            translated_loops.append(translated_loop)
                        
                        # Validate we have at least one valid loop
                        if not translated_loops:
                            logger.debug("Room {}: No valid loops after translation (likely open loop or invalid geometry), skipping".format(room_number_str))
                            fail_count += 1
                            failed_rooms.append({
                                'room': room,
                                'room_number_str': room_number_str,
                                'room_name_str': room_name_str,
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
                                
                                # Validate curve array before creating extrusion
                                if curve_arr_array.Size == 0:
                                    logger.debug("Room {}: CurveArrArray is empty (no valid loops), skipping".format(room_number_str))
                                    t.RollBack()
                                    fail_count += 1
                                    failed_rooms.append({
                                        'room': room,
                                        'room_number_str': room_number_str,
                                        'room_name_str': room_name_str,
                                        'reason': 'CurveArrArray is empty (no valid loops)'
                                    })
                                    if family_doc:
                                        try:
                                            family_doc.Close(False)
                                        except:
                                            pass
                                        family_doc = None
                                    continue
                                
                                # Validate height
                                if height <= 0 or height > 10000:  # Sanity check: 0 to 10000 feet
                                    logger.debug("Room {}: Invalid height: {} (must be > 0 and < 10000), skipping".format(room_number_str, height))
                                    t.RollBack()
                                    fail_count += 1
                                    failed_rooms.append({
                                        'room': room,
                                        'room_number_str': room_number_str,
                                        'room_name_str': room_name_str,
                                        'reason': 'Invalid height: {} (must be > 0 and < 10000)'.format(height)
                                    })
                                    if family_doc:
                                        try:
                                            family_doc.Close(False)
                                        except:
                                            pass
                                        family_doc = None
                                    continue
                                
                                # Create new extrusion using FamilyCreate (not FamilyManager)
                                # Extrusion start = 0, end = height (relative to sketch plane)
                                # Use FamilyCreate.NewExtrusion (not FamilyManager)
                                family_create = family_doc.FamilyCreate
                                
                                # #region agent log
                                debug_log("extrusion:before_create", "Before NewExtrusion", {"room_number": room_number_str, "loop_count": curve_arr_array.Size, "height": height})
                                # #endregion
                                
                                new_extrusion = family_create.NewExtrusion(True, curve_arr_array, sketch_plane, height)
                                
                                # #region agent log
                                debug_log("extrusion:after_create", "After NewExtrusion", {"room_number": room_number_str, "success": new_extrusion is not None})
                                # #endregion
                                
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
                                    
                                    # Re-associate parameters if they exist
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
                                    
                                    if template_info['material_param']:
                                        try:
                                            mat_param = new_extrusion.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                                            if mat_param:
                                                fm.AssociateElementParameterToFamilyParameter(mat_param, template_info['material_param'])
                                                logger.debug("Associated material parameter")
                                        except Exception as assoc_error:
                                            logger.debug("Could not associate material parameter: {}".format(assoc_error))
                                    
                                    # Set extrusion start/end values
                                    if template_info['extrusion_start_param']:
                                        try:
                                            fm.Set(template_info['extrusion_start_param'], 0.0)
                                        except:
                                            pass
                                    
                                    if template_info['extrusion_end_param']:
                                        try:
                                            fm.Set(template_info['extrusion_end_param'], height)
                                        except:
                                            pass
                                else:
                                    logger.debug("Room {}: Failed to create new extrusion (invalid geometry), skipping".format(room_number_str))
                                    t.RollBack()
                                    fail_count += 1
                                    failed_rooms.append({
                                        'room': room,
                                        'room_number_str': room_number_str,
                                        'room_name_str': room_name_str,
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
                                logger.debug("Room {}: Error creating extrusion: {}, skipping".format(room_number_str, recreate_error))
                                fail_count += 1
                                failed_rooms.append({
                                    'room': room,
                                    'room_number_str': room_number_str,
                                    'room_name_str': room_name_str,
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
                        logger.error("Error processing family document for room {}: {}".format(
                            room_number_str, family_error))
                        import traceback
                        logger.error(traceback.format_exc())
                        fail_count += 1
                        failed_rooms.append({
                            'room': room,
                            'room_number_str': room_number_str,
                            'room_name_str': room_name_str,
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
                    
                    # Store room data for second pass (loading/placing in project)
                    room_family_data.append({
                        'room': room,
                        'output_family_path': output_family_path,
                        'insertion_point': insertion_point,
                        'room_number_str': room_number_str,
                        'room_name_str': room_name_str,
                        'output_family_name': output_family_name,
                        'family_name_without_ext': family_name_without_ext,
                        'existing_family': existing_family,  # None if needs to be loaded, Family object if already exists
                        'family_doc': family_doc  # Keep reference to open family doc (will close after loading)
                    })
                else:
                    # Family already exists, skip family document processing
                    logger.debug("Skipping family document processing for room {} - family already exists".format(room_number_str))
                    
                    # Store room data for second pass (loading/placing in project)
                    room_family_data.append({
                        'room': room,
                        'output_family_path': None,  # No file path since family already exists
                        'insertion_point': insertion_point,
                        'room_number_str': room_number_str,
                        'room_name_str': room_name_str,
                        'output_family_name': output_family_name,
                        'family_name_without_ext': family_name_without_ext,
                        'existing_family': existing_family  # Use existing family
                    })

                    
            except Exception as e:
                logger.error("Error processing room {} (ID: {}): {}".format(
                    room_number_str if 'room_number_str' in locals() else "Unknown",
                    room_id if 'room_id' in locals() else "Unknown",
                    str(e)))
                import traceback
                logger.error(traceback.format_exc())
                fail_count += 1
                
                # Make sure family doc is closed
                if 'family_doc' in locals() and family_doc:
                    try:
                        family_doc.Close(False)
                    except:
                        pass
        
        # Second pass: Load all families and place all instances in a single project transaction
        # (Still within the same progress bar - Phase 2: 50-100% of progress)
        logger.debug("Loading families and placing instances in project (single transaction)...")
        
        if room_family_data:
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
            
            # IMPORTANT: familyDoc.LoadFamily(projectDoc, ...) requires the project document
            # to NOT be modifiable (no open transactions). We therefore split Phase 2 into:
            # (A) Load families with no transaction
            # (B) Place instances inside a single transaction

            # A) Load all required families (no transaction open)
            # Progress: 50-75% (first half of phase 2)
            total_rooms = len(room_family_data)
            for room_idx, room_data in enumerate(room_family_data):
                output_family_path = room_data.get('output_family_path')
                existing_family = room_data.get('existing_family')

                # Update progress bar: Loading phase is 50-75% (first half of phase 2)
                # Progress = 50 + (room_idx + 1) / total_rooms * 25
                loading_progress = int(50 + (room_idx + 1) / float(total_rooms) * 25)
                pb.update_progress(loading_progress, 100)

                if existing_family:
                    room_data['loaded_family'] = existing_family
                    continue

                load_result = None
                load_family_doc = room_data.get('family_doc')  # Reuse family doc from phase 1 if available
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
                            # Refresh families cache after loading new family
                            families_cache = FilteredElementCollector(doc).OfClass(Family).ToElements()
                            loaded_family = next((f for f in families_cache if f.Name == family_name), None)
                        else:
                            loaded_family = None
                    else:
                        loaded_family = load_result
                except:
                    loaded_family = None

                room_data['loaded_family'] = loaded_family

            # B) Place all instances in a single transaction
            t = Transaction(doc, "Place 3D Zone Instances")
            t.Start()
            try:
                for room_idx, room_data in enumerate(room_family_data):
                    room = room_data['room']
                    insertion_point = room_data['insertion_point']
                    room_number_str = room_data['room_number_str']
                    output_family_name = room_data['output_family_name']
                    loaded_family = room_data.get('loaded_family') or room_data.get('existing_family')

                    # Update progress bar: Placing phase is 75-100% (second half of phase 2)
                    # Progress = 75 + (room_idx + 1) / total_rooms * 25
                    placing_progress = int(75 + (room_idx + 1) / float(total_rooms) * 25)
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

                        level_id = room.LevelId if hasattr(room, 'LevelId') and room.LevelId else None
                        level = doc.GetElement(level_id) if level_id else None

                        # Get Room's Phase BEFORE instance creation (optimization)
                        # Use ROOM_PHASE parameter to get the Room's phase
                        room_phase_id = None
                        try:
                            room_phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
                            if room_phase_param and room_phase_param.HasValue:
                                room_phase_id = room_phase_param.AsElementId()
                        except Exception:
                            pass  # Room may not have phase, continue without setting phase

                        placement_point = XYZ(
                            insertion_point.X,
                            insertion_point.Y,
                            0.0
                        )

                        instance = doc.Create.NewFamilyInstance(
                            placement_point,
                            symbol,
                            level,
                            StructuralType.NonStructural
                        )

                        if not instance:
                            logger.warning("Failed to place instance for room {}".format(room_number_str))
                            fail_count += 1
                            continue

                        logger.debug("Placed instance for room {}: {}".format(room_number_str, instance.Id))

                        # Set phase IMMEDIATELY after creation (optimized for speed)
                        # Match the instance's phase to the Room's phase
                        if instance and room_phase_id and room_phase_id != ElementId.InvalidElementId:
                            try:
                                instance.CreatedPhaseId = room_phase_id
                            except Exception:
                                pass  # Instance may not support phase setting, continue anyway

                        try:
                            # Copy Name
                            room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME)
                            if room_name and room_name.HasValue:
                                inst_name = instance.LookupParameter("Name")
                                if inst_name and not inst_name.IsReadOnly:
                                    inst_name.Set(room_name.AsString())

                            # Copy Number
                            room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                            if room_number and room_number.HasValue:
                                inst_number = instance.LookupParameter("Number")
                                if inst_number and not inst_number.IsReadOnly:
                                    inst_number.Set(room_number.AsString())

                            # Copy MMI
                            configured_mmi_param = get_mmi_parameter_name(doc)
                            mmi_copied = False
                            if configured_mmi_param:
                                source_mmi = room.LookupParameter(configured_mmi_param)
                                if source_mmi and source_mmi.HasValue:
                                    target_mmi = instance.LookupParameter(configured_mmi_param)
                                    if target_mmi and not target_mmi.IsReadOnly:
                                        if source_mmi.StorageType == StorageType.String:
                                            value = source_mmi.AsString()
                                            if value:
                                                target_mmi.Set(value)
                                                mmi_copied = True

                            if not mmi_copied:
                                source_mmi = room.LookupParameter("MMI")
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

                    except Exception as room_place_error:
                        logger.error("Error placing family for room {}: {}".format(room_number_str, room_place_error))
                        import traceback
                        logger.error(traceback.format_exc())
                        fail_count += 1

                t.Commit()
            except Exception as tx_error:
                t.RollBack()
                logger.error("Transaction failed, rolled back: {}".format(tx_error))
                import traceback
                logger.error(traceback.format_exc())
                fail_count += len(room_family_data) - success_count
    
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
    logger.debug("Created {} 3D Zone families from {} Rooms ({} failed)".format(
        success_count, len(rooms_list), fail_count))
    
    # Show failed rooms with Linkify if any
    if failed_rooms:
        output = script.get_output()
        output.print_md("## ⚠️ Failed Rooms - Invalid Geometry")
        output.print_md("**{} room(s) could not be processed due to invalid geometry.**".format(len(failed_rooms)))
        output.print_md("Click room ID to zoom to room in Revit")
        output.print_md("---")
        
        # Group by reason for better overview
        by_reason = {}
        for item in failed_rooms:
            reason = item['reason']
            if reason not in by_reason:
                by_reason[reason] = []
            by_reason[reason].append(item)
        
        # Print by reason
        for reason, items in sorted(by_reason.items()):
            output.print_md("### {} ({} rooms)".format(reason, len(items)))
            output.print_md("")
            
            # Print each room with clickable link
            for item in items:
                room = item['room']
                room_number = item['room_number_str']
                room_name = item['room_name_str']
                # Use linkify method from output window to create clickable link
                room_link = output.linkify(room.Id)
                output.print_md("- **Room {} - {}**: {}".format(room_number, room_name, room_link))
            
            output.print_md("")  # Empty line between reasons
        
        output.print_md("---")
        output.print_md("**Tip:** Fix the room boundaries in Revit and run the script again.")

