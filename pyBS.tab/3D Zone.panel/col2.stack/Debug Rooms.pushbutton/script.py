# -*- coding: utf-8 -*-
"""Debug tool to visualize Rooms as DirectShape boxes.

Creates DirectShape boxes from Room geometry to help debug
geometry creation issues. Shows which rooms succeed/fail solid creation.
"""

__title__ = "Debug Rooms"
__author__ = "Byggstyrning AB"
__doc__ = "Create DirectShape boxes from Rooms to visualize geometry"

# Import standard libraries
import sys
import os

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Architecture import Room

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Import .NET collections for DirectShape
import clr
from System.Collections.Generic import List
import System

# Add the extension directory to the path
import os.path as op
script_path = __file__
pushbutton_dir = op.dirname(script_path)
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
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
    # Fallback function if import fails
    def get_mmi_parameter_name(doc):
        return "MMI"

# --- Helper Functions ---

def copy_spatial_properties_to_directshape(spatial_element, direct_shape, doc, logger):
    """Copy relevant properties from spatial element (Room/Area/Space) to DirectShape.
    
    Args:
        spatial_element: Room, Area, or Space element
        direct_shape: DirectShape element to copy properties to
        doc: Revit document (needed for MMI parameter lookup)
        logger: Logger instance
        
    Returns:
        int: Number of properties successfully copied
    """
    copied_count = 0
    
    try:
        # List of common properties to copy (parameter names)
        # Try to copy these if they exist on both elements
        properties_to_copy = [
            ("Name", "Name"),
            ("Number", "Number"),
            ("Level", "Level"),
            ("Comments", "Comments"),
            ("Description", "Description"),
        ]
        
        # For Rooms, also try Room-specific parameters
        if isinstance(spatial_element, Room):
            room_specific = [
                ("ROOM_NAME", "Name"),
                ("ROOM_NUMBER", "Number"),
                ("ROOM_LEVEL", "Level"),
                ("ROOM_COMMENTS", "Comments"),
            ]
            properties_to_copy.extend(room_specific)
        
        # Copy each property
        for source_param_name, target_param_name in properties_to_copy:
            try:
                # Get source parameter
                source_param = None
                
                # Try BuiltInParameter first
                if hasattr(BuiltInParameter, source_param_name):
                    built_in_param = getattr(BuiltInParameter, source_param_name)
                    source_param = spatial_element.get_Parameter(built_in_param)
                
                # Fallback to LookupParameter
                if not source_param:
                    source_param = spatial_element.LookupParameter(source_param_name)
                
                if not source_param or not source_param.HasValue:
                    continue
                
                # Get target parameter
                target_param = direct_shape.LookupParameter(target_param_name)
                if not target_param:
                    # Try BuiltInParameter for target
                    if hasattr(BuiltInParameter, target_param_name):
                        built_in_param = getattr(BuiltInParameter, target_param_name)
                        target_param = direct_shape.get_Parameter(built_in_param)
                
                if not target_param or target_param.IsReadOnly:
                    continue
                
                # Check storage types match
                if source_param.StorageType != target_param.StorageType:
                    continue
                
                # Copy the value based on storage type
                storage_type = source_param.StorageType
                if storage_type == StorageType.String:
                    value = source_param.AsString()
                    if value:
                        target_param.Set(value)
                        copied_count += 1
                elif storage_type == StorageType.Integer:
                    value = source_param.AsInteger()
                    target_param.Set(value)
                    copied_count += 1
                elif storage_type == StorageType.Double:
                    value = source_param.AsDouble()
                    target_param.Set(value)
                    copied_count += 1
                elif storage_type == StorageType.ElementId:
                    value = source_param.AsElementId()
                    if value and value != ElementId.InvalidElementId:
                        target_param.Set(value)
                        copied_count += 1
                        
            except Exception as prop_error:
                logger.debug("Error copying property '{}': {}".format(source_param_name, str(prop_error)))
                continue
        
        # Also copy Area as a text parameter if possible
        try:
            if hasattr(spatial_element, 'Area'):
                area_value = spatial_element.Area
                if area_value > 0:
                    # Try to set as a text parameter named "Area" or "Room Area"
                    area_param = direct_shape.LookupParameter("Area")
                    if not area_param:
                        area_param = direct_shape.LookupParameter("Room Area")
                    if area_param and not area_param.IsReadOnly:
                        if area_param.StorageType == StorageType.Double:
                            area_param.Set(area_value)
                            copied_count += 1
                        elif area_param.StorageType == StorageType.String:
                            # Format area as string
                            area_param.Set("{:.2f}".format(area_value))
                            copied_count += 1
        except Exception as area_error:
            logger.debug("Error copying Area property: {}".format(str(area_error)))
        
        # Copy Level name as text if Level parameter doesn't exist
        try:
            if hasattr(spatial_element, 'Level') and spatial_element.Level:
                level = spatial_element.Level
                level_name_param = direct_shape.LookupParameter("Level Name")
                if level_name_param and not level_name_param.IsReadOnly:
                    if level_name_param.StorageType == StorageType.String:
                        level_name_param.Set(level.Name)
                        copied_count += 1
        except Exception as level_error:
            logger.debug("Error copying Level name: {}".format(str(level_error)))
        
        # Copy MMI parameter (configured parameter first, then fallback to "MMI")
        try:
            # Get configured MMI parameter name
            configured_mmi_param = get_mmi_parameter_name(doc)
            
            # Try to copy configured MMI parameter first
            mmi_copied = False
            if configured_mmi_param:
                source_mmi_param = spatial_element.LookupParameter(configured_mmi_param)
                if source_mmi_param and source_mmi_param.HasValue:
                    target_mmi_param = direct_shape.LookupParameter(configured_mmi_param)
                    if target_mmi_param and not target_mmi_param.IsReadOnly:
                        if source_mmi_param.StorageType == target_mmi_param.StorageType:
                            if source_mmi_param.StorageType == StorageType.String:
                                value = source_mmi_param.AsString()
                                if value:
                                    target_mmi_param.Set(value)
                                    copied_count += 1
                                    mmi_copied = True
                                    logger.debug("Copied configured MMI parameter '{}' from spatial element".format(configured_mmi_param))
            
            # Fallback to "MMI" parameter if configured parameter wasn't copied
            if not mmi_copied:
                source_mmi_param = spatial_element.LookupParameter("MMI")
                if source_mmi_param and source_mmi_param.HasValue:
                    target_mmi_param = direct_shape.LookupParameter("MMI")
                    if target_mmi_param and not target_mmi_param.IsReadOnly:
                        if source_mmi_param.StorageType == target_mmi_param.StorageType:
                            if source_mmi_param.StorageType == StorageType.String:
                                value = source_mmi_param.AsString()
                                if value:
                                    target_mmi_param.Set(value)
                                    copied_count += 1
                                    logger.debug("Copied fallback 'MMI' parameter from spatial element")
        except Exception as mmi_error:
            logger.debug("Error copying MMI parameter: {}".format(str(mmi_error)))
        
    except Exception as e:
        logger.debug("Error in copy_spatial_properties_to_directshape: {}".format(str(e)))
    
    return copied_count

# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Get all Rooms
    # Use OfCategory instead of OfClass - Room is not directly filterable via OfClass
    # According to Revit API: Room exists in API but not in native object model
    # Use SpatialElement and filter, or use OfCategory(BuiltInCategory.OST_Rooms)
    spatial_elements = FilteredElementCollector(doc)\
        .OfClass(SpatialElement)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Filter for Room instances only, and only include placed rooms
    # Unplaced rooms have Area = 0 and should be excluded
    all_rooms = [elem for elem in spatial_elements if isinstance(elem, Room)]
    rooms_list = [room for room in all_rooms if room.Area > 0]
    
    unplaced_count = len(all_rooms) - len(rooms_list)
    if unplaced_count > 0:
        logger.debug("Filtered out {} unplaced rooms (Area = 0)".format(unplaced_count))
    
    if not rooms_list:
        if all_rooms:
            forms.alert("Found {} Rooms, but none are placed (all have Area = 0).\n\nPlease place rooms in the model before running this tool.".format(len(all_rooms)), 
                       title="No Placed Rooms", exitscript=True)
        else:
            forms.alert("No Rooms found in the model.", title="No Rooms", exitscript=True)
    
    logger.debug("Found {} placed Rooms ({} total, {} unplaced)".format(len(rooms_list), len(all_rooms), unplaced_count))
    
    # Ask user if they want to create debug shapes
    message = "Found {} placed Rooms".format(len(rooms_list))
    if unplaced_count > 0:
        message += " ({} unplaced rooms excluded)".format(unplaced_count)
    message += ".\n\nCreate DirectShape boxes for visualization?\n\n"
    message += "This will create boxes showing the geometry of each placed Room.\n"
    message += "You can delete them later if needed."
    
    create_shapes = forms.alert(
        message,
        title="Debug Rooms Visualization",
        ok=False,
        yes=True,
        no=True
    )
    
    if not create_shapes:
        script.exit()
    
    # Check if document supports DirectShape (must be a project document, not family)
    if doc.IsFamilyDocument:
        forms.alert("DirectShape elements can only be created in project documents, not family documents.", 
                   title="Invalid Document Type", exitscript=True)
    
    # Get or create a category for our debug shapes
    # Use Generic Models category (most common for DirectShape)
    debug_category = Category.GetCategory(doc, BuiltInCategory.OST_GenericModel)
    if not debug_category:
        forms.alert("Could not get Generic Model category", title="Error", exitscript=True)
    
    # Log category info for debugging
    category_id_value = debug_category.Id.IntegerValue if debug_category.Id else "None"
    logger.debug("Using category: {} (ID: {})".format(debug_category.Name, category_id_value))
    
    # Get or create DirectShapeType for boxes
    direct_shape_type = None
    try:
        # Try to find existing DirectShapeType in Generic Model category
        collector = FilteredElementCollector(doc).OfClass(DirectShapeType)
        for dst in collector:
            # CRITICAL: Only use DirectShapeTypes in Generic Model category
            if dst.Category and dst.Category.Id == debug_category.Id:
                # Get name via parameter (Name property may not be available)
                name_param = dst.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                if not name_param:
                    name_param = dst.LookupParameter("Type Name")
                dst_name = name_param.AsString() if name_param and name_param.HasValue else ""
                
                if dst_name == "3D Zone Debug Box":
                    direct_shape_type = dst
                    logger.debug("Found existing DirectShapeType in Generic Model: {}".format(dst_name))
                    break
        
        # Create new if not found
        if not direct_shape_type:
            logger.debug("Creating new DirectShapeType...")
            with revit.Transaction("Create Debug DirectShapeType"):
                # DirectShapeType.Create signature varies by Revit version
                # Try different parameter orders
                try:
                    # Method 1: (Document, string name, ElementId categoryId) - name first
                    direct_shape_type = DirectShapeType.Create(doc, "3D Zone Debug Box", debug_category.Id)
                    logger.debug("Created DirectShapeType with name first: (doc, name, categoryId)")
                except Exception as e1:
                    logger.debug("Method 1 failed: {}, trying method 2...".format(str(e1)))
                    try:
                        # Method 2: (Document, ElementId categoryId, string name) - categoryId first
                        direct_shape_type = DirectShapeType.Create(doc, debug_category.Id, "3D Zone Debug Box")
                        logger.debug("Created DirectShapeType with categoryId first: (doc, categoryId, name)")
                    except Exception as e2:
                        logger.debug("Method 2 failed: {}, trying method 3...".format(str(e2)))
                        try:
                            # Method 3: Create without name, set via parameter
                            direct_shape_type = DirectShapeType.Create(doc, debug_category.Id)
                            name_param = direct_shape_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                            if name_param and not name_param.IsReadOnly:
                                name_param.Set("3D Zone Debug Box")
                            logger.debug("Created DirectShapeType without name, set via parameter")
                        except Exception as e3:
                            logger.error("All methods failed. Errors: {}, {}, {}".format(str(e1), str(e2), str(e3)))
                            raise Exception("Could not create DirectShapeType. All methods failed. Check logs for details.")
        
        if direct_shape_type:
            # Verify it's in Generic Model category
            if not direct_shape_type.Category or direct_shape_type.Category.Id != debug_category.Id:
                logger.error("DirectShapeType is not in Generic Model category! Category: {}".format(
                    direct_shape_type.Category.Name if direct_shape_type.Category else "None"))
                raise Exception("DirectShapeType must be in Generic Model category for DirectShape elements")
            
            # Get name for logging
            name_param = direct_shape_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            if not name_param:
                name_param = direct_shape_type.LookupParameter("Type Name")
            type_name = name_param.AsString() if name_param and name_param.HasValue else "Unknown"
            logger.debug("Using DirectShapeType: {} (ID: {}, Category: {})".format(
                type_name, direct_shape_type.Id, direct_shape_type.Category.Name if direct_shape_type.Category else "Unknown"))
        else:
            raise Exception("Failed to get or create DirectShapeType")
            
    except Exception as e:
        logger.error("Error creating DirectShapeType: {}".format(str(e)))
        import traceback
        logger.error("Traceback: {}".format(traceback.format_exc()))
        forms.alert("Error creating DirectShapeType: {}\n\nCheck logs for details.".format(str(e)), title="Error", exitscript=True)
    
    # Check if volume computations are enabled (required for SpatialElementGeometryCalculator)
    volume_computations_enabled = False
    try:
        if hasattr(doc.Settings, 'AreaVolumeSettings'):
            volume_computations_enabled = doc.Settings.AreaVolumeSettings.ComputeVolumes
            if not volume_computations_enabled:
                logger.warning("Volume computations are not enabled. Room solid geometry may not be accurate.")
                logger.warning("Enable: Architecture > Room & Area > Area and Volume Computations > Areas and Volumes")
    except Exception as vol_check_error:
        logger.debug("Could not check volume computation settings: {}".format(str(vol_check_error)))
        # Continue anyway - may still work
    
    # Process each room
    success_count = 0
    fail_count = 0
    created_shapes = []
    
    # Create SpatialElementGeometryCalculator for room geometry
    # This calculator computes the 3D solid geometry of spatial elements
    geometry_calculator = SpatialElementGeometryCalculator(doc)
    
    with revit.Transaction("Create Debug Room Boxes"):
        for room in rooms_list:
            try:
                room_id = room.Id.IntegerValue
                
                # Get room name for logging
                room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME)
                room_name_str = room_name.AsString() if room_name and room_name.HasValue else "Unnamed"
                
                # Try to get actual room geometry using SpatialElementGeometryCalculator
                # Room inherits from SpatialElement, so this should work
                room_solid = None
                try:
                    # CalculateSpatialElementGeometry accepts SpatialElement (Room is a SpatialElement)
                    results = geometry_calculator.CalculateSpatialElementGeometry(room)
                    if results:
                        room_solid = results.GetGeometry()
                        if room_solid and hasattr(room_solid, 'Volume') and room_solid.Volume > 0:
                            logger.debug("Got room solid geometry for room {} (ID: {}, Name: {})".format(
                                room.Id, room_id, room_name_str))
                        else:
                            room_solid = None
                            logger.debug("Room {} (ID: {}, Name: {}) geometry has no volume".format(
                                room.Id, room_id, room_name_str))
                except Exception as geom_error:
                    logger.debug("Failed to get room geometry for room {} (ID: {}, Name: {}): {}".format(
                        room.Id, room_id, room_name_str, str(geom_error)))
                    room_solid = None
                
                # Fallback to bounding box if geometry calculation failed
                box_solid = None
                if room_solid:
                    box_solid = room_solid
                    logger.debug("Using actual Room solid geometry for room {} (ID: {})".format(room.Id, room_id))
                else:
                    # Fallback to bounding box
                    logger.debug("Using bounding box fallback for room {} (ID: {})".format(room.Id, room_id))
                    try:
                        bbox = room.get_BoundingBox(None)
                        if not bbox:
                            logger.warning("Room {} (ID: {}, Name: {}) has no bounding box".format(
                                room.Id, room_id, room_name_str))
                            fail_count += 1
                            continue
                        
                        min_pt = bbox.Min
                        max_pt = bbox.Max
                        
                        # Create box using GeometryCreationUtilities
                        width = max_pt.X - min_pt.X
                        height = max_pt.Y - min_pt.Y
                        depth = max_pt.Z - min_pt.Z
                        
                        # Ensure minimum dimensions
                        if width < 0.01:
                            width = 0.01
                        if height < 0.01:
                            height = 0.01
                        if depth < 0.01:
                            depth = 10.0  # Default 10 feet if calculation failed
                        
                        # Create CurveLoop for box base at the correct position
                        # Calculate center point
                        center = (min_pt + max_pt) / 2.0
                        
                        # Create rectangle at the correct position (not at origin)
                        half_width = width / 2.0
                        half_height = height / 2.0
                        p1 = XYZ(center.X - half_width, center.Y - half_height, min_pt.Z)
                        p2 = XYZ(center.X + half_width, center.Y - half_height, min_pt.Z)
                        p3 = XYZ(center.X + half_width, center.Y + half_height, min_pt.Z)
                        p4 = XYZ(center.X - half_width, center.Y + half_height, min_pt.Z)
                        
                        # Create lines for the rectangle
                        line1 = Line.CreateBound(p1, p2)
                        line2 = Line.CreateBound(p2, p3)
                        line3 = Line.CreateBound(p3, p4)
                        line4 = Line.CreateBound(p4, p1)
                        
                        # Create CurveLoop
                        curve_loop = CurveLoop()
                        curve_loop.Append(line1)
                        curve_loop.Append(line2)
                        curve_loop.Append(line3)
                        curve_loop.Append(line4)
                        
                        # Create extrusion in Z direction (vertical)
                        box_solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                            [curve_loop],
                            XYZ(0, 0, 1),
                            depth
                        )
                        
                        if not box_solid:
                            raise Exception("Failed to create extrusion geometry")
                    except Exception as e:
                        logger.warning("Error creating box solid for room {} (ID: {}, Name: {}): {}".format(
                            room.Id, room_id, room_name_str, str(e)))
                        fail_count += 1
                        continue
                
                if not box_solid:
                    logger.warning("Failed to create box solid for room {} (ID: {}, Name: {})".format(
                        room.Id, room_id, room_name_str))
                    fail_count += 1
                    continue
                
                # Create DirectShape
                try:
                    # Create name based on success/failure of actual Room solid creation
                    if room_solid:
                        shape_name = "Room_{}_{}_SUCCESS".format(room_id, room_name_str)
                    else:
                        shape_name = "Room_{}_{}_FAILED".format(room_id, room_name_str)
                    
                    # DirectShape.CreateElement takes categoryId according to Revit API docs
                    # Verify category ID is valid
                    if not debug_category or not debug_category.Id:
                        raise Exception("Invalid category for DirectShape")
                    
                    logger.debug("Creating DirectShape with category ID: {} (Category: {})".format(
                        debug_category.Id, debug_category.Name))
                    
                    direct_shape = DirectShape.CreateElement(doc, debug_category.Id)
                    
                    if not direct_shape:
                        raise Exception("DirectShape.CreateElement returned None")
                    
                    # Set the DirectShapeType
                    try:
                        direct_shape.SetTypeId(direct_shape_type.Id)
                    except Exception as set_type_error:
                        logger.warning("Could not set DirectShapeType: {}".format(str(set_type_error)))
                        # Try alternative method if SetTypeId doesn't work
                        try:
                            type_param = direct_shape.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
                            if type_param and not type_param.IsReadOnly:
                                type_param.Set(direct_shape_type.Id)
                        except:
                            pass  # Type setting is optional - geometry is more important
                    
                    # Set the geometry using SetShape
                    geometry_list = List[GeometryObject]()
                    geometry_list.Add(box_solid)
                    direct_shape.SetShape(geometry_list)
                    
                    # Set the application data (optional, for identification)
                    try:
                        direct_shape.ApplicationDataId = "3D Zone Debug"
                    except:
                        pass  # ApplicationDataId might not be available in all versions
                    
                    if direct_shape:
                        # Set name parameter if available
                        name_param = direct_shape.LookupParameter("Name")
                        if name_param and not name_param.IsReadOnly:
                            name_param.Set(shape_name)
                        
                        # Copy properties from Room to DirectShape
                        try:
                            props_copied = copy_spatial_properties_to_directshape(room, direct_shape, doc, logger)
                            if props_copied > 0:
                                logger.debug("Copied {} properties from Room {} to DirectShape".format(
                                    props_copied, room_id))
                        except Exception as prop_copy_error:
                            logger.debug("Error copying properties: {}".format(str(prop_copy_error)))
                        
                        created_shapes.append(direct_shape.Id)
                        success_count += 1
                        logger.debug("Created DirectShape for room {} (ID: {}, Name: {}): {}".format(
                            room.Id, room_id, room_name_str, "SUCCESS" if room_solid else "FAILED"))
                    else:
                        fail_count += 1
                        logger.warning("Failed to create DirectShape for room {} (ID: {}, Name: {})".format(
                            room.Id, room_id, room_name_str))
                
                except Exception as e:
                    logger.error("Error creating DirectShape for room {} (ID: {}, Name: {}): {}".format(
                        room.Id, room_id, room_name_str, str(e)))
                    fail_count += 1
                    continue
            
            except Exception as e:
                logger.error("Error processing room {} (ID: {}): {}".format(
                    room.Id, room_id if 'room_id' in locals() else "Unknown", str(e)))
                fail_count += 1
                continue
    
    # Log results
    logger.debug("Created {} DirectShape boxes from {} Rooms ({} failed)".format(
        success_count, len(rooms_list), fail_count))

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.

