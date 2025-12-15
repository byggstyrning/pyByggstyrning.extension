# -*- coding: utf-8 -*-
"""Debug tool to visualize Areas as DirectShape boxes.

Creates DirectShape boxes from Area bounding boxes to help debug
geometry creation issues. Shows which areas succeed/fail solid creation.
"""

__title__ = "Debug Areas"
__author__ = "Byggstyrning AB"
__doc__ = "Create DirectShape boxes from Areas to visualize geometry"

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

# Import zone3d libraries
try:
    from zone3d import containment
except ImportError as e:
    logger.error("Failed to import zone3d libraries: {}".format(e))
    forms.alert("Failed to import required libraries. Check logs for details.")
    script.exit()

# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Get all Areas
    areas = FilteredElementCollector(doc)\
        .OfClass(SpatialElement)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Filter for Area instances
    areas_list = [elem for elem in areas if isinstance(elem, Area)]
    
    if not areas_list:
        forms.alert("No Areas found in the model.", title="No Areas", exitscript=True)
    
    logger.info("Found {} Areas".format(len(areas_list)))
    
    # Ask user if they want to delete existing debug shapes first
    delete_existing = forms.alert(
        "Found {} Areas.\n\nCreate DirectShape boxes for visualization?\n\n"
        "This will create boxes showing the bounding boxes of each Area.\n"
        "You can delete them later if needed.".format(len(areas_list)),
        title="Debug Areas Visualization",
        ok=False,
        yes=True,
        no=True
    )
    
    if not delete_existing:
        script.exit()
    
    # Check if document supports DirectShape (must be a project document, not family)
    if doc.IsFamilyDocument:
        forms.alert("DirectShape elements can only be created in project documents, not family documents.", 
                   title="Invalid Document Type", exitscript=True)
    
    # Get or create a category for our debug shapes
    # Use Generic Models category (most common for DirectShape)
    # According to Revit API: https://rvtdocs.com/2025/bee8a24f-704e-44d9-e187-9e031548a6d2
    debug_category = Category.GetCategory(doc, BuiltInCategory.OST_GenericModel)
    if not debug_category:
        forms.alert("Could not get Generic Model category", title="Error", exitscript=True)
    
    # Log category info for debugging
    category_id_value = debug_category.Id.IntegerValue if debug_category.Id else "None"
    logger.info("Using category: {} (ID: {})".format(debug_category.Name, category_id_value))
    
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
                    logger.info("Found existing DirectShapeType in Generic Model: {}".format(dst_name))
                    break
        
        # Create new if not found
        if not direct_shape_type:
            logger.info("Creating new DirectShapeType...")
            with revit.Transaction("Create Debug DirectShapeType"):
                # DirectShapeType.Create signature varies by Revit version
                # Try different parameter orders
                try:
                    # Method 1: (Document, string name, ElementId categoryId) - name first
                    direct_shape_type = DirectShapeType.Create(doc, "3D Zone Debug Box", debug_category.Id)
                    logger.info("Created DirectShapeType with name first: (doc, name, categoryId)")
                except Exception as e1:
                    logger.info("Method 1 failed: {}, trying method 2...".format(str(e1)))
                    try:
                        # Method 2: (Document, ElementId categoryId, string name) - categoryId first
                        direct_shape_type = DirectShapeType.Create(doc, debug_category.Id, "3D Zone Debug Box")
                        logger.info("Created DirectShapeType with categoryId first: (doc, categoryId, name)")
                    except Exception as e2:
                        logger.info("Method 2 failed: {}, trying method 3...".format(str(e2)))
                        try:
                            # Method 3: Create without name, set via parameter
                            direct_shape_type = DirectShapeType.Create(doc, debug_category.Id)
                            name_param = direct_shape_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                            if name_param and not name_param.IsReadOnly:
                                name_param.Set("3D Zone Debug Box")
                            logger.info("Created DirectShapeType without name, set via parameter")
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
            logger.info("Using DirectShapeType: {} (ID: {}, Category: {})".format(
                type_name, direct_shape_type.Id, direct_shape_type.Category.Name if direct_shape_type.Category else "Unknown"))
        else:
            raise Exception("Failed to get or create DirectShapeType")
            
    except Exception as e:
        logger.error("Error creating DirectShapeType: {}".format(str(e)))
        import traceback
        logger.error("Traceback: {}".format(traceback.format_exc()))
        forms.alert("Error creating DirectShapeType: {}\n\nCheck logs for details.".format(str(e)), title="Error", exitscript=True)
    
    # Process each area
    success_count = 0
    fail_count = 0
    created_shapes = []
    
    with revit.Transaction("Create Debug Area Boxes"):
        for area in areas_list:
            try:
                area_id = area.Id.IntegerValue
                
                # Get bounding box - Areas are 2D so we need to calculate from boundaries
                bbox = area.get_BoundingBox(None)
                
                # If no bounding box, calculate from boundary segments
                if not bbox:
                    try:
                        # Get boundary segments
                        boundary_options = SpatialElementBoundaryOptions()
                        boundary_segments = area.GetBoundarySegments(boundary_options)
                        
                        if boundary_segments and len(boundary_segments) > 0:
                            # Collect all points from boundary segments
                            all_points = []
                            for segment_group in boundary_segments:
                                for segment in segment_group:
                                    curve = segment.GetCurve()
                                    if curve:
                                        try:
                                            all_points.append(curve.GetEndPoint(0))
                                            all_points.append(curve.GetEndPoint(1))
                                        except:
                                            pass
                            
                            if all_points:
                                # Calculate bounding box from points
                                min_x = min(pt.X for pt in all_points)
                                max_x = max(pt.X for pt in all_points)
                                min_y = min(pt.Y for pt in all_points)
                                max_y = max(pt.Y for pt in all_points)
                                min_z = min(pt.Z for pt in all_points)
                                max_z = max(pt.Z for pt in all_points)
                                
                                # Get level elevation for Z and calculate full level height
                                level_id = None
                                area_level = None
                                if hasattr(area, "LevelId") and area.LevelId:
                                    level_id = area.LevelId
                                elif hasattr(area, "get_Parameter"):
                                    level_param = area.get_Parameter("Level")
                                    if level_param:
                                        level_id = level_param.AsElementId()
                                
                                if level_id:
                                    area_level = doc.GetElement(level_id)
                                    if area_level:
                                        base_z = area_level.Elevation
                                        min_z = base_z
                                        
                                        # Find level above (next building storey level) for full height
                                        all_levels = FilteredElementCollector(doc)\
                                            .OfClass(Level)\
                                            .ToElements()
                                        
                                        levels_sorted = sorted(all_levels, key=lambda l: l.Elevation)
                                        
                                        # Find level above
                                        top_z = base_z + 10.0  # Default 10 feet if no level above
                                        for level in levels_sorted:
                                            if level.Elevation > base_z:
                                                top_z = level.Elevation
                                                break
                                        
                                        max_z = top_z
                                
                                min_pt = XYZ(min_x, min_y, min_z)
                                max_pt = XYZ(max_x, max_y, max_z)
                                bbox = BoundingBoxXYZ()
                                bbox.Min = min_pt
                                bbox.Max = max_pt
                            else:
                                logger.warning("Area {} (ID: {}) has no valid boundary points".format(area.Id, area_id))
                                fail_count += 1
                                continue
                        else:
                            logger.warning("Area {} (ID: {}) has no boundary segments".format(area.Id, area_id))
                            fail_count += 1
                            continue
                    except Exception as bbox_error:
                        logger.warning("Area {} (ID: {}) error calculating bounding box: {}".format(area.Id, area_id, str(bbox_error)))
                        fail_count += 1
                        continue
                
                # Create actual Area solid geometry (same as used for containment detection)
                # This uses the real Area boundary, not just a bounding box
                try:
                    area_solid = containment._create_solid_from_area(area, doc)
                except Exception as solid_error:
                    logger.debug("Failed to create Area solid: {}".format(str(solid_error)))
                    area_solid = None
                
                # Use the actual Area solid if it was created successfully
                # Otherwise fall back to bounding box
                box_solid = None
                if area_solid:
                    box_solid = area_solid
                    logger.debug("Using actual Area solid geometry for area {} (ID: {})".format(area.Id, area_id))
                else:
                    # Fallback to bounding box if solid creation failed
                    logger.debug("Using bounding box fallback for area {} (ID: {})".format(area.Id, area_id))
                    try:
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
                        # Depth should already be full level height, but ensure minimum
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
                        # Extrusion direction is vertical (0, 0, 1)
                        box_solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                            [curve_loop],
                            XYZ(0, 0, 1),
                            depth
                        )
                        
                        if not box_solid:
                            raise Exception("Failed to create extrusion geometry")
                    except Exception as e:
                        logger.warning("Error creating box solid for area {} (ID: {}): {}".format(
                            area.Id, area_id, str(e)))
                        fail_count += 1
                        continue
                
                if not box_solid:
                    logger.warning("Failed to create box solid for area {} (ID: {})".format(
                        area.Id, area_id))
                    fail_count += 1
                    continue
                
                # Create DirectShape
                try:
                    # Create name based on success/failure of actual Area solid creation
                    if area_solid:
                        shape_name = "Area_{}_SUCCESS".format(area_id)
                    else:
                        shape_name = "Area_{}_FAILED".format(area_id)
                    
                    # DirectShape.CreateElement takes categoryId according to Revit API docs
                    # https://rvtdocs.com/2025/bee8a24f-704e-44d9-e187-9e031548a6d2
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
                        
                        # Set color based on success/failure
                        # Green for success, Red for failure
                        try:
                            override_options = OverrideGraphicSettings()
                            if area_solid:
                                # Green for successful solid creation
                                override_options.SetProjectionLineColor(Color(0, 255, 0))
                            else:
                                # Red for failed solid creation
                                override_options.SetProjectionLineColor(Color(255, 0, 0))
                            
                            # Apply override to active view if available
                            if doc.ActiveView:
                                doc.ActiveView.SetElementOverrides(direct_shape.Id, override_options)
                        except:
                            pass  # Override is optional
                        
                        created_shapes.append(direct_shape.Id)
                        success_count += 1
                        logger.info("Created DirectShape for area {} (ID: {}): {}".format(
                            area.Id, area_id, "SUCCESS" if area_solid else "FAILED"))
                    else:
                        fail_count += 1
                        logger.warning("Failed to create DirectShape for area {} (ID: {})".format(
                            area.Id, area_id))
                
                except Exception as e:
                    logger.error("Error creating DirectShape for area {} (ID: {}): {}".format(
                        area.Id, area_id, str(e)))
                    fail_count += 1
                    continue
            
            except Exception as e:
                logger.error("Error processing area {} (ID: {}): {}".format(
                    area.Id, area_id, str(e)))
                fail_count += 1
                continue
    
    # Show results
    results_text = "Debug Area Visualization Complete\n\n"
    results_text += "Total Areas: {}\n".format(len(areas_list))
    results_text += "DirectShapes Created: {}\n".format(success_count)
    results_text += "Failed: {}\n\n".format(fail_count)
    results_text += "Created {} DirectShape boxes.\n\n".format(len(created_shapes))
    results_text += "Green boxes = Areas with successful solid creation\n"
    results_text += "Red boxes = Areas with failed solid creation\n\n"
    results_text += "You can delete these shapes later if needed."
    
    forms.alert(results_text, title="Debug Areas Complete")
    
    logger.info("Created {} DirectShape boxes from {} Areas".format(success_count, len(areas_list)))

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.

