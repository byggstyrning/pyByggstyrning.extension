# -*- coding: utf-8 -*-
__title__ = "Add to View"
__author__ = "PyRevit Extensions"
__doc__ = """This tool places a workplane-based family at the location and extent of the active view.
The family instance parameters "View Height" and "View Width" are set based on the view's bounding box.
"""

# Import .NET libraries
import clr
clr.AddReference("System")
from System.Collections.Generic import List

# Import Revit API
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType

# Import pyRevit libraries
from pyrevit import revit, script, forms

# Get the current script directory
logger = script.get_logger()

# Get Revit document, app and UIDocument
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_family_symbol():
    """Get the family symbol for 3D View Reference."""
    logger.debug("Searching for 3D View Reference family")
    collectors = FilteredElementCollector(doc).OfClass(FamilySymbol)
    
    # Look for the 3D View Reference face-based family
    for symbol in collectors:
        try:
            if not hasattr(symbol, "Family") or symbol.Family is None:
                continue
            
            family_name = symbol.Family.Name
            
            # Look for 3D View Reference family
            if family_name == "3D View Reference":
                return symbol
        except Exception as ex:
            logger.debug("Error checking family: {}".format(ex))
    
    # If we get here, we didn't find the family
    logger.warning("3D View Reference family not found")
    return None

def get_view_bounding_box(view):
    """Get the bounding box of the view."""
    # Get the crop box of the view if available
    if hasattr(view, "CropBox") and view.CropBoxActive:
        logger.debug("Using crop box for view: {}".format(view.Name))
        bbox = view.CropBox
        logger.debug("Crop box: Min({:.2f}, {:.2f}, {:.2f}), Max({:.2f}, {:.2f}, {:.2f})".format(
            bbox.Min.X, bbox.Min.Y, bbox.Min.Z, bbox.Max.X, bbox.Max.Y, bbox.Max.Z))
        return bbox
    
    # If no crop box, try to get bounding box
    try:
        bbox = view.get_BoundingBox(view)
        if bbox:
            logger.debug("Got bounding box: Min({:.2f}, {:.2f}, {:.2f}), Max({:.2f}, {:.2f}, {:.2f})".format(
                bbox.Min.X, bbox.Min.Y, bbox.Min.Z, bbox.Max.X, bbox.Max.Y, bbox.Max.Z))
            return bbox
    except Exception as ex:
        logger.debug("Could not get bounding box: {}".format(ex))
    
    # If no bounding box, use a default one
    logger.warning("Using default bounding box for view: {}".format(view.Name))
    bbox = BoundingBoxXYZ()
    bbox.Min = XYZ(-5, -5, 0)
    bbox.Max = XYZ(5, 5, 0)
    return bbox

def get_view_center_point(view, bbox):
    """Get the actual center point of the view in model space."""
    try:
        # Get view direction for consistent placement
        view_dir = XYZ(0, 0, 1)  # Default direction
        if hasattr(view, "ViewDirection") and view.ViewDirection is not None:
            view_dir = view.ViewDirection
            logger.debug("View direction: {}".format(view_dir))
        
        # Check if the view is a callout
        is_callout = False
        if hasattr(view, "IsCallout") and view.IsCallout:
            is_callout = True
            logger.debug("View is a callout - will use cut plane for Z coordinate")
        
        # Get crop box in model coordinates
        if hasattr(view, "CropBox") and view.CropBoxActive:
            crop_box = view.CropBox
            transform = view.CropBox.Transform
            
            # Transform to model coordinates
            min_point_transformed = transform.OfPoint(crop_box.Min)
            max_point_transformed = transform.OfPoint(crop_box.Max)
            
            # Calculate center
            center_x = (min_point_transformed.X + max_point_transformed.X) / 2
            center_y = (min_point_transformed.Y + max_point_transformed.Y) / 2
            center_z = (min_point_transformed.Z + max_point_transformed.Z) / 2
            
            # For callouts, use the view's cut plane elevation for Z
            if is_callout:
                try:
                    # Try to get from view range
                    view_range = view.GetViewRange()
                    if view_range:
                        # Try to get the cut plane elevation
                        cut_plane_param = view_range.GetOffset(PlanViewPlane.CutPlane)
                        if cut_plane_param != None:
                            cut_plane_z = cut_plane_param
                            logger.debug("Using callout cut plane for Z: {:.2f}".format(cut_plane_z))
                            center_z = cut_plane_z
                except Exception as ex:
                    logger.debug("Could not get callout cut plane elevation: {}".format(ex))
            
            center = XYZ(center_x, center_y, center_z)
            
            # Add offsets based on view direction
            if abs(view_dir.X) > 0.7:  # East-West facing view
                x_offset = 5 if view_dir.X > 0 else -5  # East/West offset
                center = XYZ(center.X + x_offset, center.Y, center.Z)
            elif abs(view_dir.Y) > 0.7:  # North-South facing view
                y_offset = 5 if view_dir.Y > 0 else -5  # North/South offset
                center = XYZ(center.X, center.Y + y_offset, center.Z)
            
            logger.debug("Center point: ({:.2f}, {:.2f}, {:.2f})".format(
                center.X, center.Y, center.Z))
            return center
        
        # If crop box not available, try section box
        if hasattr(view, "GetSectionBox"):
            try:
                section_box = view.GetSectionBox()
                if section_box:
                    center = XYZ(
                        (section_box.Min.X + section_box.Max.X) / 2, 
                        (section_box.Min.Y + section_box.Max.Y) / 2, 
                        (section_box.Min.Z + section_box.Max.Z) / 2
                    )
                    logger.debug("Using section box center: ({:.2f}, {:.2f}, {:.2f})".format(
                        center.X, center.Y, center.Z))
                    return center
            except Exception as ex:
                logger.debug("Could not get section box: {}".format(ex))
        
        # Try to get the origin directly from the view
        if hasattr(view, "Origin") and view.Origin:
            origin = view.Origin
            logger.debug("Using view's Origin property: ({:.2f}, {:.2f}, {:.2f})".format(
                origin.X, origin.Y, origin.Z))
            return origin
        
        # Final fallback - bounding box
        logger.debug("Using default center point calculation from bbox")
        return XYZ((bbox.Min.X + bbox.Max.X) / 2, (bbox.Min.Y + bbox.Max.Y) / 2, 0)
    
    except Exception as ex:
        logger.warning("Error calculating view center point: {}".format(ex))
        # Ultimate fallback
        return XYZ((bbox.Min.X + bbox.Max.X) / 2, (bbox.Min.Y + bbox.Max.Y) / 2, 0)

def set_instance_parameters(instance, width, height, view_name):
    """Set instance parameters for width and height."""
    try:
        # Set View Width parameter
        width_param = instance.LookupParameter("View Width")
        if width_param:
            width_param.Set(width)
            logger.debug("Set View Width parameter to {:.2f}".format(width))
        else:
            logger.warning("View Width parameter not found")
        
        # Set View Height parameter
        height_param = instance.LookupParameter("View Height")
        if height_param:
            height_param.Set(height)
            logger.debug("Set View Height parameter to {:.2f}".format(height))
        else:
            logger.warning("View Height parameter not found")
        
        # Set View Name parameter
        name_param = instance.LookupParameter("View Name")
        if name_param:
            safe_view_name = view_name if view_name else "Unnamed View"
            name_param.Set(safe_view_name)
            logger.debug("Set View Name parameter to '{}'".format(safe_view_name))
        else:
            logger.warning("View Name parameter not found")
    
    except Exception as ex:
        logger.error("Error setting parameters: {}".format(ex))

def create_3d_view_reference():
    """Create a 3D view reference for the active view."""
    # Get the active view
    active_view = uidoc.ActiveView
    
    # Get family symbol
    family_symbol = get_family_symbol()
    if not family_symbol:
        script.exit("Required family '3D View Reference - Face Based' not found. Please load this family into your project.")
    
    # Make sure family symbol is activated
    try:
        if not family_symbol.IsActive:
            with revit.Transaction("Activate Family Symbol"):
                family_symbol.Activate()
                logger.debug("Activated family symbol")
    except Exception as ex:
        logger.debug("Error activating family symbol: {}".format(ex))
    
    # Get view bounding box
    bbox = get_view_bounding_box(active_view)
    if not bbox:
        script.exit("Could not get bounding box for view: {}".format(active_view.Name))
    
    # Calculate width and height
    width = bbox.Max.X - bbox.Min.X
    height = bbox.Max.Y - bbox.Min.Y
    logger.debug("View dimensions - Width: {:.2f}, Height: {:.2f}".format(width, height))
    
    # Get the center point
    center_point = get_view_center_point(active_view, bbox)
    logger.debug("Center point: ({:.2f}, {:.2f}, {:.2f})".format(
        center_point.X, center_point.Y, center_point.Z))
    
    # Get view direction
    view_dir = XYZ(0, 0, 1)  # Default direction
    try:
        if hasattr(active_view, "ViewDirection") and active_view.ViewDirection is not None:
            view_dir = active_view.ViewDirection
            logger.debug("View direction: {}".format(view_dir))
    except Exception as ex:
        logger.warning("Error getting view direction: {}".format(ex))
    
    # Start transaction
    with revit.Transaction("Create 3D View Reference"):
        # Create a sketch plane aligned with the view
        normal_vector = None
        if abs(view_dir.Z) < 0.99:  # Not looking directly up/down
            up_direction = XYZ(0, 0, 1)
            normal_vector = view_dir.CrossProduct(up_direction).Normalize()
        else:
            # For plan views (looking up/down), use X axis as normal
            normal_vector = XYZ(1, 0, 0)
        
        logger.debug("View normal: ({:.6f}, {:.6f}, {:.6f})".format(
            view_dir.X, view_dir.Y, view_dir.Z))
        logger.debug("Normal vector: ({:.6f}, {:.6f}, {:.6f})".format(
            normal_vector.X, normal_vector.Y, normal_vector.Z))
        
        # Create instance
        new_instance = None
        try:
            # Create a sketch plane aligned with the view
            plane = Plane.CreateByNormalAndOrigin(view_dir, center_point)
            sketch_plane = SketchPlane.Create(doc, plane)
            
            new_instance = doc.Create.NewFamilyInstance(
                center_point, 
                family_symbol, 
                sketch_plane,
                StructuralType.NonStructural
            )
            logger.debug("Successfully placed family instance using sketch plane method")
        except Exception as ex:
            logger.error("Error creating family instance: {}".format(ex))
            script.exit("Failed to create 3D View Reference: {}".format(ex))
        
        # Check if dimensions need to be swapped
        swap_dimensions = False
        if abs(view_dir.Y) > abs(view_dir.X) and abs(view_dir.Y) > abs(view_dir.Z):
            logger.debug("North/South orientation detected - swapping width and height")
            swap_dimensions = True
        
        # Apply the swap if needed
        if swap_dimensions:
            logger.debug("Swapping dimensions: original width={:.2f}, height={:.2f}".format(width, height))
            temp = width
            width = height
            height = temp
            logger.debug("After swap: width={:.2f}, height={:.2f}".format(width, height))
        
        # Set parameters
        if new_instance:
            set_instance_parameters(new_instance, width, height, active_view.Name)
            
            # Select the element instead of isolating it
            try:
                element_ids = List[ElementId]()
                element_ids.Add(new_instance.Id)
                
                # Set selection to the new element
                uidoc.Selection.SetElementIds(element_ids)
                logger.debug("Selected newly created element")
            except Exception as ex:
                logger.debug("Error selecting element: {}".format(ex))
        
            return new_instance.Id
    
    return None

if __name__ == '__main__':
    try:
        element_id = create_3d_view_reference()
        if element_id:
            logger.debug("3D View Reference created successfully for: {}".format(uidoc.ActiveView.Name))
        else:
            logger.error("Failed to create 3D View Reference.")
    except Exception as ex:
        logger.error("Error: {}".format(ex))

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.