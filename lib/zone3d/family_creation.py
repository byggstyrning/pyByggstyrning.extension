# -*- coding: utf-8 -*-
"""Shared family creation functions for 3D Zone creation."""

import os.path as op
import os
import shutil
from Autodesk.Revit.DB import (
    FilteredElementCollector, Extrusion, BuiltInParameter, Category,
    BuiltInCategory, ElementId, Transaction, TransactionStatus, FailureProcessingResult,
    IFailuresPreprocessor, CurveLoop, CurveArray, CurveArrArray,
    XYZ, Transform, Plane, SketchPlane, SaveAsOptions, FamilyInstance, Family
)
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import script

logger = script.get_logger()


def get_template_family_path(extension_dir, template_family_name="3DZone.rfa"):
    """Get the path to the template family.
    
    Args:
        extension_dir: Extension root directory
        template_family_name: Name of template family file
        
    Returns:
        str: Template file path or None if not found
    """
    template_family_path = op.join(extension_dir, template_family_name)
    
    # Log for debugging
    logger.debug("Looking for template at: {}".format(template_family_path))
    logger.debug("Extension dir: {}".format(extension_dir))
    logger.debug("Template exists: {}".format(op.exists(template_family_path)))
    
    if not op.exists(template_family_path):
        logger.error("Template family not found at: {}".format(template_family_path))
        # Try alternative path - maybe it's in the extension root differently
        alt_path = op.join(op.dirname(extension_dir), template_family_name)
        logger.debug("Trying alternative path: {}".format(alt_path))
        if op.exists(alt_path):
            logger.debug("Found template at alternative path: {}".format(alt_path))
            return alt_path
        return None
    return template_family_path


def inspect_template_family(family_doc):
    """Inspect the template family to find extrusion and parameter names.
    
    Args:
        family_doc: Family document
        
    Returns:
        dict: {
            'extrusion': Extrusion element,
            'extrusion_start_param': FamilyParameter for start,
            'extrusion_end_param': FamilyParameter for end,
            'top_offset_param': FamilyParameter for top offset,
            'bottom_offset_param': FamilyParameter for bottom offset,
            'material_param': FamilyParameter for material,
            'subcategory': Category for '3D Zone'
        }
    """
    result = {
        'extrusion': None,
        'extrusion_start_param': None,
        'extrusion_end_param': None,
        'top_offset_param': None,
        'bottom_offset_param': None,
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
        param_names_to_try = [
            "NVExtrusionStart", "NVExtrusionEnd",
            "ExtrusionStart", "ExtrusionEnd",
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
            except Exception:
                pass
        
        # Find Top Offset and Bottom Offset parameters
        top_offset_names = ["Top Offset", "TopOffset", "Top Offset (default)", "Top_Offset"]
        for param_name in top_offset_names:
            try:
                param = fm.get_Parameter(param_name)
                if param:
                    result['top_offset_param'] = param
                    logger.debug("Found Top Offset parameter: {}".format(param_name))
                    break
            except Exception:
                pass
        
        bottom_offset_names = ["Bottom Offset", "BottomOffset", "Bottom Offset (default)", "Bottom_Offset"]
        for param_name in bottom_offset_names:
            try:
                param = fm.get_Parameter(param_name)
                if param:
                    result['bottom_offset_param'] = param
                    logger.debug("Found Bottom Offset parameter: {}".format(param_name))
                    break
            except Exception:
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
            except Exception:
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


def extract_boundary_loops(spatial_element, doc, adapter, levels_cache=None, link_transform=None):
    """Extract boundary loops from a spatial element.
    
    Args:
        spatial_element: Area or Room element
        doc: Source document containing the element (for level lookups/height calc)
        adapter: SpatialElementAdapter instance
        levels_cache: Optional pre-collected and sorted list of Level elements
        link_transform: Optional Transform from linked model to host coordinates.
                       If provided, all boundary curves are transformed to host coords.
        
    Returns:
        tuple: (loops: list of CurveLoop, insertion_point: XYZ, height: float or None)
    """
    loops = []
    insertion_point = None
    height = None
    
    try:
        # Get boundary segments using adapter
        boundary_segments = adapter.get_boundary_segments(spatial_element)
        
        if not boundary_segments or len(boundary_segments) == 0:
            logger.warning("Element {} has no boundary segments".format(spatial_element.Id))
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
                        # Transform curve to host coordinates if link_transform provided
                        if link_transform is not None:
                            curve = curve.CreateTransformed(link_transform)
                        
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
            
            # Check if loop has curves
            try:
                curve_list = list(curve_loop)
                if len(curve_list) > 0:
                    loops.append(curve_loop)
            except:
                pass
        
        # Calculate insertion point (centroid of footprint)
        # Points are already in host coordinates if link_transform was applied
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
        
        # Calculate height using adapter (uses source doc's levels)
        height = adapter.calculate_height(spatial_element, doc, levels_cache)
        
    except Exception as e:
        logger.error("Error extracting boundary loops: {}".format(e))
        import traceback
        logger.error(traceback.format_exc())
    
    return loops, insertion_point, height


def create_zone_family(spatial_element, adapter, template_info, template_path, output_path, 
                       temp_dir, doc, app, element_number_str, element_name_str,
                       source_doc=None, link_transform=None):
    """Create a zone family from a spatial element.
    
    Args:
        spatial_element: Area or Room element
        adapter: SpatialElementAdapter instance
        template_info: Dict from inspect_template_family()
        template_path: Path to template family file
        output_path: Path to save output family file
        temp_dir: Temporary directory path
        doc: Host Revit document
        app: Revit application
        element_number_str: Element number string (for logging)
        element_name_str: Element name string (for logging)
        source_doc: Source document containing the element (defaults to doc)
        link_transform: Transform from linked model to host coordinates (None for active model)
        
    Returns:
        tuple: (success: bool, family_doc: Document or None, error_reason: str or None)
    """
    # Default source_doc to host doc if not provided
    if source_doc is None:
        source_doc = doc
    family_doc = None
    
    try:
        # Copy template to output location
        shutil.copy2(template_path, output_path)
        logger.debug("Copied template to: {}".format(output_path))
        
        # Open the copied family document
        family_doc = app.OpenDocumentFile(output_path)
        if not family_doc:
            return False, None, "Failed to open family document"
        
        # Get the extrusion
        extrusions = FilteredElementCollector(family_doc).OfClass(Extrusion).ToElements()
        if not extrusions:
            return False, family_doc, "No extrusion found in copied family"
        
        extrusion = extrusions[0]
        logger.debug("Found extrusion in copied family: ID {}".format(extrusion.Id))
        
        # Get family manager
        fm = family_doc.FamilyManager
        
        # Extract boundary loops (use source_doc for level lookups, link_transform for coords)
        loops, insertion_point, height = extract_boundary_loops(
            spatial_element, source_doc, adapter, link_transform=link_transform)
        
        if not loops or not insertion_point:
            return False, family_doc, "Invalid boundary data (no loops or invalid insertion point)"
        
        # Translate loops to be relative to insertion point (center near origin)
        # Also translate to Z=0 to ensure sketch plane is perfectly horizontal
        translated_loops = []
        for loop_idx, loop in enumerate(loops):
            translated_loop = CurveLoop()
            loop_curves = []
            
            for curve in loop:
                try:
                    # Validate curve before translation
                    if curve is None:
                        logger.debug("Element {}: Found None curve in loop {}, skipping".format(element_number_str, loop_idx))
                        continue
                    
                    # Check if curve is valid
                    try:
                        start_pt = curve.GetEndPoint(0)
                        end_pt = curve.GetEndPoint(1)
                        
                        # Check for degenerate curves
                        if start_pt.DistanceTo(end_pt) < 0.001:
                            logger.debug("Element {}: Degenerate curve detected in loop {} (length < 1mm), skipping".format(element_number_str, loop_idx))
                            continue
                    except Exception as curve_check_error:
                        logger.debug("Element {}: Error checking curve in loop {}: {}, skipping".format(element_number_str, loop_idx, curve_check_error))
                        continue
                    
                    # Translate curve to be relative to insertion point AND set Z=0
                    first_pt = curve.GetEndPoint(0)
                    translation = Transform.CreateTranslation(XYZ(-insertion_point.X, -insertion_point.Y, -first_pt.Z))
                    translated_curve = curve.CreateTransformed(translation)
                    
                    # Validate translated curve
                    if translated_curve is None:
                        logger.debug("Element {}: Translated curve is None in loop {}, skipping".format(element_number_str, loop_idx))
                        continue
                    
                    translated_loop.Append(translated_curve)
                    loop_curves.append(translated_curve)
                except Exception as curve_error:
                    logger.debug("Element {}: Error processing curve in loop {}: {}, skipping".format(element_number_str, loop_idx, curve_error))
                    continue
            
            # Validate loop has at least 3 curves
            if len(loop_curves) < 3:
                logger.debug("Element {}: Loop {} has only {} curves (minimum 3 required), skipping".format(element_number_str, loop_idx, len(loop_curves)))
                continue
            
            # Check if loop is closed
            try:
                if translated_loop.IsOpen():
                    logger.debug("Element {}: Loop {} is open (not closed), attempting to close".format(element_number_str, loop_idx))
                    first_curve = loop_curves[0]
                    last_curve = loop_curves[-1]
                    first_pt = first_curve.GetEndPoint(0)
                    last_pt = last_curve.GetEndPoint(1)
                    
                    gap = first_pt.DistanceTo(last_pt)
                    if gap > 0.001:
                        logger.debug("Element {}: Loop {} gap is {} (too large to auto-close)".format(element_number_str, loop_idx, gap))
                        continue
            except Exception as loop_check_error:
                logger.debug("Element {}: Could not check if loop {} is closed: {}".format(element_number_str, loop_idx, loop_check_error))
            
            translated_loops.append(translated_loop)
        
        # Validate we have at least one valid loop
        if not translated_loops:
            return False, family_doc, "No valid loops after translation (open loop or invalid geometry)"
        
        # Delete and recreate extrusion
        logger.debug("Recreating extrusion...")
        
        # Suppress "off axis" warnings
        class OffAxisFailurePreprocessor(IFailuresPreprocessor):
            def PreprocessFailures(self, failuresAccessor):
                failures = failuresAccessor.GetFailureMessages()
                for failure in failures:
                    if failure.GetSeverity().ToString() == "Warning":
                        failuresAccessor.DeleteWarning(failure)
                return FailureProcessingResult.Continue
        
        t = Transaction(family_doc, "Recreate Extrusion")
        t.Start()
        failure_options = t.GetFailureHandlingOptions()
        failure_options.SetFailuresPreprocessor(OffAxisFailurePreprocessor())
        t.SetFailureHandlingOptions(failure_options)
        
        try:
            # Delete old extrusion
            family_doc.Delete(extrusion.Id)
            
            # Create sketch plane at Z=0
            origin = XYZ(0, 0, 0)
            normal = XYZ.BasisZ
            plane = Plane.CreateByNormalAndOrigin(normal, origin)
            sketch_plane = SketchPlane.Create(family_doc, plane)
            
            # Create CurveArrArray from loops
            curve_arr_array = CurveArrArray()
            for loop in translated_loops:
                curve_arr = CurveArray()
                for curve in loop:
                    curve_arr.Append(curve)
                curve_arr_array.Append(curve_arr)
            
            # Validate curve array
            if curve_arr_array.Size == 0:
                t.RollBack()
                return False, family_doc, "CurveArrArray is empty (no valid loops)"
            
            # Validate height
            if height is not None and (height <= 0 or height > 10000):
                t.RollBack()
                return False, family_doc, "Invalid height: {} (must be > 0 and < 10000)".format(height)
            
            # Create new extrusion
            family_create = family_doc.FamilyCreate
            extrusion_height = height if height is not None else 10.0  # Temporary default
            
            new_extrusion = family_create.NewExtrusion(True, curve_arr_array, sketch_plane, extrusion_height)
            
            if not new_extrusion:
                t.RollBack()
                return False, family_doc, "Failed to create new extrusion (invalid geometry)"
            
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
            
            # Set family parameter values FIRST (before association)
            if template_info['extrusion_start_param']:
                try:
                    fm.Set(template_info['extrusion_start_param'], 0.0)
                except Exception as pre_set_error:
                    logger.debug("Could not pre-set start family parameter: {}".format(pre_set_error))
            
            if template_info['extrusion_end_param'] and height is not None:
                try:
                    fm.Set(template_info['extrusion_end_param'], height)
                except Exception as pre_set_error:
                    logger.debug("Could not pre-set end family parameter: {}".format(pre_set_error))
            
            # Regenerate to commit family parameter values
            family_doc.Regenerate()
            
            # Set element parameters directly (they're writable before association)
            if template_info['extrusion_start_param']:
                try:
                    start_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_START_PARAM)
                    if start_param and not start_param.IsReadOnly:
                        start_param.Set(0.0)
                        logger.debug("Set extrusion start parameter directly on element: 0.0")
                except Exception as set_error:
                    logger.debug("Could not set start parameter directly: {}".format(set_error))
            
            if template_info['extrusion_end_param'] and height is not None:
                try:
                    end_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                    if end_param and not end_param.IsReadOnly:
                        end_param.Set(height)
                        logger.debug("Set extrusion end parameter directly on element: {}".format(height))
                except Exception as set_error:
                    logger.debug("Could not set end parameter directly: {}".format(set_error))
            
            # Regenerate after setting element parameters
            family_doc.Regenerate()
            
            # Set Top Offset and Bottom Offset (user-facing parameters)
            height_set_in_family = False
            if height is not None:
                # Set Bottom Offset first
                if template_info['bottom_offset_param']:
                    try:
                        fm.Set(template_info['bottom_offset_param'], 0.0)
                        logger.debug("Set bottom offset parameter: 0.0")
                    except Exception as bottom_error:
                        logger.debug("Could not set bottom offset parameter: {}".format(bottom_error))
                
                # Set Top Offset
                if template_info['top_offset_param']:
                    try:
                        fm.Set(template_info['top_offset_param'], height)
                        logger.debug("Set top offset parameter: {}".format(height))
                        height_set_in_family = True
                    except Exception as top_error:
                        logger.debug("Could not set top offset parameter: {}".format(top_error))
                else:
                    logger.debug("Top Offset parameter not found - cannot set height, will use family default")
                    height_set_in_family = False
            
            # Regenerate after setting family parameters
            family_doc.Regenerate()
            
            # Associate ExtrusionEnd/ExtrusionStart to the extrusion element
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
                logger.debug("Element {}: Height is None, using family default Top Offset".format(element_number_str))
            
            # Associate material parameter
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
                except Exception:
                    pass
            
            if template_info['extrusion_end_param'] and height is not None:
                param_set_success = False
                try:
                    # Method 1: Try setting via FamilyManager
                    try:
                        fm.Set(template_info['extrusion_end_param'], height)
                        family_doc.Regenerate()
                        param_set_success = True
                    except Exception:
                        pass
                    
                    # Method 2: Set directly on element parameter
                    try:
                        element_end_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                        if element_end_param and not element_end_param.IsReadOnly:
                            element_end_param.Set(height)
                            family_doc.Regenerate()
                            param_set_success = True
                    except Exception:
                        if not param_set_success:
                            pass
                    
                    if not param_set_success:
                        logger.warning("Failed to set extrusion end parameter for element {} - value may default to template value".format(element_number_str))
                except Exception as end_error:
                    logger.warning("Error setting extrusion end parameter for element {}: {}".format(element_number_str, end_error))
            
            t.Commit()
            
            # Save the family document
            save_options = SaveAsOptions()
            save_options.OverwriteExistingFile = True
            family_doc.SaveAs(output_path, save_options)
            logger.debug("Saved family: {}".format(op.basename(output_path)))
            
            return True, family_doc, None
            
        except Exception as recreate_error:
            t.RollBack()
            logger.debug("Element {}: Error creating extrusion: {}, skipping".format(element_number_str, recreate_error))
            return False, family_doc, "Error creating extrusion: {}".format(str(recreate_error))
            
    except Exception as family_error:
        logger.error("Error processing family document for element {}: {}".format(element_number_str, family_error))
        import traceback
        logger.error(traceback.format_exc())
        if family_doc:
            try:
                family_doc.Close(False)
            except:
                pass
            family_doc = None
            return False, None, "Error processing family document: {}".format(str(family_error))


def delete_family_instances(family, doc):
    """Delete all instances of a given family from the project.
    
    Args:
        family: Family element to delete instances for
        doc: Revit document (project document)
        
    Returns:
        tuple: (success: bool, deleted_count: int, error_reason: str or None)
    """
    # Validate family object first
    if not family:
        return False, 0, "Invalid family object (None)"
    
    # Capture family name and ID before any operations (in case family gets invalidated)
    try:
        # Check if element is valid by trying to access its properties
        family_id = family.Id
        family_name = family.Name
    except Exception as validation_error:
        # Family object is invalid or already deleted
        error_msg = "Family object is invalid or has been deleted: {}".format(str(validation_error))
        logger.error(error_msg)
        return False, 0, error_msg
    
    try:
        # Find all instances of this family
        all_instances = FilteredElementCollector(doc)\
            .OfClass(FamilyInstance)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        instances_to_delete = []
        for instance in all_instances:
            try:
                if instance.Symbol and instance.Symbol.Family and instance.Symbol.Family.Id == family_id:
                    instances_to_delete.append(instance)
            except:
                continue
        
        if not instances_to_delete:
            logger.debug("No instances found for family '{}'".format(family_name))
            return True, 0, None
        
        # Delete instances in a transaction
        t = Transaction(doc, "Delete Family Instances")
        t.Start()
        try:
            deleted_count = 0
            for instance in instances_to_delete:
                try:
                    doc.Delete(instance.Id)
                    deleted_count += 1
                except Exception as delete_error:
                    logger.warning("Failed to delete instance {}: {}".format(instance.Id, delete_error))
            
            t.Commit()
            logger.debug("Deleted {} instance(s) of family '{}'".format(deleted_count, family_name))
            return True, deleted_count, None
            
        except Exception as trans_error:
            t.RollBack()
            # Convert error to string safely
            trans_error_str = str(trans_error) if trans_error else "Unknown transaction error"
            error_msg = "Transaction failed while deleting instances of family '{}': {}".format(family_name, trans_error_str)
            logger.error(error_msg)
            return False, 0, error_msg
            
    except Exception as e:
        # Convert error to string safely (e might contain invalid object references)
        error_str = str(e) if e else "Unknown error"
        error_msg = "Error finding/deleting instances of family '{}': {}".format(family_name, error_str)
        logger.error(error_msg)
        return False, 0, error_msg


def delete_family_from_project(family, doc):
    """Delete a family from the project, including all its instances.
    
    Args:
        family: Family element to delete
        doc: Revit document (project document)
        
    Returns:
        tuple: (success: bool, error_reason: str or None)
    """
    # Validate family object first
    if not family:
        return False, "Invalid family object (None)"
    
    # Capture family name and ID before any operations (in case family gets invalidated)
    try:
        # Check if element is valid by trying to access its properties
        family_id = family.Id
        family_name = family.Name
    except Exception as validation_error:
        # Family object is invalid or already deleted
        error_msg = "Family object is invalid or has been deleted: {}".format(str(validation_error))
        logger.error(error_msg)
        return False, error_msg
    
    try:
        # First, delete all instances of this family
        instances_success, deleted_count, instances_error = delete_family_instances(family, doc)
        
        if not instances_success:
            # Convert error to string safely (it might contain invalid object references)
            try:
                error_str = str(instances_error) if instances_error else "Unknown error"
            except:
                error_str = "Error converting error message to string"
            logger.warning("Failed to delete instances before deleting family '{}': {}".format(family_name, error_str))
            # Continue anyway - try to delete family (it might fail if instances still exist)
        
        # Now delete the family itself
        t = Transaction(doc, "Delete Family")
        t.Start()
        try:
            doc.Delete(family_id)
            t.Commit()
            logger.debug("Deleted family '{}' (had {} instance(s))".format(family_name, deleted_count))
            return True, None
            
        except Exception as delete_error:
            t.RollBack()
            # Convert error to string safely
            try:
                delete_error_str = str(delete_error) if delete_error else "Unknown delete error"
            except:
                delete_error_str = "Error converting delete error to string"
            error_msg = "Failed to delete family '{}': {}".format(family_name, delete_error_str)
            logger.error(error_msg)
            return False, error_msg
            
    except Exception as e:
        # Family object might be invalidated at this point, use captured name
        # Convert error to string safely (e might contain invalid object references)
        try:
            error_str = str(e) if e else "Unknown error"
        except:
            error_str = "Error converting exception to string"
        error_msg = "Error deleting family '{}': {}".format(family_name, error_str)
        logger.error(error_msg)
        return False, error_msg


# DEPRECATED: This function is no longer used. We now delete and recreate families instead of updating them.
# Keeping it commented out for reference.
# DEPRECATED: This function is no longer used. We now delete and recreate families instead of updating them.
# Keeping it for reference but it should not be called.
def update_existing_family_geometry(spatial_element, adapter, template_info, existing_family, 
                                   doc, app, element_number_str, element_name_str):
    """Update geometry of an existing family.
    
    DEPRECATED: This function is no longer used. Use delete_family_from_project() and create_zone_family() instead.
    
    Args:
        spatial_element: Area or Room element
        adapter: SpatialElementAdapter instance
        template_info: Dict from inspect_template_family()
        existing_family: Existing Family element to update
        doc: Revit document (project document)
        app: Revit application
        element_number_str: Element number string (for logging)
        element_name_str: Element name string (for logging)
        
    Returns:
        tuple: (success: bool, error_reason: str or None)
    """
    family_doc = None
    
    try:
        # Open existing family for editing
        family_doc = doc.EditFamily(existing_family)
        if not family_doc:
            return False, "Failed to open existing family for editing"
        
        logger.debug("Opened existing family '{}' for geometry update".format(existing_family.Name))
        
        # Get the extrusion
        extrusions = FilteredElementCollector(family_doc).OfClass(Extrusion).ToElements()
        if not extrusions:
            return False, "No extrusion found in existing family"
        
        extrusion = extrusions[0]
        logger.debug("Found extrusion in existing family: ID {}".format(extrusion.Id))
        
        # Get family manager
        fm = family_doc.FamilyManager
        
        # Extract boundary loops
        loops, insertion_point, height = extract_boundary_loops(spatial_element, doc, adapter)
        
        if not loops or not insertion_point:
            return False, "Invalid boundary data (no loops or invalid insertion point)"
        
        # Translate loops to be relative to insertion point (center near origin)
        # Also translate to Z=0 to ensure sketch plane is perfectly horizontal
        translated_loops = []
        for loop_idx, loop in enumerate(loops):
            translated_loop = CurveLoop()
            loop_curves = []
            
            for curve in loop:
                try:
                    # Validate curve before translation
                    if curve is None:
                        logger.debug("Element {}: Found None curve in loop {}, skipping".format(element_number_str, loop_idx))
                        continue
                    
                    # Check if curve is valid
                    try:
                        start_pt = curve.GetEndPoint(0)
                        end_pt = curve.GetEndPoint(1)
                        
                        # Check for degenerate curves
                        if start_pt.DistanceTo(end_pt) < 0.001:
                            logger.debug("Element {}: Degenerate curve detected in loop {} (length < 1mm), skipping".format(element_number_str, loop_idx))
                            continue
                    except Exception as curve_check_error:
                        logger.debug("Element {}: Error checking curve in loop {}: {}, skipping".format(element_number_str, loop_idx, curve_check_error))
                        continue
                    
                    # Translate curve to be relative to insertion point AND set Z=0
                    first_pt = curve.GetEndPoint(0)
                    translation = Transform.CreateTranslation(XYZ(-insertion_point.X, -insertion_point.Y, -first_pt.Z))
                    translated_curve = curve.CreateTransformed(translation)
                    
                    # Validate translated curve
                    if translated_curve is None:
                        logger.debug("Element {}: Translated curve is None in loop {}, skipping".format(element_number_str, loop_idx))
                        continue
                    
                    translated_loop.Append(translated_curve)
                    loop_curves.append(translated_curve)
                except Exception as curve_error:
                    logger.debug("Element {}: Error processing curve in loop {}: {}, skipping".format(element_number_str, loop_idx, curve_error))
                    continue
            
            # Validate loop has at least 3 curves
            if len(loop_curves) < 3:
                logger.debug("Element {}: Loop {} has only {} curves (minimum 3 required), skipping".format(element_number_str, loop_idx, len(loop_curves)))
                continue
            
            # Check if loop is closed
            try:
                if translated_loop.IsOpen():
                    logger.debug("Element {}: Loop {} is open (not closed), attempting to close".format(element_number_str, loop_idx))
                    first_curve = loop_curves[0]
                    last_curve = loop_curves[-1]
                    first_pt = first_curve.GetEndPoint(0)
                    last_pt = last_curve.GetEndPoint(1)
                    
                    gap = first_pt.DistanceTo(last_pt)
                    if gap > 0.001:
                        logger.debug("Element {}: Loop {} gap is {} (too large to auto-close)".format(element_number_str, loop_idx, gap))
                        continue
            except Exception as loop_check_error:
                logger.debug("Element {}: Could not check if loop {} is closed: {}".format(element_number_str, loop_idx, loop_check_error))
            
            translated_loops.append(translated_loop)
        
        # Validate we have at least one valid loop
        if not translated_loops:
            return False, "No valid loops after translation (open loop or invalid geometry)"
        
        # Delete and recreate extrusion
        logger.debug("Updating extrusion geometry...")
        
        # Suppress "off axis" warnings
        class OffAxisFailurePreprocessor(IFailuresPreprocessor):
            def PreprocessFailures(self, failuresAccessor):
                failures = failuresAccessor.GetFailureMessages()
                for failure in failures:
                    if failure.GetSeverity().ToString() == "Warning":
                        failuresAccessor.DeleteWarning(failure)
                return FailureProcessingResult.Continue
        
        t = Transaction(family_doc, "Update Extrusion Geometry")
        
        try:
            t.Start()
        except Exception as start_error:
            if family_doc:
                try:
                    family_doc.Close(False)
                except:
                    pass
            return False, "Failed to start transaction: {}".format(str(start_error))
        
        failure_options = t.GetFailureHandlingOptions()
        failure_options.SetFailuresPreprocessor(OffAxisFailurePreprocessor())
        t.SetFailureHandlingOptions(failure_options)
        
        try:
            # Delete old extrusion
            family_doc.Delete(extrusion.Id)
            
            # Create sketch plane at Z=0
            origin = XYZ(0, 0, 0)
            normal = XYZ.BasisZ
            plane = Plane.CreateByNormalAndOrigin(normal, origin)
            sketch_plane = SketchPlane.Create(family_doc, plane)
            
            # Create CurveArrArray from loops
            curve_arr_array = CurveArrArray()
            for loop in translated_loops:
                curve_arr = CurveArray()
                for curve in loop:
                    curve_arr.Append(curve)
                curve_arr_array.Append(curve_arr)
            
            # Validate curve array
            if curve_arr_array.Size == 0:
                t.RollBack()
                return False, "CurveArrArray is empty (no valid loops)"
            
            # Validate height
            if height is not None and (height <= 0 or height > 10000):
                t.RollBack()
                return False, "Invalid height: {} (must be > 0 and < 10000)".format(height)
            
            # Create new extrusion
            family_create = family_doc.FamilyCreate
            extrusion_height = height if height is not None else 10.0  # Temporary default
            
            new_extrusion = family_create.NewExtrusion(True, curve_arr_array, sketch_plane, extrusion_height)
            
            if not new_extrusion:
                t.RollBack()
                return False, "Failed to create new extrusion (invalid geometry)"
            
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
            
            # Set family parameter values FIRST (before association)
            if template_info['extrusion_start_param']:
                try:
                    fm.Set(template_info['extrusion_start_param'], 0.0)
                except Exception as pre_set_error:
                    logger.debug("Could not pre-set start family parameter: {}".format(pre_set_error))
            
            if template_info['extrusion_end_param'] and height is not None:
                try:
                    fm.Set(template_info['extrusion_end_param'], height)
                except Exception as pre_set_error:
                    logger.debug("Could not pre-set end family parameter: {}".format(pre_set_error))
            
            # Regenerate to commit family parameter values
            family_doc.Regenerate()
            
            # Set element parameters directly (they're writable before association)
            if template_info['extrusion_start_param']:
                try:
                    start_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_START_PARAM)
                    if start_param and not start_param.IsReadOnly:
                        start_param.Set(0.0)
                        logger.debug("Set extrusion start parameter directly on element: 0.0")
                except Exception as set_error:
                    logger.debug("Could not set start parameter directly: {}".format(set_error))
            
            if template_info['extrusion_end_param'] and height is not None:
                try:
                    end_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                    if end_param and not end_param.IsReadOnly:
                        end_param.Set(height)
                        logger.debug("Set extrusion end parameter directly on element: {}".format(height))
                except Exception as set_error:
                    logger.debug("Could not set end parameter directly: {}".format(set_error))
            
            # Regenerate after setting element parameters
            family_doc.Regenerate()
            
            # Set Top Offset and Bottom Offset (user-facing parameters)
            if height is not None:
                # Set Bottom Offset first
                if template_info['bottom_offset_param']:
                    try:
                        fm.Set(template_info['bottom_offset_param'], 0.0)
                        logger.debug("Set bottom offset parameter: 0.0")
                    except Exception as bottom_error:
                        logger.debug("Could not set bottom offset parameter: {}".format(bottom_error))
                
                # Set Top Offset
                if template_info['top_offset_param']:
                    try:
                        fm.Set(template_info['top_offset_param'], height)
                        logger.debug("Set top offset parameter: {}".format(height))
                    except Exception as top_error:
                        logger.debug("Could not set top offset parameter: {}".format(top_error))
            
            # Regenerate after setting family parameters
            family_doc.Regenerate()
            
            # Associate ExtrusionEnd/ExtrusionStart to the extrusion element (if not already associated)
            if template_info['extrusion_start_param']:
                try:
                    start_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_START_PARAM)
                    if start_param:
                        # Check if already associated
                        associated_param = fm.GetAssociatedFamilyParameter(start_param)
                        if not associated_param:
                            fm.AssociateElementParameterToFamilyParameter(start_param, template_info['extrusion_start_param'])
                            logger.debug("Associated extrusion start parameter")
                except Exception as assoc_error:
                    logger.debug("Could not associate start parameter: {}".format(assoc_error))
            
            if template_info['extrusion_end_param']:
                try:
                    end_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                    if end_param:
                        # Check if already associated
                        associated_param = fm.GetAssociatedFamilyParameter(end_param)
                        if not associated_param:
                            fm.AssociateElementParameterToFamilyParameter(end_param, template_info['extrusion_end_param'])
                            logger.debug("Associated extrusion end parameter")
                except Exception as assoc_error:
                    logger.debug("Could not associate end parameter: {}".format(assoc_error))
            
            # Regenerate after association
            family_doc.Regenerate()
            
            # Associate material parameter (if not already associated)
            if template_info['material_param']:
                try:
                    mat_param = new_extrusion.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                    if mat_param:
                        associated_param = fm.GetAssociatedFamilyParameter(mat_param)
                        if not associated_param:
                            fm.AssociateElementParameterToFamilyParameter(mat_param, template_info['material_param'])
                            logger.debug("Associated material parameter")
                except Exception as assoc_error:
                    logger.debug("Could not associate material parameter: {}".format(assoc_error))
            
            # Set extrusion start/end values (verify they're still correct after association)
            if template_info['extrusion_start_param']:
                try:
                    fm.Set(template_info['extrusion_start_param'], 0.0)
                except Exception:
                    pass
            
            if template_info['extrusion_end_param'] and height is not None:
                param_set_success = False
                try:
                    # Method 1: Try setting via FamilyManager
                    try:
                        fm.Set(template_info['extrusion_end_param'], height)
                        family_doc.Regenerate()
                        param_set_success = True
                    except Exception:
                        pass
                    
                    # Method 2: Set directly on element parameter
                    try:
                        element_end_param = new_extrusion.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                        if element_end_param and not element_end_param.IsReadOnly:
                            element_end_param.Set(height)
                            family_doc.Regenerate()
                            param_set_success = True
                    except Exception:
                        if not param_set_success:
                            pass
                    
                    if not param_set_success:
                        logger.warning("Failed to set extrusion end parameter for element {} - value may default to template value".format(element_number_str))
                except Exception as end_error:
                    logger.warning("Error setting extrusion end parameter for element {}: {}".format(element_number_str, end_error))
            
            t.Commit()
            
            # When editing an existing family via EditFamily(), the family document doesn't have a file path.
            # We need to save it to a temporary file first, then load it from that file.
            # IMPORTANT: Use the EXACT same family name so LoadFamily updates the existing family.
            import tempfile
            import os
            temp_dir = tempfile.mkdtemp()
            temp_family_path = op.join(temp_dir, "{}.rfa".format(existing_family.Name))
            
            from Autodesk.Revit.DB import SaveAsOptions, IFamilyLoadOptions, FamilySource
            
            try:
                # Save family to temporary file
                save_options = SaveAsOptions()
                save_options.OverwriteExistingFile = True
                family_doc.SaveAs(temp_family_path, save_options)
                
                # Close the family document before loading
                family_doc.Close(False)
                family_doc = None
                
                # Load the family from the temporary file
                class FamilyLoadOptions(IFamilyLoadOptions):
                    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
                        overwriteParameterValues[0] = True
                        return True
                    
                    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
                        source[0] = FamilySource.Family
                        overwriteParameterValues[0] = True
                        return True
                
                load_options = FamilyLoadOptions()
                
                # Load from file path - this should update the existing family if names match
                load_result = doc.LoadFamily(temp_family_path, load_options)
                
                # Clean up temporary file and directory
                try:
                    if op.exists(temp_family_path):
                        os.remove(temp_family_path)
                    if op.exists(temp_dir):
                        os.rmdir(temp_dir)
                except:
                    pass
                
                # Check if LoadFamily succeeded
                if isinstance(load_result, tuple):
                    load_success = load_result[0] if len(load_result) > 0 else False
                elif isinstance(load_result, bool):
                    load_success = load_result
                else:
                    load_success = load_result is not None
                
                if not load_success:
                    return False, "Failed to load updated family into project (LoadFamily returned False)"
                
                logger.debug("Updated family: {}".format(existing_family.Name))
                return True, None
                
            except Exception as save_error:
                # Clean up on error
                if family_doc:
                    try:
                        family_doc.Close(False)
                    except:
                        pass
                    family_doc = None
                
                try:
                    if op.exists(temp_family_path):
                        os.remove(temp_family_path)
                    if op.exists(temp_dir):
                        os.rmdir(temp_dir)
                except:
                    pass
                
                return False, "Error saving/loading updated family: {}".format(str(save_error))
            
        except Exception as recreate_error:
                # Close document on error
                if family_doc:
                    try:
                        family_doc.Close(False)
                    except:
                        pass
                    family_doc = None
                return False, "Error saving/loading updated family: {}".format(str(save_error))
            
        except Exception as recreate_error:
            # Check transaction status before attempting rollback
            try:
                transaction_status = t.GetStatus()
                if transaction_status == TransactionStatus.Started:
                    t.RollBack()
            except Exception as rollback_error:
                pass
            
            logger.debug("Element {}: Error updating extrusion: {}, skipping".format(element_number_str, recreate_error))
            if family_doc:
                try:
                    family_doc.Close(False)
                except:
                    pass
                family_doc = None
            return False, "Error updating extrusion: {}".format(str(recreate_error))
            
    except Exception as family_error:
        logger.error("Error updating existing family for element {}: {}".format(element_number_str, family_error))
        import traceback
        logger.error(traceback.format_exc())
        if family_doc:
            try:
                family_doc.Close(False)
            except:
                pass
            family_doc = None
        return False, "Error updating existing family: {}".format(str(family_error))


def load_families(family_data_list, doc, app, progress_callback=None):
    """Load all families into the project (no transaction).
    
    Args:
        family_data_list: List of dicts with family data
        doc: Revit document
        app: Revit application
        progress_callback: Optional callback(progress_percent) for progress updates
        
    Returns:
        Updated family_data_list with 'loaded_family' keys added
    """
    from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource, Family
    
    class FamilyLoadOptions(IFamilyLoadOptions):
        def OnFamilyFound(self, familyInUse, overwriteParameterValues):
            overwriteParameterValues[0] = True
            return True
        
        def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
            source[0] = FamilySource.Family
            overwriteParameterValues[0] = True
            return True
    
    load_options = FamilyLoadOptions()
    
    total_families = len(family_data_list)
    for idx, family_data in enumerate(family_data_list):
        output_family_path = family_data.get('output_family_path')
        existing_family = family_data.get('existing_family')
        
        # Update progress if callback provided
        if progress_callback:
            progress = int(50 + (idx + 1) / float(total_families) * 25)  # 50-75%
            progress_callback(progress)
        
        if existing_family:
            family_data['loaded_family'] = existing_family
            continue
        
        if not output_family_path:
            family_data['loaded_family'] = None
            continue
        
        load_result = None
        load_family_doc = family_data.get('family_doc')
        try:
            if load_family_doc:
                # Family doc already open from phase 1 - use it directly
                load_result = load_family_doc.LoadFamily(doc, load_options)
            else:
                # Fallback: reopen from disk
                load_family_doc = app.OpenDocumentFile(output_family_path)
                load_result = load_family_doc.LoadFamily(doc, load_options)
        except Exception:
            # Safety fallback (legacy slow path)
            load_result = doc.LoadFamily(output_family_path, load_options)
        finally:
            # Close family doc after loading
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
        
        family_data['loaded_family'] = loaded_family
    
    return family_data_list


def place_instances(family_data_list, doc, adapter, template_info=None, progress_callback=None):
    """Place all instances in a single transaction.
    
    Args:
        family_data_list: List of dicts with family data (must have 'loaded_family' and 'template_info' keys)
        doc: Revit document
        adapter: SpatialElementAdapter instance
        template_info: Optional template info dict (deprecated - use template_info from family_data)
        progress_callback: Optional callback(progress_percent) for progress updates
        
    Returns:
        tuple: (success_count: int, fail_count: int, failed_elements: list)
    """
    from Autodesk.Revit.DB import Transaction
    
    success_count = 0
    fail_count = 0
    failed_elements = []
    created_instance_ids = []  # Track created instance IDs for selection
    
    t = Transaction(doc, "Place 3D Zone Instances")
    t.Start()
    try:
        total_families = len(family_data_list)
        for idx, family_data in enumerate(family_data_list):
            spatial_element = family_data['spatial_element']
            insertion_point = family_data['insertion_point']
            element_number_str = family_data['element_number_str']
            element_name_str = family_data['element_name_str']
            output_family_name = family_data['output_family_name']
            loaded_family = family_data.get('loaded_family') or family_data.get('existing_family')
            height = family_data.get('height')
            
            # Update progress if callback provided
            if progress_callback:
                progress = int(75 + (idx + 1) / float(total_families) * 25)  # 75-100%
                progress_callback(progress)
            
            try:
                # Check if instance already exists
                existing_instance = family_data.get('existing_instance')
                
                if existing_instance:
                    # Instance already exists - update its parameters instead of creating new one
                    logger.debug("Instance already exists for element {} - updating parameters".format(element_number_str))
                    
                    try:
                        # Get template_info from family_data (preferred) or use parameter
                        elem_template_info = family_data.get('template_info') or template_info
                        
                        # Update symbol parameters if adapter supports it
                        symbol = existing_instance.Symbol
                        if symbol and hasattr(adapter, 'set_symbol_parameters_before_placement') and elem_template_info:
                            adapter.set_symbol_parameters_before_placement(symbol, height, elem_template_info, doc, element_number_str)
                        
                        # Set phase if applicable
                        phase_id = adapter.get_phase_id(spatial_element)
                        adapter.set_phase_on_instance(existing_instance, phase_id)
                        
                        # Copy properties using adapter (this updates existing instance parameters)
                        adapter.copy_properties_to_instance(spatial_element, existing_instance, doc)
                        
                        logger.debug("Updated existing instance for element {}: {}".format(element_number_str, existing_instance.Id))
                        created_instance_ids.append(existing_instance.Id)
                        success_count += 1
                        
                    except Exception as update_error:
                        logger.error("Error updating existing instance for element {}: {}".format(element_number_str, update_error))
                        import traceback
                        logger.error(traceback.format_exc())
                        fail_count += 1
                        failed_elements.append({
                            'element': spatial_element,
                            'element_number_str': element_number_str,
                            'element_name_str': element_name_str,
                            'reason': 'Error updating existing instance: {}'.format(str(update_error))
                        })
                        continue
                    
                else:
                    # No existing instance - place new one
                    if not loaded_family:
                        logger.warning("Failed to load family: {}".format(output_family_name))
                        fail_count += 1
                        failed_elements.append({
                            'element': spatial_element,
                            'element_number_str': element_number_str,
                            'element_name_str': element_name_str,
                            'reason': 'Failed to load family'
                        })
                        continue
                    
                    logger.debug("Loaded family: {}".format(loaded_family.Name))
                    
                    symbol_ids_set = loaded_family.GetFamilySymbolIds()
                    symbol_ids = list(symbol_ids_set) if symbol_ids_set else []
                    if not symbol_ids or len(symbol_ids) == 0:
                        logger.warning("No symbols found in loaded family")
                        fail_count += 1
                        failed_elements.append({
                            'element': spatial_element,
                            'element_number_str': element_number_str,
                            'element_name_str': element_name_str,
                            'reason': 'No symbols found in loaded family'
                        })
                        continue
                    
                    symbol = doc.GetElement(symbol_ids[0])
                    
                    if symbol and not symbol.IsActive:
                        symbol.Activate()
                        doc.Regenerate()
                    
                    if not symbol:
                        logger.warning("No active symbol found for family {}".format(loaded_family.Name))
                        fail_count += 1
                        failed_elements.append({
                            'element': spatial_element,
                            'element_number_str': element_number_str,
                            'element_name_str': element_name_str,
                            'reason': 'No active symbol found'
                        })
                        continue
                    
                    # Rooms-specific: Set Top Offset on symbol before instance creation
                    # This is handled by checking if adapter has a method for this
                    # Get template_info from family_data (preferred) or use parameter
                    elem_template_info = family_data.get('template_info') or template_info
                    if hasattr(adapter, 'set_symbol_parameters_before_placement') and elem_template_info:
                        adapter.set_symbol_parameters_before_placement(symbol, height, elem_template_info, doc, element_number_str)
                    
                    # Get level
                    level_id = adapter.get_level_id(spatial_element)
                    level = doc.GetElement(level_id) if level_id else None
                    
                    # Create placement point with Z=0 relative to level
                    placement_point = XYZ(
                        insertion_point.X,
                        insertion_point.Y,
                        0.0
                    )
                    
                    # Place instance
                    instance = doc.Create.NewFamilyInstance(
                        placement_point,
                        symbol,
                        level,
                        StructuralType.NonStructural
                    )
                    
                    if not instance:
                        logger.warning("Failed to place instance for element {}".format(element_number_str))
                        fail_count += 1
                        failed_elements.append({
                            'element': spatial_element,
                            'element_number_str': element_number_str,
                            'element_name_str': element_name_str,
                            'reason': 'Failed to place instance'
                        })
                        continue
                    
                    logger.debug("Placed instance for element {}: {}".format(element_number_str, instance.Id))
                    
                    # Set phase if applicable
                    phase_id = adapter.get_phase_id(spatial_element)
                    adapter.set_phase_on_instance(instance, phase_id)
                    
                    # Copy properties using adapter
                    adapter.copy_properties_to_instance(spatial_element, instance, doc)
                    
                    # Track created instance ID
                    created_instance_ids.append(instance.Id)
                    success_count += 1
                
            except Exception as place_error:
                logger.error("Error placing family for element {}: {}".format(element_number_str, place_error))
                import traceback
                logger.error(traceback.format_exc())
                fail_count += 1
                failed_elements.append({
                    'element': spatial_element,
                    'element_number_str': element_number_str,
                    'element_name_str': element_name_str,
                    'reason': 'Error loading/placing family: {}'.format(str(place_error))
                })
        
        t.Commit()
    except Exception as tx_error:
        t.RollBack()
        logger.error("Transaction failed, rolled back: {}".format(tx_error))
        import traceback
        logger.error(traceback.format_exc())
        fail_count += len(family_data_list) - success_count
    
    return success_count, fail_count, failed_elements, created_instance_ids

