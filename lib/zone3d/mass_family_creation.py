# -*- coding: utf-8 -*-
"""Mass family creation functions using NewExtrusionForm API.

This module handles Mass family creation which requires different API than standard families:
- Uses NewExtrusionForm() instead of NewExtrusion()
- Requires ReferenceArray of ModelCurve references instead of CurveArrArray
- Works with Conceptual Mass family templates
"""

import os.path as op
import shutil
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInParameter, Category,
    BuiltInCategory, ElementId, Transaction, TransactionStatus, FailureProcessingResult,
    IFailuresPreprocessor, CurveLoop, XYZ, Transform, Plane, SketchPlane, 
    SaveAsOptions, FamilyInstance, Family, Line, ModelCurve, ReferenceArray,
    Form, GenericForm
)
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import script

logger = script.get_logger()


def get_mass_template_path(extension_dir, template_family_name="MassZone.rfa"):
    """Get the path to the Mass family template.
    
    Args:
        extension_dir: Extension root directory
        template_family_name: Name of template family file
        
    Returns:
        str: Template file path or None if not found
    """
    template_family_path = op.join(extension_dir, template_family_name)
    
    logger.debug("Looking for Mass template at: {}".format(template_family_path))
    
    if not op.exists(template_family_path):
        logger.error("Mass template family not found at: {}".format(template_family_path))
        return None
    return template_family_path


def inspect_mass_template(family_doc):
    """Inspect the Mass family template to find Form element and parameters.
    
    Args:
        family_doc: Family document (Mass family)
        
    Returns:
        dict: {
            'form': Form element (if found),
            'height_param': FamilyParameter for height,
            'material_param': FamilyParameter for material
        }
    """
    result = {
        'form': None,
        'height_param': None,
        'material_param': None
    }
    
    try:
        # Find Form elements (Mass families use Form, not Extrusion)
        forms = FilteredElementCollector(family_doc).OfClass(Form).ToElements()
        if forms:
            result['form'] = forms[0]
            logger.debug("Found Form in Mass template: ID {}".format(result['form'].Id))
        else:
            # Try GenericForm as fallback
            generic_forms = FilteredElementCollector(family_doc).OfClass(GenericForm).ToElements()
            if generic_forms:
                result['form'] = generic_forms[0]
                logger.debug("Found GenericForm in Mass template: ID {}".format(result['form'].Id))
            else:
                logger.warning("No Form or GenericForm found in Mass template")
                return result
        
        # Get family manager for parameters
        fm = family_doc.FamilyManager
        
        # Find height parameter
        height_param_names = ["Height", "Top Offset", "Extrusion Height", "ExtrusionHeight"]
        for param_name in height_param_names:
            try:
                param = fm.get_Parameter(param_name)
                if param:
                    result['height_param'] = param
                    logger.debug("Found height parameter: {}".format(param_name))
                    break
            except Exception:
                pass
        
        # Find material parameter
        material_param_names = ["Material", "Mass Material", "Form Material"]
        for param_name in material_param_names:
            try:
                param = fm.get_Parameter(param_name)
                if param:
                    result['material_param'] = param
                    logger.debug("Found material parameter: {}".format(param_name))
                    break
            except Exception:
                pass
        
    except Exception as e:
        logger.error("Error inspecting Mass template: {}".format(e))
        import traceback
        logger.error(traceback.format_exc())
    
    return result


def extract_boundary_loops_for_mass(spatial_element, doc, adapter, levels_cache=None):
    """Extract boundary loops from a spatial element for Mass creation.
    
    Similar to family_creation.extract_boundary_loops but returns data
    suitable for Mass family creation.
    
    Args:
        spatial_element: FilledRegion or other spatial element
        doc: Revit document
        adapter: SpatialElementAdapter instance
        levels_cache: Optional pre-collected and sorted list of Level elements
        
    Returns:
        tuple: (curves: list of Curve objects, insertion_point: XYZ, height: float or None)
    """
    curves = []
    insertion_point = None
    height = None
    
    try:
        # Get boundary segments using adapter
        boundary_segments = adapter.get_boundary_segments(spatial_element)
        
        if not boundary_segments or len(boundary_segments) == 0:
            logger.warning("Element {} has no boundary segments".format(spatial_element.Id))
            return curves, insertion_point, height
        
        # Process boundary loops - collect all curves
        all_points = []
        for segment_group in boundary_segments:
            for segment in segment_group:
                curve = segment.GetCurve()
                if curve:
                    try:
                        start_pt = curve.GetEndPoint(0)
                        end_pt = curve.GetEndPoint(1)
                        all_points.append(start_pt)
                        all_points.append(end_pt)
                        curves.append(curve)
                    except Exception as e:
                        logger.debug("Error processing curve segment: {}".format(e))
                        continue
        
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
        
        # Calculate height using adapter
        height = adapter.calculate_height(spatial_element, doc, levels_cache)
        
    except Exception as e:
        logger.error("Error extracting boundary loops for Mass: {}".format(e))
        import traceback
        logger.error(traceback.format_exc())
    
    return curves, insertion_point, height


def translate_curves_to_origin(curves, insertion_point):
    """Translate curves to be relative to origin (for family creation).
    
    Args:
        curves: List of Curve objects
        insertion_point: XYZ point to use as new origin
        
    Returns:
        List of translated Curve objects
    """
    translated_curves = []
    
    for curve in curves:
        try:
            if curve is None:
                continue
            
            # Check for degenerate curves
            start_pt = curve.GetEndPoint(0)
            end_pt = curve.GetEndPoint(1)
            if start_pt.DistanceTo(end_pt) < 0.001:
                logger.debug("Degenerate curve detected, skipping")
                continue
            
            # Translate curve to be relative to insertion point AND set Z=0
            translation = Transform.CreateTranslation(
                XYZ(-insertion_point.X, -insertion_point.Y, -start_pt.Z)
            )
            translated_curve = curve.CreateTransformed(translation)
            
            if translated_curve:
                translated_curves.append(translated_curve)
                
        except Exception as e:
            logger.debug("Error translating curve: {}".format(e))
            continue
    
    return translated_curves


def create_model_curves_and_references(family_doc, curves):
    """Create ModelCurves in family document and collect their references.
    
    This is the key difference from standard family creation - Mass families
    require ModelCurve references in a ReferenceArray, not CurveArrArray.
    
    Args:
        family_doc: Mass family document
        curves: List of translated Curve objects
        
    Returns:
        tuple: (ReferenceArray, list of ModelCurve objects) or (None, None) on failure
    """
    ref_array = ReferenceArray()
    model_curves = []
    
    try:
        # Create a single sketch plane at Z=0 for all curves
        origin = XYZ(0, 0, 0)
        normal = XYZ.BasisZ
        plane = Plane.CreateByNormalAndOrigin(normal, origin)
        sketch_plane = SketchPlane.Create(family_doc, plane)
        
        for curve in curves:
            try:
                # Create ModelCurve
                model_curve = family_doc.FamilyCreate.NewModelCurve(curve, sketch_plane)
                
                if model_curve:
                    # Get the reference from the curve
                    geom_curve = model_curve.GeometryCurve
                    if geom_curve and hasattr(geom_curve, 'Reference'):
                        ref = geom_curve.Reference
                        if ref:
                            ref_array.Append(ref)
                            model_curves.append(model_curve)
                            logger.debug("Created ModelCurve and added reference")
                        else:
                            logger.warning("ModelCurve has no valid reference")
                    else:
                        logger.warning("Could not get GeometryCurve reference")
                else:
                    logger.warning("Failed to create ModelCurve")
                    
            except Exception as curve_error:
                logger.debug("Error creating ModelCurve: {}".format(curve_error))
                continue
        
        if ref_array.Size == 0:
            logger.error("No valid references created from curves")
            return None, None
        
        logger.debug("Created {} ModelCurves with references".format(ref_array.Size))
        return ref_array, model_curves
        
    except Exception as e:
        logger.error("Error creating ModelCurves: {}".format(e))
        import traceback
        logger.error(traceback.format_exc())
        return None, None


def create_mass_family(spatial_element, adapter, template_info, template_path, output_path, 
                       temp_dir, doc, app, element_number_str, element_name_str):
    """Create a Mass family from a spatial element.
    
    This uses NewExtrusionForm() API which is required for Mass families.
    
    Args:
        spatial_element: FilledRegion or other spatial element
        adapter: SpatialElementAdapter instance
        template_info: Dict from inspect_mass_template()
        template_path: Path to Mass template family file
        output_path: Path to save output family file
        temp_dir: Temporary directory path
        doc: Revit document
        app: Revit application
        element_number_str: Element number string (for logging)
        element_name_str: Element name string (for logging)
        
    Returns:
        tuple: (success: bool, family_doc: Document or None, error_reason: str or None)
    """
    family_doc = None
    
    try:
        # Copy template to output location
        shutil.copy2(template_path, output_path)
        logger.debug("Copied Mass template to: {}".format(output_path))
        
        # Open the copied family document
        family_doc = app.OpenDocumentFile(output_path)
        if not family_doc:
            return False, None, "Failed to open Mass family document"
        
        # Verify this is a Mass family by checking for Form elements
        forms = FilteredElementCollector(family_doc).OfClass(Form).ToElements()
        generic_forms = FilteredElementCollector(family_doc).OfClass(GenericForm).ToElements()
        
        existing_form = None
        if forms:
            existing_form = forms[0]
        elif generic_forms:
            existing_form = generic_forms[0]
        
        if not existing_form:
            logger.warning("No existing Form found in template - will create new one")
        
        # Get family manager
        fm = family_doc.FamilyManager
        
        # Extract boundary curves
        curves, insertion_point, height = extract_boundary_loops_for_mass(
            spatial_element, doc, adapter
        )
        
        if not curves or not insertion_point:
            return False, family_doc, "Invalid boundary data (no curves or invalid insertion point)"
        
        # Translate curves to origin
        translated_curves = translate_curves_to_origin(curves, insertion_point)
        
        if not translated_curves or len(translated_curves) < 3:
            return False, family_doc, "Not enough valid curves after translation (need at least 3)"
        
        # Suppress warnings
        class MassFailurePreprocessor(IFailuresPreprocessor):
            def PreprocessFailures(self, failuresAccessor):
                failures = failuresAccessor.GetFailureMessages()
                for failure in failures:
                    if failure.GetSeverity().ToString() == "Warning":
                        failuresAccessor.DeleteWarning(failure)
                return FailureProcessingResult.Continue
        
        t = Transaction(family_doc, "Create Mass Form")
        t.Start()
        failure_options = t.GetFailureHandlingOptions()
        failure_options.SetFailuresPreprocessor(MassFailurePreprocessor())
        t.SetFailureHandlingOptions(failure_options)
        
        try:
            # Delete existing form if present
            if existing_form:
                try:
                    family_doc.Delete(existing_form.Id)
                    logger.debug("Deleted existing Form from template")
                except Exception as del_error:
                    logger.warning("Could not delete existing Form: {}".format(del_error))
            
            # Create ModelCurves and get references
            ref_array, model_curves = create_model_curves_and_references(
                family_doc, translated_curves
            )
            
            if not ref_array or ref_array.Size == 0:
                t.RollBack()
                return False, family_doc, "Failed to create ModelCurves with valid references"
            
            # Calculate extrusion direction (Z-up with height as length)
            extrusion_height = height if height and height > 0 else 10.0
            direction = XYZ(0, 0, extrusion_height)
            
            # Create the extrusion form using Mass API
            logger.debug("Creating NewExtrusionForm with {} references, height={}".format(
                ref_array.Size, extrusion_height))
            
            try:
                new_form = family_doc.FamilyCreate.NewExtrusionForm(True, ref_array, direction)
                
                if not new_form:
                    t.RollBack()
                    return False, family_doc, "NewExtrusionForm returned None"
                
                logger.debug("Created new Mass Form: ID {}".format(new_form.Id))
                
            except Exception as form_error:
                t.RollBack()
                error_msg = "NewExtrusionForm failed: {}".format(str(form_error))
                logger.error(error_msg)
                return False, family_doc, error_msg
            
            # Set height parameter if available
            if template_info.get('height_param') and height:
                try:
                    fm.Set(template_info['height_param'], height)
                    logger.debug("Set height parameter to: {}".format(height))
                except Exception as param_error:
                    logger.debug("Could not set height parameter: {}".format(param_error))
            
            # Regenerate
            family_doc.Regenerate()
            
            t.Commit()
            
            # Save the family document
            save_options = SaveAsOptions()
            save_options.OverwriteExistingFile = True
            family_doc.SaveAs(output_path, save_options)
            logger.debug("Saved Mass family: {}".format(op.basename(output_path)))
            
            return True, family_doc, None
            
        except Exception as recreate_error:
            if t.GetStatus() == TransactionStatus.Started:
                t.RollBack()
            logger.error("Error creating Mass Form: {}".format(recreate_error))
            import traceback
            logger.error(traceback.format_exc())
            return False, family_doc, "Error creating Mass Form: {}".format(str(recreate_error))
            
    except Exception as family_error:
        logger.error("Error processing Mass family for element {}: {}".format(
            element_number_str, family_error))
        import traceback
        logger.error(traceback.format_exc())
        if family_doc:
            try:
                family_doc.Close(False)
            except:
                pass
            family_doc = None
        return False, None, "Error processing Mass family: {}".format(str(family_error))


def delete_mass_family_instances(family, doc):
    """Delete all instances of a Mass family from the project.
    
    Args:
        family: Family element to delete instances for
        doc: Revit document (project document)
        
    Returns:
        tuple: (success: bool, deleted_count: int, error_reason: str or None)
    """
    if not family:
        return False, 0, "Invalid family object (None)"
    
    try:
        family_id = family.Id
        family_name = family.Name
    except Exception as validation_error:
        return False, 0, "Family object is invalid: {}".format(str(validation_error))
    
    try:
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
            return True, 0, None
        
        t = Transaction(doc, "Delete Mass Family Instances")
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
            return True, deleted_count, None
            
        except Exception as trans_error:
            t.RollBack()
            return False, 0, "Transaction failed: {}".format(str(trans_error))
            
    except Exception as e:
        return False, 0, "Error deleting instances: {}".format(str(e))


def delete_mass_family_from_project(family, doc):
    """Delete a Mass family from the project, including all its instances.
    
    Args:
        family: Family element to delete
        doc: Revit document (project document)
        
    Returns:
        tuple: (success: bool, error_reason: str or None)
    """
    if not family:
        return False, "Invalid family object (None)"
    
    try:
        family_id = family.Id
        family_name = family.Name
    except Exception as validation_error:
        return False, "Family object is invalid: {}".format(str(validation_error))
    
    try:
        # First delete instances
        instances_success, deleted_count, instances_error = delete_mass_family_instances(family, doc)
        
        if not instances_success:
            logger.warning("Failed to delete instances: {}".format(instances_error))
        
        # Delete the family
        t = Transaction(doc, "Delete Mass Family")
        t.Start()
        try:
            doc.Delete(family_id)
            t.Commit()
            logger.debug("Deleted Mass family '{}' (had {} instances)".format(family_name, deleted_count))
            return True, None
            
        except Exception as delete_error:
            t.RollBack()
            return False, "Failed to delete family: {}".format(str(delete_error))
            
    except Exception as e:
        return False, "Error deleting family: {}".format(str(e))


def load_mass_families(family_data_list, doc, app, progress_callback=None):
    """Load all Mass families into the project.
    
    Args:
        family_data_list: List of dicts with family data
        doc: Revit document
        app: Revit application
        progress_callback: Optional callback(progress_percent) for progress updates
        
    Returns:
        Updated family_data_list with 'loaded_family' keys added
    """
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
    
    total_families = len(family_data_list)
    for idx, family_data in enumerate(family_data_list):
        output_family_path = family_data.get('output_family_path')
        existing_family = family_data.get('existing_family')
        
        if progress_callback:
            progress = int(50 + (idx + 1) / float(total_families) * 25)
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
                load_result = load_family_doc.LoadFamily(doc, load_options)
            else:
                load_family_doc = app.OpenDocumentFile(output_family_path)
                load_result = load_family_doc.LoadFamily(doc, load_options)
        except Exception:
            load_result = doc.LoadFamily(output_family_path, load_options)
        finally:
            if load_family_doc:
                try:
                    load_family_doc.Close(False)
                except:
                    pass
        
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
                loaded_family = load_result
        except:
            loaded_family = None
        
        family_data['loaded_family'] = loaded_family
    
    return family_data_list


def place_mass_instances(family_data_list, doc, adapter, progress_callback=None):
    """Place all Mass instances in a single transaction.
    
    Args:
        family_data_list: List of dicts with family data
        doc: Revit document
        adapter: SpatialElementAdapter instance
        progress_callback: Optional callback(progress_percent) for progress updates
        
    Returns:
        tuple: (success_count: int, fail_count: int, failed_elements: list, created_instance_ids: list)
    """
    success_count = 0
    fail_count = 0
    failed_elements = []
    created_instance_ids = []
    
    t = Transaction(doc, "Place Mass Instances")
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
            
            if progress_callback:
                progress = int(75 + (idx + 1) / float(total_families) * 25)
                progress_callback(progress)
            
            try:
                if not loaded_family:
                    logger.warning("Failed to load Mass family: {}".format(output_family_name))
                    fail_count += 1
                    failed_elements.append({
                        'element': spatial_element,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'reason': 'Failed to load Mass family'
                    })
                    continue
                
                symbol_ids_set = loaded_family.GetFamilySymbolIds()
                symbol_ids = list(symbol_ids_set) if symbol_ids_set else []
                if not symbol_ids:
                    fail_count += 1
                    failed_elements.append({
                        'element': spatial_element,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'reason': 'No symbols found in loaded Mass family'
                    })
                    continue
                
                symbol = doc.GetElement(symbol_ids[0])
                
                if symbol and not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()
                
                if not symbol:
                    fail_count += 1
                    failed_elements.append({
                        'element': spatial_element,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'reason': 'No active symbol found'
                    })
                    continue
                
                # Get level
                level_id = adapter.get_level_id(spatial_element)
                level = doc.GetElement(level_id) if level_id else None
                
                # Create placement point
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
                    fail_count += 1
                    failed_elements.append({
                        'element': spatial_element,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'reason': 'Failed to place Mass instance'
                    })
                    continue
                
                logger.debug("Placed Mass instance for element {}: {}".format(
                    element_number_str, instance.Id))
                
                # Set phase if applicable
                phase_id = adapter.get_phase_id(spatial_element)
                adapter.set_phase_on_instance(instance, phase_id)
                
                # Copy properties
                adapter.copy_properties_to_instance(spatial_element, instance, doc)
                
                created_instance_ids.append(instance.Id)
                success_count += 1
                
            except Exception as place_error:
                logger.error("Error placing Mass for element {}: {}".format(
                    element_number_str, place_error))
                fail_count += 1
                failed_elements.append({
                    'element': spatial_element,
                    'element_number_str': element_number_str,
                    'element_name_str': element_name_str,
                    'reason': 'Error placing Mass: {}'.format(str(place_error))
                })
        
        t.Commit()
    except Exception as tx_error:
        t.RollBack()
        logger.error("Transaction failed: {}".format(tx_error))
        fail_count += len(family_data_list) - success_count
    
    return success_count, fail_count, failed_elements, created_instance_ids
