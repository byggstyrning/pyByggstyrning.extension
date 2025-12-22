# -*- coding: utf-8 -*-
"""Main orchestration function for creating 3D Zones from spatial elements."""

import os.path as op
import tempfile
import shutil
from Autodesk.Revit.DB import (
    FilteredElementCollector, Family, SpatialElement, SaveAsOptions, Level
)
from pyrevit import script, forms, revit
from zone3d import family_creation as fc_module

logger = script.get_logger()


def report_failed_elements(failed_elements, element_type_name):
    """Report failed elements with linkify.
    
    Args:
        failed_elements: List of dicts with failed element info
        element_type_name: String like "Areas" or "Rooms"
    """
    if not failed_elements:
        return
    
    output = script.get_output()
    output.print_md("## ⚠️ Failed {} - Invalid Geometry".format(element_type_name))
    output.print_md("**{} {}(s) could not be processed due to invalid geometry.**".format(
        len(failed_elements), element_type_name.rstrip('s')))
    output.print_md("Click {} ID to zoom to {} in Revit".format(
        element_type_name.rstrip('s'), element_type_name.rstrip('s').lower()))
    output.print_md("---")
    
    # Group by reason for better overview
    by_reason = {}
    for item in failed_elements:
        reason = item['reason']
        if reason not in by_reason:
            by_reason[reason] = []
        by_reason[reason].append(item)
    
    # Print by reason
    for reason, items in sorted(by_reason.items()):
        output.print_md("### {} ({} {})".format(reason, len(items), element_type_name.rstrip('s').lower()))
        output.print_md("")
        
        # Print each element with clickable link
        for item in items:
            element = item['element']
            element_number = item['element_number_str']
            element_name = item['element_name_str']
            element_link = output.linkify(element.Id)
            output.print_md("- **{} {} - {}**: {}".format(
                element_type_name.rstrip('s'), element_number, element_name, element_link))
        
        output.print_md("")  # Empty line between reasons
    
    output.print_md("---")
    output.print_md("**Tip:** Fix the {} boundaries in Revit and run the script again.".format(
        element_type_name.rstrip('s').lower()))


def create_zones_from_spatial_elements(
    spatial_elements,
    doc,
    adapter,
    extension_dir,
    pushbutton_dir,
    show_filter_dialog_func,
    element_type_name,
    template_family_name="3DZone.rfa"
):
    """Main orchestration function for creating 3D Zones.
    
    Args:
        spatial_elements: List of Area or Room elements
        doc: Revit document
        adapter: SpatialElementAdapter instance
        extension_dir: Extension root directory
        pushbutton_dir: Pushbutton directory (for XAML files)
        show_filter_dialog_func: Function to show filter dialog
        element_type_name: String for progress/error messages (e.g., "Areas" or "Rooms")
        template_family_name: Name of template family file
        
    Returns:
        tuple: (success_count, fail_count, failed_elements)
    """
    app = doc.Application
    
    # Check if document supports families (must be a project document)
    if doc.IsFamilyDocument:
        forms.alert("This tool can only be used in project documents, not family documents.",
                   title="Invalid Document Type", exitscript=True)
    
    # Verify template family exists
    template_path = fc_module.get_template_family_path(extension_dir, template_family_name)
    if not template_path:
        template_family_path = op.join(extension_dir, template_family_name)
        logger.error("Extension dir: {}".format(extension_dir))
        logger.error("Template path: {}".format(template_family_path))
        logger.error("Template exists: {}".format(op.exists(template_family_path)))
        forms.alert("Template family '{}' not found at:\n{}\n\nExtension dir: {}".format(
            template_family_name, template_family_path, extension_dir),
            title="Template Not Found", exitscript=True)
    
    # Create temporary directory for RFA files
    temp_dir = tempfile.mkdtemp(prefix="pyBS_3DZone_")
    logger.debug("Created temporary directory: {}".format(temp_dir))
    
    # Show filter dialog
    selected_elements = show_filter_dialog_func(spatial_elements, doc)
    
    # Check if user cancelled or selected no elements
    if not selected_elements:
        logger.debug("No {} selected, exiting.".format(element_type_name.lower()))
        script.exit()
    
    logger.debug("User selected {} {} for 3D Zone creation".format(len(selected_elements), element_type_name.lower()))
    
    # Check if template family exists in project first
    template_family_name_in_project = "3DZone"  # Name without .rfa extension
    project_template_family = None
    families_cache = FilteredElementCollector(doc).OfClass(Family).ToElements()
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
                
                # Handle LoadFamily return value
                if isinstance(load_result, tuple):
                    project_template_family = load_result[1] if len(load_result) > 1 else None
                elif isinstance(load_result, bool):
                    if load_result:
                        families_cache = FilteredElementCollector(doc).OfClass(Family).ToElements()
                        project_template_family = next((f for f in families_cache if f.Name == template_family_name_in_project), None)
                    else:
                        project_template_family = None
                else:
                    project_template_family = load_result
                
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
        try:
            logger.debug("Exporting template family from project...")
            temp_template_path = op.join(temp_dir, "temp_3DZone_template.rfa")
            
            template_family_doc = doc.EditFamily(project_template_family)
            if template_family_doc:
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
    
    # Open template family document and inspect
    logger.debug("Opening template family for inspection...")
    try:
        template_family_doc = app.OpenDocumentFile(template_path_to_use)
        if not template_family_doc:
            raise Exception("Failed to open template family document")
        
        template_info = fc_module.inspect_template_family(template_family_doc)
        
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
    
    # Process elements
    success_count = 0
    fail_count = 0
    failed_elements = []
    family_data_list = []
    created_instance_ids = []  # Initialize instance IDs list
    
    logger.debug("Starting to process {} {}...".format(len(selected_elements), element_type_name.lower()))
    
    # Cache levels collection (sorted by elevation for height calculations)
    levels_cache = FilteredElementCollector(doc).OfClass(Level).ToElements()
    levels_cache_sorted = sorted(levels_cache, key=lambda l: l.Elevation)
    
    # Cache families collection
    families_cache = FilteredElementCollector(doc).OfClass(Family).ToElements()
    
    # Cache zone instances for checking existing instances
    from Autodesk.Revit.DB import FamilyInstance, BuiltInCategory, Category
    all_generic_instances = FilteredElementCollector(doc)\
        .OfClass(FamilyInstance)\
        .OfCategory(BuiltInCategory.OST_GenericModel)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Filter to only 3DZone families
    zone_instances_cache = []
    family_name_prefix = adapter.get_family_name_prefix()
    for instance in all_generic_instances:
        try:
            instance_family_name = instance.Symbol.Family.Name
            if instance_family_name.startswith(family_name_prefix):
                zone_instances_cache.append(instance)
        except:
            pass
    
    # Phase 1: Process elements and create family documents (0-50% progress)
    total_elements = len(selected_elements)
    with forms.ProgressBar(title="Creating 3D Zone Families and Placing Instances ({} {})".format(
            total_elements, element_type_name.lower())) as pb:
        
        for elem_idx, spatial_element in enumerate(selected_elements):
            try:
                element_id = spatial_element.Id.IntegerValue
                
                # Get element info using adapter
                element_number_str = adapter.get_number(spatial_element)
                if not element_number_str or element_number_str == "?":
                    element_number_str = str(element_id)
                
                element_name_str = adapter.get_name(spatial_element)
                
                logger.debug("Processing {} {} ({}/{}): {} - {}".format(
                    element_type_name.rstrip('s').lower(), element_number_str, elem_idx + 1,
                    len(selected_elements), element_name_str, element_id))
                
                # Update progress bar: Phase 1 is 0-50%
                phase1_progress = int((elem_idx + 1) / float(total_elements) * 50)
                pb.update_progress(phase1_progress, 100)
                
                # Extract boundary loops using shared function
                loops, insertion_point, height = fc_module.extract_boundary_loops(
                    spatial_element, doc, adapter, levels_cache_sorted)
                
                # Check for valid boundary data
                if not loops or not insertion_point:
                    logger.warning("{} {} (ID: {}) - invalid boundary data, skipping".format(
                        element_type_name.rstrip('s'), element_number_str, element_id))
                    fail_count += 1
                    failed_elements.append({
                        'element': spatial_element,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'reason': 'Invalid boundary data (no loops or no insertion point)'
                    })
                    continue
                
                # If height is explicitly set but <= 0, that's invalid
                if height is not None and height <= 0:
                    logger.warning("{} {} (ID: {}) - invalid height ({}), skipping".format(
                        element_type_name.rstrip('s'), element_number_str, element_id, height))
                    fail_count += 1
                    failed_elements.append({
                        'element': spatial_element,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'reason': 'Invalid height (height <= 0)'
                    })
                    continue
                
                # Create output family name and path
                sanitized_number = adapter.sanitize_number(element_number_str)
                family_name_prefix = adapter.get_family_name_prefix()
                output_family_name = "{}{}_{}.rfa".format(family_name_prefix, sanitized_number, element_id)
                output_family_path = op.join(temp_dir, output_family_name)
                family_name_without_ext = op.splitext(output_family_name)[0]
                
                # Check if family already exists in project
                existing_family = None
                for fam in families_cache:
                    try:
                        # Check if family is valid before accessing Name property
                        # Deleted families will throw exceptions when accessing properties
                        if fam and fam.Name == family_name_without_ext:
                            existing_family = fam
                            logger.debug("Family '{}' already exists in project".format(family_name_without_ext))
                            break
                    except Exception as fam_access_error:
                        # Skip invalid/deleted families - they may have been deleted earlier in the loop
                        logger.debug("Skipping invalid family reference in cache (may have been deleted): {}".format(fam_access_error))
                        continue
                
                # If family exists, delete it and all its instances, then create a new one
                if existing_family:
                    # Store family name before deletion (in case it becomes invalid)
                    existing_family_name = None
                    try:
                        existing_family_name = existing_family.Name
                    except:
                        existing_family_name = family_name_without_ext  # Fallback to expected name
                    
                    logger.debug("Family '{}' already exists for {} {} - deleting and recreating".format(
                        existing_family_name, element_type_name.rstrip('s').lower(), element_number_str))
                    
                    # Delete all instances and the family itself
                    success, error_reason = fc_module.delete_family_from_project(existing_family, doc)
                    
                    if not success:
                        logger.warning("Failed to delete existing family '{}' for {} {}: {}. Will attempt to create new family anyway.".format(
                            existing_family_name, element_type_name.rstrip('s').lower(), element_number_str, error_reason))
                        # Continue to create new family - Revit may handle duplicate names or the deletion might have partially succeeded
                    
                    # Clear the existing_family reference so we create a new one
                    existing_family = None
                
                # Create family if it doesn't exist (or was just deleted)
                if not existing_family:
                    # Create family using shared function
                    success, family_doc, error_reason = fc_module.create_zone_family(
                        spatial_element,
                        adapter,
                        template_info,
                        template_path_to_use,
                        output_family_path,
                        temp_dir,
                        doc,
                        app,
                        element_number_str,
                        element_name_str
                    )
                    
                    if not success:
                        logger.warning("Failed to create family for {} {}: {}".format(
                            element_type_name.rstrip('s').lower(), element_number_str, error_reason))
                        fail_count += 1
                        failed_elements.append({
                            'element': spatial_element,
                            'element_number_str': element_number_str,
                            'element_name_str': element_name_str,
                            'reason': error_reason or 'Failed to create family'
                        })
                        if family_doc:
                            try:
                                family_doc.Close(False)
                            except:
                                pass
                        continue
                    
                    # Store family data for second pass
                    family_data_list.append({
                        'spatial_element': spatial_element,
                        'output_family_path': output_family_path,
                        'insertion_point': insertion_point,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'output_family_name': output_family_name,
                        'family_name_without_ext': family_name_without_ext,
                        'existing_family': None,
                        'family_doc': family_doc,  # Keep reference to open family doc
                        'height': height,
                        'template_info': template_info
                    })
                    
            except Exception as e:
                logger.error("Error processing {} {} (ID: {}): {}".format(
                    element_type_name.rstrip('s').lower(),
                    element_number_str if 'element_number_str' in locals() else "Unknown",
                    element_id if 'element_id' in locals() else "Unknown",
                    str(e)))
                import traceback
                logger.error(traceback.format_exc())
                fail_count += 1
                try:
                    if 'spatial_element' in locals():
                        failed_elements.append({
                            'element': spatial_element,
                            'element_number_str': element_number_str if 'element_number_str' in locals() else "Unknown",
                            'element_name_str': element_name_str if 'element_name_str' in locals() else "Unknown",
                            'reason': 'Error processing {}: {}'.format(element_type_name.rstrip('s').lower(), str(e))
                        })
                except:
                    pass
        
        # Phase 2: Load families and place instances (50-100% progress)
        if family_data_list:
            # Load families (no transaction)
            def progress_callback(progress):
                pb.update_progress(progress, 100)
            
            fc_module.load_families(family_data_list, doc, app, progress_callback)
            
            # Place instances (single transaction)
            # template_info is stored in each family_data dict
            place_success, place_fail, place_failed, created_instance_ids = fc_module.place_instances(
                family_data_list, doc, adapter, None, progress_callback)
            
            success_count += place_success
            fail_count += place_fail
            failed_elements.extend(place_failed)
    
    # Close template family doc
    if template_family_doc:
        template_family_doc.Close(False)
    
    # Clean up temporary files
    try:
        if op.exists(temp_dir):
            shutil.rmtree(temp_dir)
            logger.debug("Cleaned up temporary directory: {}".format(temp_dir))
    except Exception as cleanup_error:
        logger.warning("Error cleaning up temporary directory: {}".format(cleanup_error))
    
    # Log results
    logger.debug("Created {} 3D Zone families from {} {} ({} failed)".format(
        success_count, len(selected_elements), element_type_name, fail_count))
    
    # Report failed elements
    report_failed_elements(failed_elements, element_type_name)
    
    # Return instance IDs for selection (empty list if no instances created)
    return success_count, fail_count, failed_elements, created_instance_ids

