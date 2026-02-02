# -*- coding: utf-8 -*-
"""Main orchestration function for creating Mass elements from spatial elements.

This module is parallel to zone_creator.py but uses the Mass-specific
family creation API (NewExtrusionForm instead of NewExtrusion).
"""

import os.path as op
import tempfile
import shutil
from Autodesk.Revit.DB import (
    FilteredElementCollector, Family, Level
)
from pyrevit import script, forms, revit
from zone3d import mass_family_creation as mfc

logger = script.get_logger()


def report_failed_elements(failed_elements, element_type_name):
    """Report failed elements with linkify.
    
    Args:
        failed_elements: List of dicts with failed element info
        element_type_name: String like "Regions"
    """
    if not failed_elements:
        return
    
    output = script.get_output()
    output.print_md("## Failed {} - Invalid Geometry".format(element_type_name))
    output.print_md("**{} {}(s) could not be processed due to invalid geometry.**".format(
        len(failed_elements), element_type_name.rstrip('s')))
    output.print_md("Click {} ID to zoom to element in Revit".format(
        element_type_name.rstrip('s')))
    output.print_md("---")
    
    # Group by reason
    by_reason = {}
    for item in failed_elements:
        reason = item['reason']
        if reason not in by_reason:
            by_reason[reason] = []
        by_reason[reason].append(item)
    
    for reason, items in sorted(by_reason.items()):
        output.print_md("### {} ({} {})".format(reason, len(items), element_type_name.rstrip('s').lower()))
        output.print_md("")
        
        for item in items:
            element = item['element']
            element_number = item['element_number_str']
            element_name = item['element_name_str']
            element_link = output.linkify(element.Id)
            output.print_md("- **{} {} - {}**: {}".format(
                element_type_name.rstrip('s'), element_number, element_name, element_link))
        
        output.print_md("")
    
    output.print_md("---")
    output.print_md("**Tip:** Fix the {} boundaries and run the script again.".format(
        element_type_name.rstrip('s').lower()))


def create_masses_from_spatial_elements(
    spatial_elements,
    doc,
    adapter,
    extension_dir,
    pushbutton_dir,
    show_filter_dialog_func,
    element_type_name,
    template_family_name="MassZone.rfa"
):
    """Main orchestration function for creating Mass elements.
    
    Args:
        spatial_elements: List of spatial elements (FilledRegions, etc.)
        doc: Revit document
        adapter: SpatialElementAdapter instance
        extension_dir: Extension root directory
        pushbutton_dir: Pushbutton directory (for XAML files)
        show_filter_dialog_func: Function to show filter dialog
        element_type_name: String for progress/error messages (e.g., "Regions")
        template_family_name: Name of Mass template family file
        
    Returns:
        tuple: (success_count, fail_count, failed_elements, created_instance_ids)
    """
    app = doc.Application
    
    # Check if document supports families
    if doc.IsFamilyDocument:
        forms.alert("This tool can only be used in project documents, not family documents.",
                   title="Invalid Document Type", exitscript=True)
    
    # Verify Mass template exists
    template_path = mfc.get_mass_template_path(extension_dir, template_family_name)
    if not template_path:
        template_family_path = op.join(extension_dir, template_family_name)
        forms.alert("Mass template family '{}' not found at:\n{}\n\nPlease create the MassZone.rfa template in Revit using:\nFile > New > Conceptual Mass".format(
            template_family_name, template_family_path),
            title="Mass Template Not Found", exitscript=True)
    
    # Create temporary directory
    temp_dir = tempfile.mkdtemp(prefix="pyBS_Mass_")
    logger.debug("Created temporary directory: {}".format(temp_dir))
    
    # Show filter dialog
    selected_elements = show_filter_dialog_func(spatial_elements, doc)
    
    if not selected_elements:
        logger.debug("No {} selected, exiting.".format(element_type_name.lower()))
        script.exit()
    
    logger.debug("User selected {} {} for Mass creation".format(
        len(selected_elements), element_type_name.lower()))
    
    # Check if template family exists in project
    template_family_name_in_project = op.splitext(template_family_name)[0]
    project_template_family = None
    families_cache = FilteredElementCollector(doc).OfClass(Family).ToElements()
    for fam in families_cache:
        if fam.Name == template_family_name_in_project:
            project_template_family = fam
            logger.debug("Found Mass template family '{}' in project".format(
                template_family_name_in_project))
            break
    
    # Load template if not in project
    if not project_template_family:
        logger.debug("Mass template '{}' not found in project, loading...".format(
            template_family_name_in_project))
        try:
            from Autodesk.Revit.DB import IFamilyLoadOptions, FamilySource, SaveAsOptions
            
            class FamilyLoadOptions(IFamilyLoadOptions):
                def OnFamilyFound(self, familyInUse, overwriteParameterValues):
                    overwriteParameterValues[0] = True
                    return True
                
                def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
                    source[0] = FamilySource.Family
                    overwriteParameterValues[0] = True
                    return True
            
            load_options = FamilyLoadOptions()
            
            with revit.Transaction("Load Mass Template Family"):
                load_result = doc.LoadFamily(template_path, load_options)
                
                if isinstance(load_result, tuple):
                    project_template_family = load_result[1] if len(load_result) > 1 else None
                elif isinstance(load_result, bool):
                    if load_result:
                        families_cache = FilteredElementCollector(doc).OfClass(Family).ToElements()
                        project_template_family = next(
                            (f for f in families_cache if f.Name == template_family_name_in_project), None)
                else:
                    project_template_family = load_result
                
                if project_template_family:
                    logger.debug("Loaded Mass template family into project")
                else:
                    forms.alert("Failed to load Mass template family. Please check the file.",
                               title="Failed to Load Template", exitscript=True)
        except Exception as load_error:
            logger.error("Error loading Mass template: {}".format(load_error))
            forms.alert("Error loading Mass template:\n{}\n\nCheck logs for details.".format(
                str(load_error)), title="Error Loading Template", exitscript=True)
    
    # Export template for use
    template_family_doc = None
    template_path_to_use = template_path
    
    if project_template_family:
        try:
            from Autodesk.Revit.DB import SaveAsOptions
            
            logger.debug("Exporting Mass template from project...")
            temp_template_path = op.join(temp_dir, "temp_MassZone_template.rfa")
            
            template_family_doc = doc.EditFamily(project_template_family)
            if template_family_doc:
                save_options = SaveAsOptions()
                save_options.OverwriteExistingFile = True
                template_family_doc.SaveAs(temp_template_path, save_options)
                template_family_doc.Close(False)
                template_family_doc = None
                template_path_to_use = temp_template_path
                logger.debug("Exported Mass template to: {}".format(temp_template_path))
            else:
                logger.warning("Could not edit project family, using file template")
                project_template_family = None
        except Exception as export_error:
            logger.warning("Error exporting project family: {}".format(export_error))
            project_template_family = None
            if template_family_doc:
                try:
                    template_family_doc.Close(False)
                except:
                    pass
                template_family_doc = None
    
    # Open and inspect template
    logger.debug("Opening Mass template for inspection...")
    try:
        template_family_doc = app.OpenDocumentFile(template_path_to_use)
        if not template_family_doc:
            raise Exception("Failed to open Mass template document")
        
        template_info = mfc.inspect_mass_template(template_family_doc)
        
        logger.debug("Mass template inspection complete")
        
    except Exception as e:
        if template_family_doc:
            template_family_doc.Close(False)
        logger.error("Error opening Mass template: {}".format(e))
        forms.alert("Error opening Mass template: {}\n\nCheck logs for details.".format(str(e)),
                   title="Error", exitscript=True)
    
    # Process elements
    success_count = 0
    fail_count = 0
    failed_elements = []
    family_data_list = []
    created_instance_ids = []
    
    logger.debug("Starting to process {} {}...".format(len(selected_elements), element_type_name.lower()))
    
    # Cache levels
    levels_cache = FilteredElementCollector(doc).OfClass(Level).ToElements()
    levels_cache_sorted = sorted(levels_cache, key=lambda l: l.Elevation)
    
    # Cache families
    families_cache = FilteredElementCollector(doc).OfClass(Family).ToElements()
    
    # Phase 1: Process elements and create Mass families (0-50% progress)
    total_elements = len(selected_elements)
    with forms.ProgressBar(title="Creating Mass Families ({} {})".format(
            total_elements, element_type_name.lower())) as pb:
        
        for elem_idx, spatial_element in enumerate(selected_elements):
            try:
                element_id = spatial_element.Id.IntegerValue
                
                # Get element info
                element_number_str = adapter.get_number(spatial_element)
                if not element_number_str or element_number_str == "?":
                    element_number_str = str(element_id)
                
                element_name_str = adapter.get_name(spatial_element)
                
                logger.debug("Processing {} {} ({}/{}): {} - {}".format(
                    element_type_name.rstrip('s').lower(), element_number_str, elem_idx + 1,
                    len(selected_elements), element_name_str, element_id))
                
                # Update progress
                phase1_progress = int((elem_idx + 1) / float(total_elements) * 50)
                pb.update_progress(phase1_progress, 100)
                
                # Extract boundary data
                curves, insertion_point, height = mfc.extract_boundary_loops_for_mass(
                    spatial_element, doc, adapter, levels_cache_sorted)
                
                if not curves or not insertion_point:
                    logger.warning("{} {} - invalid boundary data, skipping".format(
                        element_type_name.rstrip('s'), element_number_str))
                    fail_count += 1
                    failed_elements.append({
                        'element': spatial_element,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'reason': 'Invalid boundary data'
                    })
                    continue
                
                if height is not None and height <= 0:
                    logger.warning("{} {} - invalid height, skipping".format(
                        element_type_name.rstrip('s'), element_number_str))
                    fail_count += 1
                    failed_elements.append({
                        'element': spatial_element,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'reason': 'Invalid height'
                    })
                    continue
                
                # Create family path
                sanitized_number = adapter.sanitize_number(element_number_str)
                family_name_prefix = "Mass_Region_"
                output_family_name = "{}{}_{}.rfa".format(family_name_prefix, sanitized_number, element_id)
                output_family_path = op.join(temp_dir, output_family_name)
                family_name_without_ext = op.splitext(output_family_name)[0]
                
                # Check if family exists
                existing_family = None
                for fam in families_cache:
                    try:
                        if fam and fam.Name == family_name_without_ext:
                            existing_family = fam
                            break
                    except:
                        continue
                
                # Delete existing family if found
                if existing_family:
                    logger.debug("Family '{}' exists - deleting and recreating".format(
                        family_name_without_ext))
                    success, error_reason = mfc.delete_mass_family_from_project(existing_family, doc)
                    if not success:
                        logger.warning("Failed to delete existing family: {}".format(error_reason))
                    existing_family = None
                
                # Create Mass family
                success, family_doc, error_reason = mfc.create_mass_family(
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
                    logger.warning("Failed to create Mass family for {} {}: {}".format(
                        element_type_name.rstrip('s').lower(), element_number_str, error_reason))
                    fail_count += 1
                    failed_elements.append({
                        'element': spatial_element,
                        'element_number_str': element_number_str,
                        'element_name_str': element_name_str,
                        'reason': error_reason or 'Failed to create Mass family'
                    })
                    if family_doc:
                        try:
                            family_doc.Close(False)
                        except:
                            pass
                    continue
                
                # Store data for second pass
                family_data_list.append({
                    'spatial_element': spatial_element,
                    'output_family_path': output_family_path,
                    'insertion_point': insertion_point,
                    'element_number_str': element_number_str,
                    'element_name_str': element_name_str,
                    'output_family_name': output_family_name,
                    'family_name_without_ext': family_name_without_ext,
                    'existing_family': None,
                    'family_doc': family_doc,
                    'height': height,
                    'template_info': template_info
                })
                
            except Exception as e:
                logger.error("Error processing {} {}: {}".format(
                    element_type_name.rstrip('s').lower(),
                    element_number_str if 'element_number_str' in locals() else "Unknown",
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
                            'reason': 'Error processing: {}'.format(str(e))
                        })
                except:
                    pass
        
        # Phase 2: Load families and place instances (50-100% progress)
        if family_data_list:
            def progress_callback(progress):
                pb.update_progress(progress, 100)
            
            mfc.load_mass_families(family_data_list, doc, app, progress_callback)
            
            place_success, place_fail, place_failed, created_instance_ids = mfc.place_mass_instances(
                family_data_list, doc, adapter, progress_callback)
            
            success_count += place_success
            fail_count += place_fail
            failed_elements.extend(place_failed)
    
    # Close template doc
    if template_family_doc:
        template_family_doc.Close(False)
    
    # Cleanup
    try:
        if op.exists(temp_dir):
            shutil.rmtree(temp_dir)
            logger.debug("Cleaned up temporary directory")
    except Exception as cleanup_error:
        logger.warning("Error cleaning up: {}".format(cleanup_error))
    
    # Log results
    logger.debug("Created {} Mass elements from {} {} ({} failed)".format(
        success_count, len(selected_elements), element_type_name, fail_count))
    
    # Report failures
    report_failed_elements(failed_elements, element_type_name)
    
    return success_count, fail_count, failed_elements, created_instance_ids
