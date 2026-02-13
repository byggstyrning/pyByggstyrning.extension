# -*- coding: utf-8 -*-
"""Clash detection utilities for element search, highlighting, and visualization.

This module provides fast GUID search across all Revit models (open + linked),
view creation, and clash result highlighting with color coding.
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List

from pyrevit import script
from pyrevit import revit

# Initialize logger
logger = script.get_logger()

def build_guid_lookup_dict(doc):
    """Build a fast lookup dictionary mapping IFC GUIDs to elements across ALL models.
    
    This function searches through both the open document and all linked models
    to create a comprehensive GUID lookup dictionary.
    
    Args:
        doc: The active Revit document
        
    Returns:
        Dictionary mapping IFC GUID strings to tuples of (element, is_linked, link_instance)
        where is_linked is True for linked elements and False for host model elements
    """
    guid_dict = {}
    
    try:
        logger.debug("Building GUID lookup dictionary for host model...")
        
        # First, collect all elements from the host model
        host_elements = FilteredElementCollector(doc)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        for element in host_elements:
            try:
                # Check for different variations of the IFC GUID parameter name
                ifc_guid_param = element.LookupParameter("IFCGuid")
                if not ifc_guid_param:
                    ifc_guid_param = element.LookupParameter("IfcGUID")
                if not ifc_guid_param:
                    ifc_guid_param = element.LookupParameter("IFC GUID")
                
                if ifc_guid_param and ifc_guid_param.HasValue and ifc_guid_param.StorageType == StorageType.String:
                    guid_value = ifc_guid_param.AsString()
                    if guid_value:
                        # Store as (element, is_linked=False, link_instance=None)
                        guid_dict[guid_value] = (element, False, None)
            except Exception as e:
                logger.debug("Error reading GUID from element {}: {}".format(element.Id, str(e)))
                continue
        
        logger.debug("Found {} elements with IFC GUIDs in host model".format(len(guid_dict)))
        
        # Now collect from all linked models
        link_instances = FilteredElementCollector(doc)\
            .OfClass(RevitLinkInstance)\
            .ToElements()
        
        logger.debug("Found {} linked models".format(len(link_instances)))
        
        for link_instance in link_instances:
            try:
                link_doc = link_instance.GetLinkDocument()
                if not link_doc:
                    logger.debug("Could not get document for link: {}".format(link_instance.Name))
                    continue
                
                logger.debug("Processing linked model: {}".format(link_instance.Name))
                
                # Collect elements from linked document
                link_elements = FilteredElementCollector(link_doc)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
                
                link_count = 0
                for element in link_elements:
                    try:
                        # Check for different variations of the IFC GUID parameter name
                        ifc_guid_param = element.LookupParameter("IFCGuid")
                        if not ifc_guid_param:
                            ifc_guid_param = element.LookupParameter("IfcGUID")
                        if not ifc_guid_param:
                            ifc_guid_param = element.LookupParameter("IFC GUID")
                        
                        if ifc_guid_param and ifc_guid_param.HasValue and ifc_guid_param.StorageType == StorageType.String:
                            guid_value = ifc_guid_param.AsString()
                            if guid_value:
                                # Store as (element, is_linked=True, link_instance)
                                # If GUID already exists, we keep the first one (prioritize host model)
                                if guid_value not in guid_dict:
                                    guid_dict[guid_value] = (element, True, link_instance)
                                    link_count += 1
                    except Exception as e:
                        logger.debug("Error reading GUID from linked element: {}".format(str(e)))
                        continue
                
                logger.debug("Found {} elements with IFC GUIDs in {}".format(link_count, link_instance.Name))
                
            except Exception as e:
                logger.error("Error processing linked model: {}".format(str(e)))
                continue
        
        logger.debug("Total elements in GUID lookup dictionary: {}".format(len(guid_dict)))
        return guid_dict
        
    except Exception as e:
        logger.error("Error building GUID lookup dictionary: {}".format(str(e)))
        return {}

def find_elements_by_guids(doc, guids):
    """Find Revit elements by their IFC GUIDs across all models.
    
    Args:
        doc: The active Revit document
        guids: List of IFC GUID strings to search for
        
    Returns:
        Dictionary mapping GUIDs to tuples of (element, is_linked, link_instance)
    """
    # Build the complete lookup dictionary
    guid_dict = build_guid_lookup_dict(doc)
    
    # Filter to only requested GUIDs
    found_elements = {}
    for guid in guids:
        if guid in guid_dict:
            found_elements[guid] = guid_dict[guid]
    
    return found_elements

def create_clash_view(doc, view_name, element_dict):
    """Create a new 3D view for visualizing clash results.
    
    Args:
        doc: The active Revit document
        view_name: Name for the new view
        element_dict: Dictionary of elements to show (from find_elements_by_guids)
        
    Returns:
        The created View3D or None if creation failed
    """
    try:
        logger.debug("Creating clash view: {}".format(view_name))
        
        # Find a 3D view type
        view_family_types = FilteredElementCollector(doc)\
            .OfClass(ViewFamilyType)\
            .ToElements()
        
        view_type = None
        for vft in view_family_types:
            if vft.ViewFamily == ViewFamily.ThreeDimensional:
                view_type = vft
                break
        
        if not view_type:
            logger.error("Could not find a 3D view type")
            return None
        
        # Create the 3D view
        with revit.Transaction("Create Clash View"):
            new_view = View3D.CreateIsometric(doc, view_type.Id)
            new_view.Name = view_name
            
            # Optionally set detail level
            new_view.DetailLevel = ViewDetailLevel.Fine
            
            logger.debug("Created 3D view: {}".format(view_name))
            return new_view
            
    except Exception as e:
        logger.error("Error creating clash view: {}".format(str(e)))
        return None

def highlight_clash_elements(doc, view, element_dict):
    """Highlight clash elements in a view with color coding.
    
    Own model elements are colored green, linked elements are colored red.
    Also creates section boxes around the elements for better visualization.
    
    Args:
        doc: The active Revit document
        view: The view to apply highlighting in
        element_dict: Dictionary mapping GUIDs to (element, is_linked, link_instance) tuples
        
    Returns:
        True if successful, False otherwise
    """
    try:
        logger.debug("Highlighting {} elements in view".format(len(element_dict)))
        
        # Define colors
        green_color = Color(0, 200, 0)  # Green for own elements
        red_color = Color(255, 0, 0)    # Red for linked elements
        
        # Get solid fill pattern
        solid_fill_id = None
        fillpatterns = FilteredElementCollector(doc).OfClass(FillPatternElement)
        for pat in fillpatterns:
            if pat.GetFillPattern().IsSolidFill:
                solid_fill_id = pat.Id
                break
        
        # Lists to track elements for section box calculation
        all_bboxes = []
        
        with revit.Transaction("Highlight Clash Elements"):
            for guid, (element, is_linked, link_instance) in element_dict.items():
                try:
                    # Choose color based on whether element is linked
                    color = red_color if is_linked else green_color
                    
                    # Create override settings
                    ogs = OverrideGraphicSettings()
                    ogs.SetProjectionLineColor(color)
                    ogs.SetCutLineColor(color)
                    
                    if solid_fill_id:
                        ogs.SetSurfaceForegroundPatternColor(color)
                        ogs.SetCutForegroundPatternColor(color)
                        ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                        ogs.SetCutForegroundPatternId(solid_fill_id)
                    
                    # Apply override
                    if is_linked and link_instance:
                        # For linked elements, we need to override through the link instance
                        view.SetElementOverrides(link_instance.Id, ogs)
                    else:
                        # For host model elements
                        view.SetElementOverrides(element.Id, ogs)
                    
                    # Get bounding box for section box calculation
                    try:
                        if is_linked and link_instance:
                            # For linked elements, transform the bounding box
                            bbox = element.get_BoundingBox(None)
                            if bbox:
                                transform = link_instance.GetTransform()
                                bbox_min = transform.OfPoint(bbox.Min)
                                bbox_max = transform.OfPoint(bbox.Max)
                                all_bboxes.append((bbox_min, bbox_max))
                        else:
                            bbox = element.get_BoundingBox(view)
                            if bbox:
                                all_bboxes.append((bbox.Min, bbox.Max))
                    except Exception as e:
                        logger.debug("Could not get bounding box for element: {}".format(str(e)))
                    
                except Exception as e:
                    logger.error("Error highlighting element with GUID {}: {}".format(guid, str(e)))
                    continue
            
            # Create section box around all elements
            if all_bboxes and isinstance(view, View3D):
                try:
                    # Calculate overall bounding box
                    min_x = min([bbox[0].X for bbox in all_bboxes])
                    min_y = min([bbox[0].Y for bbox in all_bboxes])
                    min_z = min([bbox[0].Z for bbox in all_bboxes])
                    max_x = max([bbox[1].X for bbox in all_bboxes])
                    max_y = max([bbox[1].Y for bbox in all_bboxes])
                    max_z = max([bbox[1].Z for bbox in all_bboxes])
                    
                    # Add some padding (10% of each dimension)
                    padding_x = (max_x - min_x) * 0.1
                    padding_y = (max_y - min_y) * 0.1
                    padding_z = (max_z - min_z) * 0.1
                    
                    # Create bounding box with padding
                    section_box = BoundingBoxXYZ()
                    section_box.Min = XYZ(min_x - padding_x, min_y - padding_y, min_z - padding_z)
                    section_box.Max = XYZ(max_x + padding_x, max_y + padding_y, max_z + padding_z)
                    
                    # Apply section box to view
                    view.SetSectionBox(section_box)
                    view.CropBoxActive = True
                    view.CropBoxVisible = True
                    
                    logger.debug("Applied section box to view")
                except Exception as e:
                    logger.error("Error creating section box: {}".format(str(e)))
        
        logger.debug("Successfully highlighted clash elements")
        return True
        
    except Exception as e:
        logger.error("Error highlighting clash elements: {}".format(str(e)))
        return False

def set_other_models_as_underlay(doc, view, element_dict):
    """Set models not involved in clash as underlay/halftone.
    
    Args:
        doc: The active Revit document
        view: The view to modify
        element_dict: Dictionary of clash elements
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Get list of link instances involved in clashes
        involved_links = set()
        for guid, (element, is_linked, link_instance) in element_dict.items():
            if is_linked and link_instance:
                involved_links.add(link_instance.Id)
        
        # Get all link instances in model
        all_links = FilteredElementCollector(doc)\
            .OfClass(RevitLinkInstance)\
            .ToElements()
        
        # Create override for underlay
        underlay_ogs = OverrideGraphicSettings()
        underlay_ogs.SetHalftone(True)
        
        with revit.Transaction("Set Underlay Models"):
            for link in all_links:
                if link.Id not in involved_links:
                    # This link is not involved in clash, set as halftone
                    view.SetElementOverrides(link.Id, underlay_ogs)
        
        logger.debug("Set uninvolved models as underlay")
        return True
        
    except Exception as e:
        logger.error("Error setting underlay models: {}".format(str(e)))
        return False

def enrich_clash_data_with_revit_info(doc, clash_results):
    """Enrich clash results with Revit category and level information.
    
    This function loads all GUIDs from clash results and matches them against
    the model to add Revit-specific metadata for filtering and grouping.
    
    Works with ifcclash format where clashes have:
    - a_global_id: IFC GUID for element A
    - b_global_id: IFC GUID for element B
    - a_ifc_class: IFC class name (e.g., "IfcWall")
    - b_ifc_class: IFC class name
    
    Args:
        doc: The active Revit document
        clash_results: List or dict of clash result objects with GUIDs
        
    Returns:
        Enriched clash results with Revit category and level information
    """
    try:
        logger.debug("Enriching clash data with Revit information...")
        
        # Build GUID lookup dictionary
        guid_dict = build_guid_lookup_dict(doc)
        
        # Handle both list and dict formats
        is_dict = isinstance(clash_results, dict)
        clashes_to_process = clash_results.values() if is_dict else clash_results
        
        # Process each clash result
        enriched_results = [] if not is_dict else {}
        
        for clash_key, clash in (clash_results.items() if is_dict else enumerate(clashes_to_process)):
            try:
                # Get GUIDs from ifcclash format
                guid_a = clash.get('a_global_id')
                guid_b = clash.get('b_global_id')
                
                # Look up elements
                element_a_data = guid_dict.get(guid_a)
                element_b_data = guid_dict.get(guid_b)
                
                # Add Revit metadata for element A
                if element_a_data:
                    element_a, is_linked_a, link_a = element_a_data
                    clash['revit_category_a'] = element_a.Category.Name if element_a.Category else clash.get('a_ifc_class', 'Unknown')
                    clash['revit_is_linked_a'] = is_linked_a
                    
                    # Get level
                    try:
                        level_param = element_a.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
                        if level_param and level_param.HasValue:
                            level_id = level_param.AsElementId()
                            level = doc.GetElement(level_id)
                            clash['revit_level_a'] = level.Name if level else "Unknown"
                        else:
                            clash['revit_level_a'] = "Unknown"
                    except:
                        clash['revit_level_a'] = "Unknown"
                else:
                    # Element not found in current model, use IFC class as fallback
                    clash['revit_category_a'] = clash.get('a_ifc_class', 'Not Found')
                    clash['revit_level_a'] = "Not Found"
                    clash['revit_is_linked_a'] = True
                
                # Add Revit metadata for element B
                if element_b_data:
                    element_b, is_linked_b, link_b = element_b_data
                    clash['revit_category_b'] = element_b.Category.Name if element_b.Category else clash.get('b_ifc_class', 'Unknown')
                    clash['revit_is_linked_b'] = is_linked_b
                    
                    # Get level
                    try:
                        level_param = element_b.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
                        if level_param and level_param.HasValue:
                            level_id = level_param.AsElementId()
                            level = doc.GetElement(level_id)
                            clash['revit_level_b'] = level.Name if level else "Unknown"
                        else:
                            clash['revit_level_b'] = "Unknown"
                    except:
                        clash['revit_level_b'] = "Unknown"
                else:
                    # Element not found in current model, use IFC class as fallback
                    clash['revit_category_b'] = clash.get('b_ifc_class', 'Not Found')
                    clash['revit_level_b'] = "Not Found"
                    clash['revit_is_linked_b'] = True
                
                # Store enriched result
                if is_dict:
                    enriched_results[clash_key] = clash
                else:
                    enriched_results.append(clash)
                
            except Exception as e:
                logger.error("Error enriching clash data: {}".format(str(e)))
                if is_dict:
                    enriched_results[clash_key] = clash
                else:
                    enriched_results.append(clash)
                continue
        
        logger.debug("Enriched {} clash results".format(len(enriched_results)))
        return enriched_results
        
    except Exception as e:
        logger.error("Error enriching clash data: {}".format(str(e)))
        return clash_results
