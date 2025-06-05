# -*- coding: utf-8 -*-
"""Sync checker module for MMI parameter validation after document synchronization."""

import datetime
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import LocationPoint, LocationCurve
from Autodesk.Revit.UI import UIDocument
from pyrevit import script, forms, revit
from System.Collections.Generic import List

# Initialize logger
logger = script.get_logger()

# Global storage for pre-sync element state
_user_modified_elements = {}
_sync_in_progress = False

def get_user_owned_elements(doc):
    """Get elements owned by current user using WorksharingUtils.
    
    Args:
        doc: The active Revit document
        
    Returns:
        list: List of element IDs owned by current user
    """
    try:
        if not doc.IsWorkshared:
            logger.debug("Document is not workshared, returning empty list")
            return []
            
        user_owned_elements = []
        
        # Get all elements in the model
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        
        for element in all_elements:
            try:
                # Skip non-model elements (views, annotations, etc.)
                if not element.Category or element.Category.CategoryType != CategoryType.Model:
                    continue
                    
                # Get worksharing tooltip info to determine ownership
                tooltip_info = WorksharingUtils.GetWorksharingTooltipInfo(doc, element.Id)
                
                # Check if current user owns this element
                if tooltip_info and tooltip_info.Owner and tooltip_info.Owner.lower() == doc.Application.Username.lower():
                    user_owned_elements.append(element.Id)
                    
            except Exception as e:
                # Some elements might not support worksharing queries
                logger.debug("Error checking ownership for element {}: {}".format(element.Id, e))
                continue
                
        logger.debug("Found {} user-owned elements".format(len(user_owned_elements)))
        return user_owned_elements
        
    except Exception as e:
        logger.error("Error getting user owned elements: {}".format(e))
        return []

def track_modified_elements_before_sync(doc):
    """Store modified elements before sync to compare after.
    
    Args:
        doc: The active Revit document
    """
    global _user_modified_elements, _sync_in_progress
    
    try:
        _sync_in_progress = True
        
        # Get user-owned elements
        user_owned_ids = get_user_owned_elements(doc)
        
        # Store current state of user-owned elements
        _user_modified_elements = {}
        
        # Get MMI parameter name
        from mmi.core import get_mmi_parameter_name
        from mmi.utils import get_element_mmi_value
        
        mmi_param_name = get_mmi_parameter_name(doc)
        
        for element_id in user_owned_ids:
            try:
                element = doc.GetElement(element_id)
                if element:
                    # Double-check that this is a model element
                    if not element.Category or element.Category.CategoryType != CategoryType.Model:
                        continue
                        
                    # Store current MMI state
                    mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
                    _user_modified_elements[element_id.IntegerValue] = {
                        "element_id": element_id,
                        "mmi_value": mmi_value,
                        "mmi_string": value_str,
                        "has_mmi_param": param is not None,
                        "timestamp": datetime.datetime.now()
                    }
                    
            except Exception as e:
                logger.debug("Error tracking element {}: {}".format(element_id, e))
                
        logger.debug("Tracked {} user-owned elements before sync".format(len(_user_modified_elements)))
        
    except Exception as e:
        logger.error("Error tracking elements before sync: {}".format(e))

def validate_post_sync_mmi(doc):
    """Validate MMI on user's modified elements after sync.
    
    Args:
        doc: The active Revit document
        
    Returns:
        dict: Results of post-sync validation
    """
    global _user_modified_elements, _sync_in_progress
    
    try:
        if not _sync_in_progress or not _user_modified_elements:
            logger.debug("No sync in progress or no tracked elements")
            return {"elements_missing_mmi": [], "elements_invalid_mmi": [], "total_checked": 0}
            
        # Reset sync flag
        _sync_in_progress = False
        
        # Get MMI parameter name
        from mmi.core import get_mmi_parameter_name
        from mmi.utils import get_element_mmi_value, validate_mmi_value
        
        mmi_param_name = get_mmi_parameter_name(doc)
        
        elements_missing_mmi = []
        elements_invalid_mmi = []
        
        for element_int_id, pre_sync_data in _user_modified_elements.items():
            try:
                element_id = ElementId(element_int_id)
                element = doc.GetElement(element_id)
                
                if not element:
                    continue
                
                # Skip non-model elements
                if not element.Category or element.Category.CategoryType != CategoryType.Model:
                    continue
                    
                # Check current MMI state
                current_mmi_value, current_value_str, current_param = get_element_mmi_value(
                    element, mmi_param_name, doc)
                
                # Check if element is missing MMI value
                if current_param is None or not current_value_str or current_value_str.strip() == "":
                    elements_missing_mmi.append({
                        "element_id": element_id,
                        "element": element,
                        "category": element.Category.Name if element.Category else "Unknown",
                        "pre_sync_mmi": pre_sync_data.get("mmi_string", "None")
                    })
                    logger.debug("Element {} missing MMI after sync".format(element_id))
                    
                # Check if MMI value is invalid
                elif current_value_str:
                    original, fixed = validate_mmi_value(current_value_str)
                    if original and fixed:
                        elements_invalid_mmi.append({
                            "element_id": element_id,
                            "element": element,
                            "category": element.Category.Name if element.Category else "Unknown",
                            "current_value": current_value_str,
                            "suggested_value": fixed
                        })
                        logger.debug("Element {} has invalid MMI: '{}'".format(element_id, current_value_str))
                        
            except Exception as e:
                logger.debug("Error validating element {}: {}".format(element_int_id, e))
                
        results = {
            "elements_missing_mmi": elements_missing_mmi,
            "elements_invalid_mmi": elements_invalid_mmi,
            "total_checked": len(_user_modified_elements)
        }
        
        # Clear tracked elements
        _user_modified_elements = {}
        
        logger.debug("Post-sync validation complete: {} missing, {} invalid out of {} checked".format(
            len(elements_missing_mmi), len(elements_invalid_mmi), results["total_checked"]))
            
        return results
        
    except Exception as e:
        logger.error("Error in post-sync MMI validation: {}".format(e))
        return {"elements_missing_mmi": [], "elements_invalid_mmi": [], "total_checked": 0}

def create_post_sync_view(doc, problematic_elements):
    """Create/update a special 3D view showing elements with missing/invalid MMI.
    
    Args:
        doc: The active Revit document
        problematic_elements: List of elements that need attention
        
    Returns:
        View3D: The created/updated view or None if failed
    """
    try:
        if not problematic_elements:
            logger.debug("No problematic elements to show")
            return None
            
        VIEW_NAME = "Post-Sync MMI Check"
        
        # Find or create the view
        post_sync_view = None
        collector = FilteredElementCollector(doc).OfClass(View3D)
        
        for view in collector:
            if not view.IsTemplate and view.Name == VIEW_NAME:
                post_sync_view = view
                break
                
        if not post_sync_view:
            # Create new view
            with revit.Transaction("Create Post-Sync MMI View", doc):
                vft_collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
                vft_3d = next((vft for vft in vft_collector 
                              if vft.ViewFamily == ViewFamily.ThreeDimensional), None)
                              
                if not vft_3d:
                    logger.error("Could not find a 3D ViewFamilyType")
                    return None
                    
                post_sync_view = View3D.CreateIsometric(doc, vft_3d.Id)
                post_sync_view.Name = VIEW_NAME
                logger.debug("Created new post-sync MMI view")
        
        # Calculate bounding box for all problematic elements
        all_element_ids = []
        min_pt = None
        max_pt = None
        
        for item in problematic_elements:
            element = item["element"]
            
            # Skip if not a model element (safety check)
            if not element.Category or element.Category.CategoryType != CategoryType.Model:
                continue
                
            all_element_ids.append(element.Id)
            
            # Get element bounding box
            try:
                bbox = element.get_BoundingBox(post_sync_view)
                if bbox:
                    if min_pt is None:
                        min_pt = bbox.Min
                        max_pt = bbox.Max
                    else:
                        min_pt = XYZ(
                            min(min_pt.X, bbox.Min.X),
                            min(min_pt.Y, bbox.Min.Y), 
                            min(min_pt.Z, bbox.Min.Z)
                        )
                        max_pt = XYZ(
                            max(max_pt.X, bbox.Max.X),
                            max(max_pt.Y, bbox.Max.Y),
                            max(max_pt.Z, bbox.Max.Z)
                        )
            except Exception as e:
                logger.debug("Error getting bounding box for element {}: {}".format(element.Id, e))
        
        # Apply section box and select elements
        with revit.Transaction("Update Post-Sync MMI View", doc):
            if min_pt and max_pt:
                # Add padding to bounding box
                padding = 3.0
                padded_min = min_pt - XYZ(padding, padding, padding)
                padded_max = max_pt + XYZ(padding, padding, padding)
                
                section_box = BoundingBoxXYZ()
                section_box.Min = padded_min
                section_box.Max = padded_max
                
                post_sync_view.IsSectionBoxActive = True
                post_sync_view.SetSectionBox(section_box)
                logger.debug("Applied section box to post-sync view")
        
        # Activate view and select elements
        try:
            uidoc = UIDocument(doc)
            uidoc.ActiveView = post_sync_view
            
            if all_element_ids:
                uidoc.Selection.SetElementIds(List[ElementId](all_element_ids))
                # Zoom to selection
                if len(all_element_ids) == 1:
                    uidoc.ShowElements(all_element_ids[0])
                else:
                    uidoc.ShowElements(List[ElementId](all_element_ids))
                    
            logger.debug("Activated post-sync view and selected {} elements".format(len(all_element_ids)))
            
        except Exception as e:
            logger.warning("Error activating view or selecting elements: {}".format(e))
            
        return post_sync_view
        
    except Exception as e:
        logger.error("Error creating post-sync view: {}".format(e))
        return None

def process_post_sync_check(doc):
    """Main function to process post-sync MMI check.
    
    Args:
        doc: The active Revit document
    """
    try:
        # Validate MMI after sync
        results = validate_post_sync_mmi(doc)
        
        total_issues = len(results["elements_missing_mmi"]) + len(results["elements_invalid_mmi"])
        
        if total_issues == 0:
            if results["total_checked"] > 0:
                # Show success message
                forms.show_balloon(
                    header="Post-Sync MMI Check",
                    text="✅ All {} checked elements have valid MMI values".format(results["total_checked"]),
                    tooltip="No MMI issues found after synchronization",
                    is_new=True
                )
            return
            
        # Combine all problematic elements
        all_problematic = results["elements_missing_mmi"] + results["elements_invalid_mmi"]
        
        # Create/update view
        view = create_post_sync_view(doc, all_problematic)
        
        # Create detailed message
        missing_count = len(results["elements_missing_mmi"])
        invalid_count = len(results["elements_invalid_mmi"])
        
        message_parts = []
        if missing_count > 0:
            message_parts.append("{} missing MMI".format(missing_count))
        if invalid_count > 0:
            message_parts.append("{} invalid MMI".format(invalid_count))
            
        message = "Found: " + ", ".join(message_parts)
        
        # Create tooltip with details
        tooltip_lines = ["Post-sync MMI validation results:"]
        if missing_count > 0:
            tooltip_lines.append("Missing MMI:")
            for item in results["elements_missing_mmi"][:3]:  # Show first 3
                tooltip_lines.append("  • {} (ID: {})".format(
                    item["category"], item["element_id"].IntegerValue))
            if missing_count > 3:
                tooltip_lines.append("  • ... and {} more".format(missing_count - 3))
                
        if invalid_count > 0:
            tooltip_lines.append("Invalid MMI:")
            for item in results["elements_invalid_mmi"][:3]:  # Show first 3
                tooltip_lines.append("  • {} '{}' → '{}'".format(
                    item["category"], item["current_value"], item["suggested_value"]))
            if invalid_count > 3:
                tooltip_lines.append("  • ... and {} more".format(invalid_count - 3))
        
        tooltip = "\n".join(tooltip_lines)
        
        # Show notification
        forms.show_balloon(
            header="Post-Sync MMI Check",
            text=message + "\nView: '{}'".format(view.Name if view else "Creation failed"),
            tooltip=tooltip,
            is_new=True
        )
        
        logger.debug("Post-sync check completed: {} total issues found".format(total_issues))
        
    except Exception as e:
        logger.error("Error in post-sync check process: {}".format(e)) 