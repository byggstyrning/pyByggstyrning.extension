# -*- coding: utf-8 -*-
"""Core mapping functions for 3D Zone parameter mapping."""

from collections import defaultdict
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction,
    StorageType
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.DB.Mechanical import Space
from pyrevit import revit, script
try:
    from zone3d import containment
    from zone3d import config
except ImportError:
    # Handle relative imports
    import containment
    import config

# Initialize logger
logger = script.get_logger()

def copy_parameter_value(source_param, target_param):
    """Copy a parameter value from source to target.
    
    Args:
        source_param: Source parameter
        target_param: Target parameter
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        if not source_param or not source_param.HasValue:
            return False
        
        if not target_param or target_param.IsReadOnly:
            return False
        
        # Check storage types match
        if source_param.StorageType != target_param.StorageType:
            return False
        
        storage_type = source_param.StorageType
        
        if storage_type == StorageType.String:
            value = source_param.AsString()
            if value:
                target_param.Set(value)
        elif storage_type == StorageType.Integer:
            value = source_param.AsInteger()
            target_param.Set(value)
        elif storage_type == StorageType.Double:
            value = source_param.AsDouble()
            target_param.Set(value)
        elif storage_type == StorageType.ElementId:
            value = source_param.AsElementId()
            if value and value != source_param.ElementId.InvalidElementId:
                target_param.Set(value)
        else:
            return False
        
        return True
    except Exception as e:
        return False

def has_source_parameter(element, source_param_names):
    """Check if an element has at least one of the source parameters with a value.
    
    Checks both instance and type parameters, and verifies the parameter has a value
    (not empty/not set). Empty strings are considered as "no value".
    
    Args:
        element: Element to check
        source_param_names: List of parameter names to check
        
    Returns:
        bool: True if element has at least one source parameter with a value
    """
    if not source_param_names:
        return False
    
    for param_name in source_param_names:
        try:
            # Check instance parameter
            param = element.LookupParameter(param_name)
            if param and param.HasValue:
                # Check if value is not empty (for strings)
                if param.StorageType == StorageType.String:
                    value = param.AsString()
                    if value and value.strip():  # Not empty or whitespace
                        return True
                else:
                    # For non-string types, HasValue is sufficient
                    return True
            
            # Check type parameter
            if hasattr(element, "GetTypeId"):
                type_id = element.GetTypeId()
                if type_id and type_id != element.Id.InvalidElementId:
                    element_type = element.Document.GetElement(type_id)
                    if element_type:
                        param = element_type.LookupParameter(param_name)
                        if param and param.HasValue:
                            # Check if value is not empty (for strings)
                            if param.StorageType == StorageType.String:
                                value = param.AsString()
                                if value and value.strip():  # Not empty or whitespace
                                    return True
                            else:
                                # For non-string types, HasValue is sufficient
                                return True
        except:
            continue
    
    return False

def copy_parameters(source_element, target_element, source_param_names, target_param_names):
    """Copy multiple parameters from source to target element.
    
    Args:
        source_element: Source element
        target_element: Target element
        source_param_names: List of source parameter names
        target_param_names: List of target parameter names (must match source count)
        
    Returns:
        int: Number of parameters successfully copied
    """
    if len(source_param_names) != len(target_param_names):
        logger.warning("Parameter name lists don't match in length")
        return 0
    
    copied_count = 0
    
    for source_name, target_name in zip(source_param_names, target_param_names):
        try:
            # Get source parameter
            source_param = source_element.LookupParameter(source_name)
            if not source_param:
                # Try element type parameter
                if hasattr(source_element, "GetTypeId"):
                    type_id = source_element.GetTypeId()
                    if type_id and type_id != source_element.Id.InvalidElementId:
                        element_type = source_element.Document.GetElement(type_id)
                        if element_type:
                            source_param = element_type.LookupParameter(source_name)
            
            if not source_param:
                logger.debug("[DEBUG] Source parameter '{}' not found on element {} (ID: {})".format(
                    source_name, source_element.GetType().Name, source_element.Id))
                continue
            
            if not source_param.HasValue:
                logger.debug("[DEBUG] Source parameter '{}' found but has no value on element {} (ID: {})".format(
                    source_name, source_element.GetType().Name, source_element.Id))
                continue
            
            # Get target parameter
            target_param = target_element.LookupParameter(target_name)
            if not target_param:
                # Try element type parameter
                if hasattr(target_element, "GetTypeId"):
                    type_id = target_element.GetTypeId()
                    if type_id and type_id != target_element.Id.InvalidElementId:
                        element_type = target_element.Document.GetElement(type_id)
                        if element_type:
                            target_param = element_type.LookupParameter(target_name)
            
            if not target_param:
                logger.debug("[DEBUG] Target parameter '{}' not found on element {} (ID: {})".format(
                    target_name, target_element.GetType().Name, target_element.Id))
                continue
            
            if target_param.IsReadOnly:
                logger.debug("[DEBUG] Target parameter '{}' is read-only on element {} (ID: {})".format(
                    target_name, target_element.GetType().Name, target_element.Id))
                continue
            
            # Copy the value
            if copy_parameter_value(source_param, target_param):
                copied_count += 1
                logger.debug("[DEBUG] Successfully copied parameter '{}' -> '{}' (source: {} ID: {}, target: {} ID: {})".format(
                    source_name, target_name, 
                    source_element.GetType().Name, source_element.Id,
                    target_element.GetType().Name, target_element.Id))
            else:
                logger.debug("[DEBUG] Failed to copy parameter '{}' -> '{}' (copy_parameter_value returned False)".format(
                    source_name, target_name))
        
        except Exception as e:
            logger.debug("[DEBUG] Exception copying parameter '{}' -> '{}': {}".format(source_name, target_name, str(e)))
            continue
    
    return copied_count

def write_parameters_to_elements(doc, zone_config, progress_bar=None):
    """Write parameters to elements based on a zone configuration.
    
    Args:
        doc: Revit document
        zone_config: Configuration dictionary
        progress_bar: Optional progress bar for tracking progress
        
    Returns:
        dict: Results dictionary with counts and errors
    """
    results = {
        "elements_processed": 0,
        "elements_updated": 0,
        "parameters_copied": 0,
        "errors": []
    }
    
    try:
        # Detect containment strategy
        source_categories = zone_config.get("source_categories", [])
        strategy = containment.detect_containment_strategy(source_categories)
        
        if not strategy:
            error_msg = "Could not detect containment strategy for categories: {}".format(source_categories)
            logger.error(error_msg)
            results["errors"].append(error_msg)
            return results
        
        # Get parameter names early (needed for filtering)
        source_param_names = zone_config.get("source_params", [])
        
        if not source_param_names:
            return results
        
        # Get source elements
        # Collect elements from each category separately and combine (OR logic)
        # Multiple OfCategory() calls create AND logic (elements in ALL categories), which is wrong
        if source_categories:
            source_elements = []
            element_ids = set()  # Track IDs to avoid duplicates
            for category in source_categories:
                category_elements = FilteredElementCollector(doc)\
                    .WhereElementIsNotElementType()\
                    .OfCategory(category)\
                    .ToElements()
                for el in category_elements:
                    if el.Id.IntegerValue not in element_ids:
                        element_ids.add(el.Id.IntegerValue)
                        source_elements.append(el)
        else:
            source_elements = FilteredElementCollector(doc)\
                .WhereElementIsNotElementType()\
                .ToElements()
        
        logger.info("[DEBUG] Found {} total source elements (before filtering)".format(len(source_elements)))
        
        if not source_elements:
            logger.warning("[DEBUG] No source elements found at all for categories: {}".format(source_categories))
            return results
        
        # Filter source elements to only include those with source parameters
        # This avoids processing elements that can't provide values anyway
        filtered_source_elements = []
        elements_without_params = []
        for source_el in source_elements:
            if has_source_parameter(source_el, source_param_names):
                filtered_source_elements.append(source_el)
            else:
                elements_without_params.append(source_el.Id)
        
        logger.info("[DEBUG] Source elements with parameters '{}': {}".format(source_param_names, len(filtered_source_elements)))
        if elements_without_params:
            logger.info("[DEBUG] Source elements WITHOUT parameters (first 10): {}".format(elements_without_params[:10]))
        
        source_elements = filtered_source_elements
        
        if not source_elements:
            logger.warning("[DEBUG] No source elements found with required parameters: {}".format(source_param_names))
            return results
        
        # Pre-compute geometries for Mass/Generic Model elements and Areas (batch operation)
        # This improves performance by calculating all geometries in one pass
        if strategy in ["element", "area"]:
            containment.precompute_geometries(source_elements, doc)
        
        # Group source elements by level for performance
        if strategy in ["room", "space"]:
            source_by_level = defaultdict(list)
            for source_el in source_elements:
                if hasattr(source_el, "LevelId") and source_el.LevelId:
                    source_by_level[source_el.LevelId].append(source_el)
        else:
            source_by_level = None
        
        # Get target elements
        target_filter_categories = zone_config.get("target_filter_categories", [])
        
        # Collect elements from each category separately and combine (OR logic)
        # Multiple OfCategory() calls create AND logic (elements in ALL categories), which is wrong
        if target_filter_categories:
            target_elements = []
            element_ids = set()  # Track IDs to avoid duplicates
            for category in target_filter_categories:
                category_elements = FilteredElementCollector(doc)\
                    .WhereElementIsNotElementType()\
                    .OfCategory(category)\
                    .ToElements()
                for el in category_elements:
                    if el.Id.IntegerValue not in element_ids:
                        element_ids.add(el.Id.IntegerValue)
                        target_elements.append(el)
        else:
            target_elements = FilteredElementCollector(doc)\
                .WhereElementIsNotElementType()\
                .ToElements()
        
        logger.info("[DEBUG] Found {} target elements for categories: {}".format(len(target_elements), target_filter_categories))
        
        if not target_elements:
            logger.warning("[DEBUG] No target elements found for categories: {}".format(target_filter_categories))
            return results
        
        # Get target parameter names (source_param_names already retrieved above)
        target_param_names = zone_config.get("target_params", [])
        
        if len(source_param_names) != len(target_param_names):
            error_msg = "Source and target parameter lists don't match"
            logger.error(error_msg)
            results["errors"].append(error_msg)
            return results
        
        # Process elements (element-driven loop for performance)
        elements_updated = 0
        total_params_copied = 0
        
        # Pre-group rooms/spaces/areas by level if needed
        rooms_by_level = None
        spaces_by_level = None
        areas_by_level = None
        
        if strategy == "room":
            rooms_by_level = defaultdict(list)
            for source_el in source_elements:
                if isinstance(source_el, Room) and source_el.LevelId:
                    rooms_by_level[source_el.LevelId].append(source_el)
        elif strategy == "space":
            spaces_by_level = defaultdict(list)
            for source_el in source_elements:
                if isinstance(source_el, Space) and source_el.LevelId:
                    spaces_by_level[source_el.LevelId].append(source_el)
        elif strategy == "area":
            from Autodesk.Revit.DB import Area
            areas_by_level = defaultdict(list)
            for source_el in source_elements:
                if isinstance(source_el, Area):
                    # Areas may have LevelId or get it from parameter
                    level_id = None
                    if hasattr(source_el, "LevelId") and source_el.LevelId:
                        level_id = source_el.LevelId
                    elif hasattr(source_el, "get_Parameter"):
                        level_param = source_el.get_Parameter("Level")
                        if level_param:
                            level_id = level_param.AsElementId()
                    
                    if level_id:
                        areas_by_level[level_id].append(source_el)
                    else:
                        # If no level, add to a default list (use None as key)
                        areas_by_level[None].append(source_el)
        
        # Process each target element
        total_elements = len(target_elements)
        # Calculate update interval for 5% increments (20 steps = 5% each)
        update_interval = max(1, int(total_elements / 20.0)) if total_elements > 0 else 1
        last_update = 0
        
        # Debug counters
        containment_found_count = 0
        containment_not_found_count = 0
        params_copy_failed_count = 0
        
        logger.info("[DEBUG] Starting to process {} target elements using strategy '{}'".format(total_elements, strategy))
        
        for idx, target_el in enumerate(target_elements):
            try:
                results["elements_processed"] += 1
                
                # Update progress bar in 5% increments
                if progress_bar and total_elements > 0:
                    current_progress = idx + 1
                    if current_progress - last_update >= update_interval or current_progress == total_elements:
                        progress_bar.update_progress(current_progress, total_elements)
                        last_update = current_progress
                
                # Find containing element
                containing_el = containment.get_containing_element_by_strategy(
                    target_el, doc, strategy, source_categories,
                    rooms_by_level, spaces_by_level, areas_by_level
                )
                
                if not containing_el:
                    containment_not_found_count += 1
                    # Log first few failures for debugging
                    if containment_not_found_count <= 5:
                        logger.debug("[DEBUG] Target element {} (ID: {}) - no containing element found".format(
                            idx + 1, target_el.Id))
                    continue
                
                containment_found_count += 1
                
                # Log first few successes for debugging
                if containment_found_count <= 5:
                    logger.debug("[DEBUG] Target element {} (ID: {}) - found containing element {} (ID: {})".format(
                        idx + 1, target_el.Id, containing_el.GetType().Name, containing_el.Id))
                
                # Copy parameters
                params_copied = copy_parameters(
                    containing_el, target_el,
                    source_param_names, target_param_names
                )
                
                if params_copied > 0:
                    elements_updated += 1
                    total_params_copied += params_copied
                    if elements_updated <= 5:
                        logger.debug("[DEBUG] Successfully copied {} parameters to element {} (ID: {})".format(
                            params_copied, idx + 1, target_el.Id))
                else:
                    params_copy_failed_count += 1
                    if params_copy_failed_count <= 5:
                        logger.debug("[DEBUG] Failed to copy parameters to element {} (ID: {}) - containing element found but params not copied".format(
                            idx + 1, target_el.Id))
            
            except Exception as e:
                error_msg = "Error processing element {}: {}".format(target_el.Id, str(e))
                logger.error(error_msg)
                results["errors"].append(error_msg)
                continue
        
        # Log summary statistics
        logger.info("[DEBUG] Processing complete:")
        logger.info("[DEBUG]   Containment found: {} / {} target elements".format(containment_found_count, total_elements))
        logger.info("[DEBUG]   Containment not found: {} / {} target elements".format(containment_not_found_count, total_elements))
        logger.info("[DEBUG]   Parameters copied successfully: {} elements".format(elements_updated))
        logger.info("[DEBUG]   Parameters copy failed (containment found but no params copied): {} elements".format(params_copy_failed_count))
        
        
        results["elements_updated"] = elements_updated
        results["parameters_copied"] = total_params_copied
        
        return results
    
    except Exception as e:
        error_msg = "Error in write_parameters_to_elements: {}".format(str(e))
        logger.error(error_msg)
        results["errors"].append(error_msg)
        return results

def execute_configuration(doc, zone_config, progress_bar=None):
    """Execute a single configuration within a transaction.
    
    Args:
        doc: Revit document
        zone_config: Configuration dictionary
        progress_bar: Optional progress bar for tracking progress
        
    Returns:
        dict: Results dictionary
    """
    config_name = zone_config.get("name", "Unknown")
    
    with revit.Transaction("3D Zone: {}".format(config_name), doc):
        return write_parameters_to_elements(doc, zone_config, progress_bar)

def execute_all_configurations(doc):
    """Execute all enabled configurations in order.
    
    Args:
        doc: Revit document
        
    Returns:
        dict: Summary results with per-configuration results
    """
    from pyrevit import forms
    
    configs = config.get_enabled_configs(doc)
    
    if not configs:
        logger.info("No enabled configurations found")
        return {
            "total_configs": 0,
            "config_results": [],
            "total_elements_updated": 0,
            "total_parameters_copied": 0
        }
    
    summary = {
        "total_configs": len(configs),
        "config_results": [],
        "total_elements_updated": 0,
        "total_parameters_copied": 0
    }
    
    # Process each configuration with its own progress bar
    for config_idx, zone_config in enumerate(configs):
        config_name = zone_config.get("name", "Unknown")
        config_order = zone_config.get("order", 0)
        
        # Create progress bar with configuration name
        progress_title = "3D Zone: {} ({} of {})".format(
            config_name, config_idx + 1, len(configs)
        )
        
        with forms.ProgressBar(title=progress_title) as pb:
            # Reset progress bar to 0
            pb.update_progress(0, 100)
            
            # Execute configuration with progress bar
            result = execute_configuration(doc, zone_config, pb)
            result["config_name"] = config_name
            result["config_order"] = config_order
            
            summary["config_results"].append(result)
            summary["total_elements_updated"] += result.get("elements_updated", 0)
            summary["total_parameters_copied"] += result.get("parameters_copied", 0)
    
    # Log summary
    logger.info("3D Zone Write Complete: {} configs, {} elements updated, {} parameters copied".format(
        summary["total_configs"],
        summary["total_elements_updated"],
        summary["total_parameters_copied"]
    ))
    
    return summary

