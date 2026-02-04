# -*- coding: utf-8 -*-
"""Core mapping functions for 3D Zone parameter mapping."""

from collections import defaultdict
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, Transaction,
    StorageType, BuiltInParameter, Category, RevitLinkInstance, ElementId
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.DB.Mechanical import Space
from pyrevit import revit, script
import time
try:
    from zone3d import containment
    from zone3d import config
    import sys
    import os.path as op
    # Add lib path for revit_utils import
    script_path = __file__
    lib_dir = op.dirname(op.dirname(op.dirname(script_path)))
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)
    from revit.revit_utils import is_element_editable
except ImportError:
    # Handle relative imports
    import containment
    import config
    import sys
    import os.path as op
    # Add lib path for revit_utils import
    script_path = __file__
    lib_dir = op.dirname(op.dirname(op.dirname(script_path)))
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)
    try:
        from revit.revit_utils import is_element_editable
    except ImportError:
        # Fallback if import fails
        def is_element_editable(doc, element):
            return True, "Editable"

# Initialize logger
logger = script.get_logger()

def copy_parameter_value(source_param, target_param, return_value=False):
    """Copy a parameter value from source to target.
    
    Optimized to skip writes when values haven't changed (reduces instance-level writes).
    
    Args:
        source_param: Source parameter
        target_param: Target parameter
        return_value: If True, return (success, value) tuple instead of just bool
        
    Returns:
        bool or tuple: True if successful, False otherwise. If return_value=True, returns (success, value) tuple.
    """
    try:
        if not source_param or not source_param.HasValue:
            return (False, None) if return_value else False
        
        if not target_param or target_param.IsReadOnly:
            return (False, None) if return_value else False
        
        # Check storage types match
        if source_param.StorageType != target_param.StorageType:
            return (False, None) if return_value else False
        
        storage_type = source_param.StorageType
        value = None
        
        if storage_type == StorageType.String:
            value = source_param.AsString()
            if value:
                # Skip write if value hasn't changed
                if target_param.HasValue:
                    current_value = target_param.AsString()
                    if current_value == value:
                        return (False, value) if return_value else False  # Value unchanged, skip write
                target_param.Set(value)
        elif storage_type == StorageType.Integer:
            value = source_param.AsInteger()
            # Skip write if value hasn't changed
            if target_param.HasValue:
                current_value = target_param.AsInteger()
                if current_value == value:
                    return (False, value) if return_value else False  # Value unchanged, skip write
            target_param.Set(value)
        elif storage_type == StorageType.Double:
            value = source_param.AsDouble()
            # Skip write if value hasn't changed (with tolerance for floating point)
            if target_param.HasValue:
                current_value = target_param.AsDouble()
                if abs(current_value - value) < 1e-9:  # Very small tolerance for floating point
                    return (False, value) if return_value else False  # Value unchanged, skip write
            target_param.Set(value)
        elif storage_type == StorageType.ElementId:
            value = source_param.AsElementId()
            if value and value != source_param.ElementId.InvalidElementId:
                # Skip write if value hasn't changed
                if target_param.HasValue:
                    current_value = target_param.AsElementId()
                    if current_value == value:
                        return (False, value) if return_value else False  # Value unchanged, skip write
                target_param.Set(value)
                # Convert ElementId to integer for caching
                value = value.IntegerValue if value else None
        else:
            return (False, None) if return_value else False
        
        return (True, value) if return_value else True
    except Exception as e:
        return (False, None) if return_value else False

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

def has_target_parameter(element, target_param_names):
    """Check if an element has at least one of the target parameters (exists and is writable).
    
    Checks both instance and type parameters. The parameter must exist and not be read-only.
    We don't check for values since we're writing TO these parameters.
    
    Args:
        element: Element to check
        target_param_names: List of parameter names to check
        
    Returns:
        bool: True if element has at least one target parameter that exists and is writable
    """
    if not target_param_names:
        return False
    
    for param_name in target_param_names:
        try:
            # Check instance parameter
            param = element.LookupParameter(param_name)
            if param and not param.IsReadOnly:
                return True
            
            # Check type parameter
            if hasattr(element, "GetTypeId"):
                type_id = element.GetTypeId()
                if type_id and type_id != element.Id.InvalidElementId:
                    element_type = element.Document.GetElement(type_id)
                    if element_type:
                        param = element_type.LookupParameter(param_name)
                        if param and not param.IsReadOnly:
                            return True
        except:
            continue
    
    return False

def sort_source_elements(source_elements, sort_property="ElementId", descending=False):
    """Sort source elements by a property value.
    
    Elements with empty/missing sort property values come first,
    then elements with values are sorted alphabetically/numerically.
    
    Args:
        source_elements: List of source elements to sort
        sort_property: Name of property to sort by (default: "ElementId")
        descending: If True, sort in descending order (default: False)
        
    Returns:
        list: Sorted list of source elements
    """
    if not source_elements:
        return []
    
    # If sorting by ElementId, use simple integer value sort
    if sort_property == "ElementId":
        return sorted(source_elements, key=lambda el: el.Id.IntegerValue, reverse=descending)
    
    def get_sort_value(element):
        """Get sort value from element property.
        
        Returns tuple (has_value, sort_value, element_id) where:
        - has_value: 0 if empty/missing, 1 if has value
        - sort_value: The actual value to sort by (or None if empty)
        - element_id: ElementId as tiebreaker
        """
        try:
            # Try to get parameter
            param = element.LookupParameter(sort_property)
            if not param:
                # Parameter not found - empty value, comes first
                return (0, None, element.Id.IntegerValue)
            
            # Check if parameter has value
            if not param.HasValue:
                # No value - empty value, comes first
                return (0, None, element.Id.IntegerValue)
            
            # Get value based on storage type
            storage_type = param.StorageType
            if storage_type == StorageType.String:
                value = param.AsString()
                if not value or not value.strip():
                    # Empty string - empty value, comes first
                    return (0, None, element.Id.IntegerValue)
                # Has value - return (1, value, element_id) for sorting
                return (1, value, element.Id.IntegerValue)
            elif storage_type == StorageType.Integer:
                value = param.AsInteger()
                # Has value - return (1, value, element_id) for sorting
                return (1, value, element.Id.IntegerValue)
            elif storage_type == StorageType.Double:
                value = param.AsDouble()
                # Has value - return (1, value, element_id) for sorting
                return (1, value, element.Id.IntegerValue)
            elif storage_type == StorageType.ElementId:
                elem_id = param.AsElementId()
                if elem_id and elem_id != ElementId.InvalidElementId:
                    # Has value - return (1, elem_id, element_id) for sorting
                    return (1, elem_id.IntegerValue, element.Id.IntegerValue)
                else:
                    # Invalid ElementId - empty value, comes first
                    return (0, None, element.Id.IntegerValue)
            else:
                # Unknown storage type - empty value, comes first
                return (0, None, element.Id.IntegerValue)
        except Exception as e:
            logger.debug("Error getting sort value for property '{}': {}".format(sort_property, str(e)))
            # Error occurred - empty value, comes first
            return (0, None, element.Id.IntegerValue)
    
    # Sort elements using the sort value function
    # Elements with empty values (has_value=0) come first, then sorted by value
    # CRITICAL: descending only affects sort_value, not has_value (empty values always first)
    try:
        # Separate empty and non-empty elements
        empty_elements = []
        non_empty_elements = []
        for element in source_elements:
            has_value, sort_value, element_id = get_sort_value(element)
            if has_value == 0:
                empty_elements.append(element)
            else:
                non_empty_elements.append(element)
        
        # Sort empty elements by element_id (deterministic)
        empty_elements.sort(key=lambda el: el.Id.IntegerValue)
        
        # Sort non-empty elements by sort_value
        # For descending, we need to handle different types differently
        if descending:
            # For descending, sort ascending first, then reverse
            # This ensures proper descending order for all types (strings, numbers, etc.)
            def get_non_empty_sort_key(element):
                """Get sort key for non-empty elements (ascending)."""
                has_value, sort_value, element_id = get_sort_value(element)
                return (sort_value, element_id)
            
            non_empty_elements.sort(key=get_non_empty_sort_key)
            non_empty_elements.reverse()
        else:
            # For ascending, normal sort
            def get_non_empty_sort_key(element):
                """Get sort key for non-empty elements (ascending)."""
                has_value, sort_value, element_id = get_sort_value(element)
                return (sort_value, element_id)
            
            non_empty_elements.sort(key=get_non_empty_sort_key)
        
        # Combine: empty elements first, then non-empty elements
        result = empty_elements + non_empty_elements
        
        return result
    except Exception as e:
        logger.warning("Error sorting source elements by property '{}': {}".format(sort_property, str(e)))
        # Fallback to ElementId sorting
        return sorted(source_elements, key=lambda el: el.Id.IntegerValue, reverse=descending)

def is_3dzone_family(element):
    """Check if an element is a 3DZone family (Generic Model with family name containing "3DZone").
    
    Args:
        element: Element to check
        
    Returns:
        bool: True if element is a 3DZone family, False otherwise
    """
    try:
        # Check if element is a Generic Model
        if not element.Category:
            return False
        
        generic_model_cat = Category.GetCategory(element.Document, BuiltInCategory.OST_GenericModel)
        if not generic_model_cat or element.Category.Id != generic_model_cat.Id:
            return False
        
        # Check if it's a FamilyInstance with Symbol
        if hasattr(element, "Symbol"):
            symbol = element.Symbol
            if symbol and hasattr(symbol, "FamilyName"):
                family_name = symbol.FamilyName
                if family_name and "3DZone" in family_name:
                    return True
        
        return False
    except Exception:
        return False

def are_target_parameters_empty(element, target_param_names):
    """Check if all target parameters are empty (no value or empty string).
    
    Args:
        element: Element to check
        target_param_names: List of target parameter names
        
    Returns:
        bool: True if all target parameters are empty, False otherwise
    """
    if not target_param_names:
        return True  # No parameters to check, consider as empty
    
    for param_name in target_param_names:
        try:
            # Check instance parameter first
            param = element.LookupParameter(param_name)
            if param:
                if param.HasValue:
                    # Check if value is non-empty
                    if param.StorageType == StorageType.String:
                        value = param.AsString()
                        if value and value.strip():  # Non-empty string
                            return False
                    elif param.StorageType == StorageType.Integer:
                        if param.AsInteger() != 0:  # Non-zero integer
                            return False
                    elif param.StorageType == StorageType.Double:
                        if abs(param.AsDouble()) > 1e-9:  # Non-zero double (with tolerance)
                            return False
                    elif param.StorageType == StorageType.ElementId:
                        eid = param.AsElementId()
                        if eid and eid != element.Id.InvalidElementId:  # Valid element ID
                            return False
                    else:
                        # For other types, if HasValue is True, consider it non-empty
                        return False
            
            # Check type parameter if instance parameter doesn't exist or is empty
            if hasattr(element, "GetTypeId"):
                type_id = element.GetTypeId()
                if type_id and type_id != element.Id.InvalidElementId:
                    element_type = element.Document.GetElement(type_id)
                    if element_type:
                        type_param = element_type.LookupParameter(param_name)
                        if type_param and type_param.HasValue:
                            # Check if value is non-empty
                            if type_param.StorageType == StorageType.String:
                                value = type_param.AsString()
                                if value and value.strip():  # Non-empty string
                                    return False
                            elif type_param.StorageType == StorageType.Integer:
                                if type_param.AsInteger() != 0:  # Non-zero integer
                                    return False
                            elif type_param.StorageType == StorageType.Double:
                                if abs(type_param.AsDouble()) > 1e-9:  # Non-zero double (with tolerance)
                                    return False
                            elif type_param.StorageType == StorageType.ElementId:
                                eid = type_param.AsElementId()
                                if eid and eid != element.Id.InvalidElementId:  # Valid element ID
                                    return False
                            else:
                                # For other types, if HasValue is True, consider it non-empty
                                return False
        except:
            # If we can't check a parameter, assume it might have a value (be safe)
            continue
    
    # All parameters are empty
    return True

# Cache for element types to avoid repeated GetElement() calls
_element_type_cache = {}

def _get_element_type_cached(doc, type_id):
    """Get element type with caching to avoid repeated GetElement() calls.
    
    Args:
        doc: Revit document
        type_id: ElementId of the type
        
    Returns:
        ElementType or None
    """
    if not type_id or type_id.IntegerValue < 0:
        return None
    
    cache_key = type_id.IntegerValue
    if cache_key not in _element_type_cache:
        _element_type_cache[cache_key] = doc.GetElement(type_id)
    return _element_type_cache[cache_key]

def copy_parameters(source_element, target_element, source_param_names, target_param_names, debug_log_count=None, cache_values=False):
    """Copy multiple parameters from source to target element.
    
    Optimized with element type caching to reduce repeated GetElement() calls.
    
    Args:
        source_element: Source element
        target_element: Target element
        source_param_names: List of source parameter names
        target_param_names: List of target parameter names (must match source count)
        debug_log_count: Optional count for limiting debug logs (for first N attempts)
        cache_values: If True, return written parameter values in result dict
        
    Returns:
        dict: {"copied": count, "already_correct": count, "written_values": {param_name: value}}
    """

    if len(source_param_names) != len(target_param_names):
        logger.warning("Parameter name lists don't match in length")
        return {"copied": 0, "already_correct": 0}
    
    copied_count = 0
    already_correct_count = 0
    written_values = {}  # Cache of written parameter values: {param_name: value}
    
    # Cache element types to avoid repeated lookups
    source_type = None
    target_type = None
    
    # Pre-fetch element types if needed
    source_needs_type = False
    target_needs_type = False
    
    for source_name, target_name in zip(source_param_names, target_param_names):
        # Check if we need type parameters
        source_param = source_element.LookupParameter(source_name)
        if not source_param:
            source_needs_type = True
        
        target_param = target_element.LookupParameter(target_name)
        if not target_param:
            target_needs_type = True
    
    # Fetch types once if needed
    if source_needs_type and hasattr(source_element, "GetTypeId"):
        type_id = source_element.GetTypeId()
        if type_id and type_id != source_element.Id.InvalidElementId:
            source_type = _get_element_type_cached(source_element.Document, type_id)
    
    if target_needs_type and hasattr(target_element, "GetTypeId"):
        type_id = target_element.GetTypeId()
        if type_id and type_id != target_element.Id.InvalidElementId:
            target_type = _get_element_type_cached(target_element.Document, type_id)

    # Now copy parameters using cached types
    loop_iteration_count = 0
    for source_name, target_name in zip(source_param_names, target_param_names):
        loop_iteration_count += 1

        try:
            # Get source parameter
            source_param = source_element.LookupParameter(source_name)
            if not source_param and source_type:
                source_param = source_type.LookupParameter(source_name)

            if not source_param:

                continue
            
            if not source_param.HasValue:

                continue
            
            # Get target parameter
            target_param = target_element.LookupParameter(target_name)
            if not target_param and target_type:
                target_param = target_type.LookupParameter(target_name)

            if not target_param:

                continue
            
            if target_param.IsReadOnly:

                continue
            
            # Copy the value
            copy_result = copy_parameter_value(source_param, target_param)

            if copy_result:
                copied_count += 1

            elif not copy_result and source_param and target_param and source_param.HasValue and target_param.HasValue:
                # Check if values already match (already correct, not a failure)
                try:
                    if source_param.StorageType == target_param.StorageType:
                        if source_param.StorageType == StorageType.String:
                            if source_param.AsString() == target_param.AsString():
                                already_correct_count += 1
                        elif source_param.StorageType == StorageType.Integer:
                            if source_param.AsInteger() == target_param.AsInteger():
                                already_correct_count += 1
                        elif source_param.StorageType == StorageType.Double:
                            if abs(source_param.AsDouble() - target_param.AsDouble()) < 1e-9:
                                already_correct_count += 1
                except: pass
        
        except Exception as e:

            continue

    result = {"copied": copied_count, "already_correct": already_correct_count}
    if cache_values:
        result["written_values"] = written_values
    return result

def get_source_document(doc, zone_config):
    """Get the source document (main doc or linked doc) based on configuration.
    
    Args:
        doc: Main Revit document
        zone_config: Configuration dictionary
        
    Returns:
        tuple: (source_doc, link_instance_or_none) - The document to search for source elements
    """
    use_linked = zone_config.get("use_linked_document", False)
    if not use_linked:
        return (doc, None)
    
    linked_doc_name = zone_config.get("linked_document_name")
    if not linked_doc_name:
        logger.warning("use_linked_document is True but linked_document_name is not set, using main document")
        return (doc, None)
    
    try:
        # Find the linked document by name
        link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
        
        for link in link_instances:
            if link.Name == linked_doc_name:
                try:
                    link_doc = link.GetLinkDocument()
                    if link_doc:
                        logger.debug("Using linked document '{}' for source elements".format(linked_doc_name))
                        return (link_doc, link)
                    else:
                        logger.warning("Linked document '{}' found but GetLinkDocument() returned None".format(linked_doc_name))
                except Exception as e:
                    logger.warning("Error accessing linked document '{}': {}".format(linked_doc_name, str(e)))
        
        logger.warning("Linked document '{}' not found, using main document".format(linked_doc_name))
        return (doc, None)
    except Exception as e:
        logger.error("Error finding linked document '{}': {}".format(linked_doc_name, str(e)))
        return (doc, None)

def write_parameters_to_elements(doc, zone_config, progress_bar=None, view_id=None, cache_dict=None):
    """Write parameters to elements based on a zone configuration.
    
    Args:
        doc: Revit document
        zone_config: Configuration dictionary
        progress_bar: Optional progress bar for tracking progress
        view_id: Optional ElementId of view to filter elements by visibility
        cache_dict: Optional dict to cache parameter values: {element_id: {param_name: value}}
        
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
        # Special marker for 3D Zone filter
        THREE_D_ZONE_MARKER = "3DZONE_FILTER"
        
        # Convert special marker to BuiltInCategory for strategy detection
        # (3D Zone uses Generic Model strategy)
        categories_for_strategy = []
        for cat in source_categories:
            # Use str() comparison for IronPython 2.7 compatibility
            if cat == THREE_D_ZONE_MARKER or str(cat) == THREE_D_ZONE_MARKER:
                categories_for_strategy.append(BuiltInCategory.OST_GenericModel)
            else:
                categories_for_strategy.append(cat)
        
        strategy = containment.detect_containment_strategy(categories_for_strategy)
        
        if not strategy:
            error_msg = "Could not detect containment strategy for categories: {} (converted: {})".format(source_categories, categories_for_strategy)
            logger.error(error_msg)
            results["errors"].append(error_msg)
            return results
        
        # Get parameter names early (needed for filtering)
        try:
            source_param_names = zone_config.get("source_params", [])
        except Exception as e:
            raise

        if not source_param_names:

            return results
        
        # Get source document (main doc or linked doc)
        source_doc, link_instance = get_source_document(doc, zone_config)
        
        # Note: view_id filtering doesn't work with linked documents, so we skip it when using linked doc
        use_view_filter = view_id is not None and link_instance is None
        
        # Get source elements
        # Collect elements from each category separately and combine (OR logic)
        # Multiple OfCategory() calls create AND logic (elements in ALL categories), which is wrong
        # Special marker for 3D Zone filter
        THREE_D_ZONE_MARKER = "3DZONE_FILTER"
        
        if source_categories:
            source_elements = []
            element_ids = set()  # Track IDs to avoid duplicates
            for category in source_categories:
                # Handle special 3D Zone marker
                if category == THREE_D_ZONE_MARKER:
                    # Filter Generic Models by family name containing "3DZone"
                    # If view_id is provided and not using linked doc, filter by view visibility
                    if use_view_filter:
                        category_elements = FilteredElementCollector(source_doc, view_id)\
                            .WhereElementIsNotElementType()\
                            .OfCategory(BuiltInCategory.OST_GenericModel)\
                            .ToElements()
                    else:
                        category_elements = FilteredElementCollector(source_doc)\
                            .WhereElementIsNotElementType()\
                            .OfCategory(BuiltInCategory.OST_GenericModel)\
                            .ToElements()
                    
                    # Filter by family name
                    filtered_count = 0
                    total_generic_models = len(category_elements)
                    for el in category_elements:
                        try:
                            if hasattr(el, "Symbol"):
                                symbol = el.Symbol
                                if symbol and hasattr(symbol, "FamilyName"):
                                    family_name = symbol.FamilyName
                                    if family_name and "3DZone" in family_name:
                                        if el.Id.IntegerValue not in element_ids:
                                            element_ids.add(el.Id.IntegerValue)
                                            source_elements.append(el)
                                            filtered_count += 1
                        except:
                            continue
                else:
                    # Regular category
                    # If view_id is provided and not using linked doc, filter by view visibility
                    if use_view_filter:
                        category_elements = FilteredElementCollector(source_doc, view_id)\
                            .WhereElementIsNotElementType()\
                            .OfCategory(category)\
                            .ToElements()
                    else:
                        category_elements = FilteredElementCollector(source_doc)\
                            .WhereElementIsNotElementType()\
                            .OfCategory(category)\
                            .ToElements()
                    for el in category_elements:
                        if el.Id.IntegerValue not in element_ids:
                            element_ids.add(el.Id.IntegerValue)
                            source_elements.append(el)
        else:
            # If view_id is provided and not using linked doc, filter by view visibility
            if use_view_filter:
                source_elements = FilteredElementCollector(source_doc, view_id)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
            else:
                source_elements = FilteredElementCollector(source_doc)\
                    .WhereElementIsNotElementType()\
                    .ToElements()

        if use_view_filter:
            logger.debug("[DEBUG] Found {} total source elements (filtered by view visibility)".format(len(source_elements)))
        else:
            logger.debug("[DEBUG] Found {} total source elements (before filtering)".format(len(source_elements)))
        
        if not source_elements:

            logger.warning("[DEBUG] No source elements found at all for categories: {}".format(source_categories))
            return results
        
        # Filter source elements to only include those with source parameters that have values
        # This avoids processing elements that can't provide values anyway
        filtered_source_elements = []
        elements_without_params = []
        
        for source_el in source_elements:
            has_param = has_source_parameter(source_el, source_param_names)
            
            if has_param:
                filtered_source_elements.append(source_el)
            else:
                elements_without_params.append(source_el.Id)

        logger.debug("[DEBUG] Source elements with parameters '{}': {}".format(source_param_names, len(filtered_source_elements)))
        if elements_without_params:
            logger.debug("[DEBUG] Source elements WITHOUT parameters (first 10): {}".format(elements_without_params[:10]))
        
        source_elements = filtered_source_elements
        
        if not source_elements:

            logger.warning("[DEBUG] No source elements found with required parameters: {}".format(source_param_names))
            return results
        
        # Sort source elements by configured sort property (default: ElementId)
        sort_property = zone_config.get("source_sort_property", "ElementId")
        sort_descending = zone_config.get("source_sort_descending", False)
        
        source_elements = sort_source_elements(source_elements, sort_property, descending=sort_descending)
        logger.debug("[DEBUG] Sorted {} source elements by property: {} (descending: {})".format(len(source_elements), sort_property, sort_descending))
        
        # Pre-compute geometries for Mass/Generic Model elements and Areas (batch operation)
        # This improves performance by calculating all geometries in one pass
        # Use source_doc for geometry operations (works with linked document elements)
        if strategy in ["element", "area"]:
            containment.precompute_geometries(source_elements, source_doc)
        
        # Build spatial index for element strategy (fast path - eliminates per-target DB queries)
        element_index = None
        element_index_cell_size = 50.0  # Default cell size in feet
        if strategy == "element":
            index_start_time = time.time()
            element_index = containment.build_source_element_spatial_index(
                source_elements, source_doc, element_index_cell_size, sort_property=sort_property, sort_descending=sort_descending
            )
            index_time = time.time() - index_start_time
            logger.debug("[DEBUG] Built spatial index in {:.2f}s".format(index_time))
        
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
                # If view_id is provided, filter by view visibility
                if view_id:
                    category_elements = FilteredElementCollector(doc, view_id)\
                        .WhereElementIsNotElementType()\
                        .OfCategory(category)\
                        .ToElements()
                else:
                    category_elements = FilteredElementCollector(doc)\
                        .WhereElementIsNotElementType()\
                        .OfCategory(category)\
                        .ToElements()
                for el in category_elements:
                    if el.Id.IntegerValue not in element_ids:
                        element_ids.add(el.Id.IntegerValue)
                        target_elements.append(el)
        else:
            # If view_id is provided, filter by view visibility
            if view_id:
                target_elements = FilteredElementCollector(doc, view_id)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
            else:
                target_elements = FilteredElementCollector(doc)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
        
        # CRITICAL: When using room-based containment strategy, exclude Rooms and 3DZone families from target elements
        # Rooms always contain themselves, which causes incorrect parameter writes
        # 3DZone families should not be considered as target elements in room computations
        # This must be done AFTER collecting elements, regardless of whether categories were specified
        if strategy == "room":
            count_before = len(target_elements)
            filtered_target_elements = []
            rooms_excluded = 0
            zones_excluded = 0
            for el in target_elements:
                if isinstance(el, Room):
                    rooms_excluded += 1
                    continue
                if is_3dzone_family(el):
                    zones_excluded += 1
                    continue
                filtered_target_elements.append(el)
            target_elements = filtered_target_elements
            if rooms_excluded > 0 or zones_excluded > 0:
                logger.debug("[DEBUG] Excluded {} Rooms and {} 3DZone families from target elements (room-based strategy)".format(rooms_excluded, zones_excluded))

        if view_id:
            logger.debug("[DEBUG] Found {} target elements (filtered by view visibility) for categories: {}".format(len(target_elements), target_filter_categories))
        else:
            logger.debug("[DEBUG] Found {} target elements for categories: {}".format(len(target_elements), target_filter_categories))
        
        if not target_elements:

            logger.warning("[DEBUG] No target elements found for categories: {}".format(target_filter_categories))
            return results
        
        # Get target parameter names (source_param_names already retrieved above)
        target_param_names = zone_config.get("target_params", [])
        
        # Get ifc_export_only_empty flag (default False)
        ifc_export_only_empty = zone_config.get("ifc_export_only_empty", False)
        
        if len(source_param_names) != len(target_param_names):
            error_msg = "Source and target parameter lists don't match"
            logger.error(error_msg)
            results["errors"].append(error_msg)
            return results
        
        # Pre-filter target elements to only include those with target parameters
        # This avoids expensive containment checks on elements that can't accept the parameters anyway
        logger.debug("[DEBUG] Pre-filtering target elements for target parameters: {}".format(target_param_names))
        filtered_target_elements = []
        elements_without_target_params = []
        filter_start_time = time.time()
        
        for target_el in target_elements:
            if has_target_parameter(target_el, target_param_names):
                filtered_target_elements.append(target_el)
            else:
                elements_without_target_params.append(target_el.Id.IntegerValue)
        
        filter_time = time.time() - filter_start_time
        logger.debug("[DEBUG] Pre-filtering complete in {:.2f}s: {} elements have target parameters, {} elements filtered out".format(
            filter_time, len(filtered_target_elements), len(elements_without_target_params)))
        if elements_without_target_params:
            logger.debug("[DEBUG] Elements WITHOUT target parameters (first 10): {}".format(elements_without_target_params[:10]))

        target_elements = filtered_target_elements
        
        if not target_elements:

            logger.warning("[DEBUG] No target elements found with required target parameters: {}".format(target_param_names))
            return results
        
        # Process elements (element-driven loop for performance)
        elements_updated = 0
        elements_already_correct = 0
        total_params_copied = 0
        total_params_already_correct = 0
        updated_element_ids = []  # Track IDs of elements that were updated
        
        # Pre-group rooms/spaces/areas by level if needed
        rooms_by_level = None
        spaces_by_level = None
        areas_by_level = None
        
        # Phase-aware room containment: build phase index for room strategy
        ordered_phases = None
        rooms_by_phase_by_level = None
        phase_map = None  # Phase map from RevitLinkType.GetPhaseMap() for linked documents
        
        if strategy == "room":
            # Build ordered phases once (needed for phase-aware containment)
            # CRITICAL: When using linked documents, we need phases from BOTH documents:
            # - Source doc phases: for organizing rooms (rooms are in linked doc)
            # - Main doc phases: for checking target element phases (targets are in main doc)
            # Phase IDs are document-specific, so we can't mix them!
            source_ordered_phases = containment.get_ordered_phases(source_doc)  # For organizing rooms
            main_ordered_phases = containment.get_ordered_phases(doc) if link_instance else source_ordered_phases  # For checking target elements
            
            # Build rooms by phase and level (phase-aware index) - uses source doc phases
            rooms_list = [el for el in source_elements if isinstance(el, Room)]
            rooms_by_phase_by_level = containment.build_rooms_by_phase_and_level(rooms_list)
            
            # Also build legacy rooms_by_level for fallback/compatibility
            rooms_by_level = defaultdict(list)
            for source_el in source_elements:
                if isinstance(source_el, Room) and source_el.LevelId:
                    rooms_by_level[source_el.LevelId].append(source_el)
            
            # Store both phase lists for phase-aware containment
            ordered_phases = source_ordered_phases  # For room organization
            main_doc_ordered_phases = main_ordered_phases  # For target element phase checking
            
            # Get phase map from RevitLinkType (uses user-configured phase mappings)
            # This provides reliable phase mapping that respects Revit UI settings
            phase_map = containment.get_phase_map_for_link(doc, link_instance) if link_instance else None
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
        # Calculate update interval for progress bar (every 5%)
        update_interval = max(1, int(total_elements / 20.0)) if total_elements > 0 else 1
        last_update = 0
        last_log_time = time.time()
        log_interval = 5.0  # Log status every 5 seconds
        
        # Debug counters
        containment_found_count = 0
        containment_not_found_count = 0
        params_copy_failed_count = 0
        elements_skipped_not_writable = 0
        
        # Performance timing
        start_time = time.time()
        containment_time = 0.0
        param_copy_time = 0.0
        editable_check_time = 0.0
        
        # Optimize: Skip editable check for non-workshared projects
        is_workshared = doc.IsWorkshared

        logger.debug("[DEBUG] Starting to process {} target elements using strategy '{}'".format(total_elements, strategy))
        logger.debug("[DEBUG] Workshared: {}, Skipping editable check: {}".format(is_workshared, not is_workshared))
        
        # Log start of processing
        logger.debug("[PROGRESS] Starting to process {} target elements...".format(total_elements))
        
        for idx, target_el in enumerate(target_elements):
            try:
                results["elements_processed"] += 1
                
                # Log first element to confirm loop is running
                if idx == 0:
                    logger.debug("[PROGRESS] Processing first element (ID: {})...".format(target_el.Id.IntegerValue))
                
                # Update progress bar every 5%
                if progress_bar and total_elements > 0:
                    current_progress = idx + 1
                    if current_progress - last_update >= update_interval or current_progress == total_elements:
                        progress_bar.update_progress(current_progress, total_elements)
                        last_update = current_progress
                
                # Log periodic status updates so user can see progress even if progress bar is slow
                current_time = time.time()
                if current_time - last_log_time >= log_interval:
                    elapsed = current_time - start_time
                    rate = float(idx) / elapsed if elapsed > 0 else 0.0
                    remaining = float(total_elements - idx) / rate if rate > 0 else 0.0
                    percent_complete = float(idx) / total_elements * 100.0 if total_elements > 0 else 0.0
                    logger.debug("[PROGRESS] Processed {}/{} elements ({:.1f}%), {:.1f} elements/sec, ~{:.0f}s remaining".format(
                        idx, total_elements, percent_complete, rate, remaining))
                    last_log_time = current_time
                
                # Check if element is writable (not owned by other users)
                # Skip check for non-workshared projects (performance optimization)
                if is_workshared:
                    editable_start = time.time()
                    is_editable, edit_reason = is_element_editable(doc, target_el)
                    editable_check_time += time.time() - editable_start
                    
                    if not is_editable:
                        elements_skipped_not_writable += 1
                        continue
                
                # Check if we should only process elements with empty target parameters
                if ifc_export_only_empty:
                    if not are_target_parameters_empty(target_el, target_param_names):
                        # Element has at least one filled parameter, skip it
                        continue
                
                # Find containing element
                containment_start = time.time()
                
                # Use phase-aware room containment for room strategy
                if strategy == "room" and ordered_phases is not None and rooms_by_phase_by_level is not None:
                    # Pass main doc phases for element phase checking when using linked documents
                    element_phases_for_checking = main_doc_ordered_phases if link_instance else ordered_phases
                    
                    containing_el = containment.get_containing_room_phase_aware(
                        target_el, source_doc, rooms_by_phase_by_level, ordered_phases,
                        element_phases_for_checking, link_instance,
                        host_doc=doc, phase_map=phase_map  # Use Revit's phase map for reliable cross-doc mapping
                    )
                else:
                    # Convert special marker to BuiltInCategory for containment check
                    # (3D Zone uses Generic Model category)
                    categories_for_containment = []
                    for cat in source_categories:
                        if cat == THREE_D_ZONE_MARKER or str(cat) == THREE_D_ZONE_MARKER:
                            categories_for_containment.append(BuiltInCategory.OST_GenericModel)
                        else:
                            categories_for_containment.append(cat)
                    
                    containing_el = containment.get_containing_element_by_strategy(
                        target_el, source_doc, strategy, categories_for_containment,  # Use source_doc, not doc
                        rooms_by_level, spaces_by_level, areas_by_level,
                        element_index, element_index_cell_size,
                        sort_property=sort_property, sort_descending=sort_descending
                    )
                
                # Additional check: if using 3D Zone marker, verify family name matches
                if containing_el and THREE_D_ZONE_MARKER in source_categories:
                    try:
                        if hasattr(containing_el, "Symbol"):
                            symbol = containing_el.Symbol
                            if symbol and hasattr(symbol, "FamilyName"):
                                family_name = symbol.FamilyName
                                if not family_name or "3DZone" not in family_name:
                                    containing_el = None  # Doesn't match filter
                    except Exception as e:
                        containing_el = None  # Error checking, skip this element
                containment_elapsed = time.time() - containment_start
                containment_time += containment_elapsed
                
                if not containing_el:
                    containment_not_found_count += 1

                    continue
                
                containment_found_count += 1

                # Copy parameters (with caching if cache_dict provided)
                param_start = time.time()
                cache_values = cache_dict is not None
                copy_result = copy_parameters(
                    containing_el, target_el,
                    source_param_names, target_param_names,
                    debug_log_count=containment_found_count,
                    cache_values=cache_values
                )
                
                # Cache written values if cache_dict provided
                if cache_dict is not None and isinstance(copy_result, dict):
                    written_vals = copy_result.get("written_values", {})
                    if written_vals:
                        element_id = target_el.Id.IntegerValue
                        if element_id not in cache_dict:
                            cache_dict[element_id] = {}
                        cache_dict[element_id].update(written_vals)
                param_elapsed = time.time() - param_start
                param_copy_time += param_elapsed
                
                # Handle both old int return and new dict return for backward compatibility
                if isinstance(copy_result, dict):
                    params_copied = copy_result.get("copied", 0)
                    params_already_correct = copy_result.get("already_correct", 0)
                else:
                    # Legacy: treat int as copied count
                    params_copied = copy_result if isinstance(copy_result, int) else 0
                    params_already_correct = 0

                if params_copied > 0:
                    elements_updated += 1
                    total_params_copied += params_copied
                    updated_element_ids.append(target_el.Id.IntegerValue)  # Track updated element ID

                elif params_already_correct > 0:
                    # Element already has correct values - count separately
                    elements_already_correct += 1
                    total_params_already_correct += params_already_correct

                else:
                    params_copy_failed_count += 1

            except Exception as e:
                element_id_val = target_el.Id.IntegerValue if hasattr(target_el.Id, 'IntegerValue') else str(target_el.Id)
                error_msg = "Error processing element {}: {}".format(element_id_val, str(e))
                logger.error(error_msg)
                results["errors"].append(error_msg)
                continue
        
        # Calculate total time
        total_time = time.time() - start_time
        
        # Log summary statistics with performance metrics
        logger.debug("[DEBUG] Processing complete:")
        logger.debug("[DEBUG]   Total time: {:.2f}s".format(total_time))
        logger.debug("[DEBUG]   Time per element: {:.4f}s".format(total_time / total_elements if total_elements > 0 else 0))
        logger.debug("[DEBUG]   Containment check time: {:.2f}s ({:.1f}%)".format(
            containment_time, (containment_time / total_time * 100) if total_time > 0 else 0))
        logger.debug("[DEBUG]   Parameter copy time: {:.2f}s ({:.1f}%)".format(
            param_copy_time, (param_copy_time / total_time * 100) if total_time > 0 else 0))
        logger.debug("[DEBUG]   Editable check time: {:.2f}s ({:.1f}%)".format(
            editable_check_time, (editable_check_time / total_time * 100) if total_time > 0 else 0))
        logger.debug("[DEBUG]   Containment found: {} / {} target elements".format(containment_found_count, total_elements))
        logger.debug("[DEBUG]   Containment not found: {} / {} target elements".format(containment_not_found_count, total_elements))
        logger.debug("[DEBUG]   Elements skipped (not writable): {} / {} target elements".format(elements_skipped_not_writable, total_elements))
        logger.debug("[DEBUG]   Elements updated (values changed): {} elements, {} parameters".format(elements_updated, total_params_copied))
        logger.debug("[DEBUG]   Elements already correct (values matched): {} elements, {} parameters".format(elements_already_correct, total_params_already_correct))
        logger.debug("[DEBUG]   Parameters copy failed (containment found but no params copied): {} elements".format(params_copy_failed_count))
        
        # Performance warnings
        # if total_time > 60:  # More than 1 minute
        #     logger.warning("[PERF] Processing took {:.1f}s - consider optimizing containment strategy or reducing target elements".format(total_time))
        # if containment_time / total_time > 0.7 if total_time > 0 else False:
        #     logger.warning("[PERF] Containment checks are taking {:.1f}% of total time - consider using Rooms/Spaces instead of Areas/Elements".format(
        #         containment_time / total_time * 100))
        if editable_check_time / total_time > 0.2 if total_time > 0 else False:
            logger.warning("[PERF] Editable checks are taking {:.1f}% of total time - consider batch checking or skipping for non-workshared projects".format(
                editable_check_time / total_time * 100))

        results["elements_updated"] = elements_updated
        results["elements_already_correct"] = elements_already_correct
        results["parameters_copied"] = total_params_copied
        results["parameters_already_correct"] = total_params_already_correct
        results["updated_element_ids"] = updated_element_ids  # Track IDs for post-write verification

        return results
    
    except Exception as e:
        error_msg = "Error in write_parameters_to_elements: {}".format(str(e))
        logger.error(error_msg)
        results["errors"].append(error_msg)
        return results

def write_cached_parameters(doc, cache_dict, target_param_names, only_empty=False):
    """Write cached parameter values directly to elements (no recalculation).
    
    Args:
        doc: Revit document
        cache_dict: Dict of cached values: {element_id: {param_name: value}}
        target_param_names: List of target parameter names to write
        only_empty: If True, only write to elements where all target parameters are empty
        
    Returns:
        dict: Results dictionary with counts
    """
    from Autodesk.Revit.DB import ElementId, StorageType
    
    results = {
        "elements_updated": 0,
        "elements_already_correct": 0,
        "parameters_copied": 0,
        "parameters_already_correct": 0,
        "updated_element_ids": []
    }
    
    if not cache_dict:
        return results
    
    for element_id, param_values in cache_dict.items():
        try:
            element = doc.GetElement(ElementId(element_id))
            if not element:
                continue
            
            # Check if we should only process elements with empty target parameters
            if only_empty:
                if not are_target_parameters_empty(element, target_param_names):
                    # Element has at least one filled parameter, skip it
                    continue
            
            params_copied = 0
            params_already_correct = 0
            
            for param_name in target_param_names:
                if param_name not in param_values:
                    continue
                
                cached_value = param_values[param_name]
                param = element.LookupParameter(param_name)
                if not param or param.IsReadOnly:
                    continue
                
                # Check if value already matches
                try:
                    if param.StorageType == StorageType.String:
                        if param.HasValue and param.AsString() == cached_value:
                            params_already_correct += 1
                            continue
                        param.Set(str(cached_value))
                    elif param.StorageType == StorageType.Integer:
                        if param.HasValue and param.AsInteger() == cached_value:
                            params_already_correct += 1
                            continue
                        param.Set(int(cached_value))
                    elif param.StorageType == StorageType.Double:
                        if param.HasValue and abs(param.AsDouble() - cached_value) < 1e-9:
                            params_already_correct += 1
                            continue
                        param.Set(float(cached_value))
                    elif param.StorageType == StorageType.ElementId:
                        cached_eid = ElementId(int(cached_value)) if cached_value else ElementId.InvalidElementId
                        if param.HasValue and param.AsElementId() == cached_eid:
                            params_already_correct += 1
                            continue
                        param.Set(cached_eid)
                    else:
                        continue
                    params_copied += 1
                except Exception:
                    continue
            
            if params_copied > 0:
                results["elements_updated"] += 1
                results["parameters_copied"] += params_copied
                results["updated_element_ids"].append(element_id)
            elif params_already_correct > 0:
                results["elements_already_correct"] += 1
                results["parameters_already_correct"] += params_already_correct
                
        except Exception:
            continue
    
    return results

def execute_configuration(doc, zone_config, progress_bar=None, view_id=None, force_transaction=False, use_subtransaction=False, cache_dict=None, skip_cache_clear=False):
    """Execute a single configuration within a transaction.
    
    Args:
        doc: Revit document
        zone_config: Configuration dictionary
        progress_bar: Optional progress bar for tracking progress
        view_id: Optional ElementId of view to filter elements by visibility
        force_transaction: If True, force normal transaction path even if doc is modifiable
        use_subtransaction: If True, use SubTransaction instead of Transaction (for nesting inside parent transactions)
        cache_dict: Optional dict to cache parameter values: {element_id: {param_name: value}}
        skip_cache_clear: If True, skip clearing geometry cache (for batch executions where cache is cleared once at start)
        
    Returns:
        dict: Results dictionary
    """
    config_name = zone_config.get("name", "Unknown")

    # Clear element type cache at start of each configuration
    global _element_type_cache
    _element_type_cache = {}
    
    # Clear geometry cache at start of each configuration to ensure fresh geometry
    # This is critical because 3D zones can change between runs
    # NOTE: When called from execute_all_configurations or batch execution, cache is cleared once at start
    # This per-config clearing is for standalone calls to execute_configuration
    if not skip_cache_clear:
        containment.clear_geometry_cache()
    
    config_start_time = time.time()
    
    try:

        # Determine transaction strategy based on document modifiability
        # If doc.IsModifiable is True, Revit may have an internal transaction open
        # In that case, we can write directly without starting our own transaction
        use_no_transaction_path = False
        modifiability_check_result = None
        try:
            # Try to access IsModifiable property
            if hasattr(doc, "IsModifiable"):
                try:
                    modifiability_check_result = doc.IsModifiable

                    # If document is modifiable, it means we're inside an existing transaction
                    # We can write parameters without starting a new transaction
                    # UNLESS force_transaction is True (e.g., during FileExporting where Revit's transaction may roll back)
                    if modifiability_check_result and not force_transaction:
                        use_no_transaction_path = True

                except Exception as mod_check_error:
                    pass
            else:
                pass
        except Exception as mod_check_outer_error:
            pass
        
        # Check if we're already in a transaction (from execute_all_configurations)
        # If doc.IsModifiable is True, we're in a transaction, so use no-transaction path
        # UNLESS force_transaction is True (which overrides this behavior)
        if use_no_transaction_path or (not force_transaction and hasattr(doc, "IsModifiable") and doc.IsModifiable):
            # No-transaction path: write directly (assuming we're inside Revit's internal transaction)
            try:
                result = write_parameters_to_elements(doc, zone_config, progress_bar, view_id, cache_dict=cache_dict)

            except Exception as write_error:

                raise
        else:
            # Normal transaction path: create and manage our own transaction
            # If use_subtransaction is True, use SubTransaction to nest inside parent transaction
            transaction_type = "SubTransaction" if use_subtransaction else "Transaction"

            # Use Transaction or SubTransaction based on parameter
            if use_subtransaction:
                from Autodesk.Revit.DB import SubTransaction
                transaction = SubTransaction(doc)
            else:
                transaction = Transaction(doc, "3D Zone: {}".format(config_name))
            
            try:
                transaction.Start()

                result = write_parameters_to_elements(doc, zone_config, progress_bar, view_id, cache_dict=cache_dict)

                # Commit transaction explicitly
                transaction.Commit()

            except Exception as tx_error:

                # Only rollback if transaction was started
                try:
                    if transaction.HasStarted() and not transaction.HasEnded():
                        transaction.RollBack()
                except:
                    pass  # Ignore rollback errors if transaction wasn't started
                raise
            finally:
                # Ensure transaction is disposed
                try:
                    if transaction.HasStarted() and not transaction.HasEnded():
                        transaction.RollBack()
                    transaction.Dispose()
                except:
                    pass  # Ignore disposal errors

    except Exception as e:

        # Check if this is a transaction error - if so, don't log generic error here
        # (the caller will show a formatted message)
        error_msg = str(e)
        is_transaction_error = "not permitted" in error_msg.lower() or "read-only" in error_msg.lower() or "cannot start" in error_msg.lower()
        
        if not is_transaction_error:
            logger.error("Error executing configuration '{}': {}".format(config_name, error_msg))
        
        result = {
            "elements_processed": 0,
            "elements_updated": 0,
            "parameters_copied": 0,
            "errors": [error_msg]
        }
    
    config_time = time.time() - config_start_time
    logger.debug("[PERF] Configuration '{}' completed in {:.2f}s".format(config_name, config_time))
    
    return result

def execute_all_configurations(doc):
    """Execute all enabled configurations in order within a single transaction.
    
    Args:
        doc: Revit document
        
    Returns:
        dict: Summary results with per-configuration results
    """
    from pyrevit import forms
    
    configs = config.get_enabled_configs(doc)
    
    if not configs:
        logger.debug("No enabled configurations found")
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
        "total_elements_already_correct": 0,
        "total_parameters_copied": 0,
        "total_parameters_already_correct": 0
    }
    
    # Clear geometry cache once at the start (not per config)
    containment.clear_geometry_cache()
    
    # Start single transaction for all configurations
    transaction = Transaction(doc, "3D Zone: All Configurations")
    
    try:
        transaction.Start()
        
        class BatchProgressAdapter(object):
            """Adapter to map per-config progress (0..N) into one batch progress bar (0..100%)."""
            def __init__(self, progress_bar, config_idx, total_configs, steps_per_config=100):
                self._pb = progress_bar
                self._config_idx = int(config_idx)
                self._total_configs = int(total_configs) if total_configs else 1
                # steps_per_config is kept for tuning, but output is always mapped to 0..100.
                self._steps_per_config = int(steps_per_config) if steps_per_config else 100
            
            def update_progress(self, current, total):
                try:
                    cur = int(current) if current is not None else 0
                    tot = int(total) if total is not None else 0
                except Exception:
                    cur = 0
                    tot = 0
                
                if tot <= 0:
                    local = 0
                else:
                    ratio = float(cur) / float(tot)
                    if ratio < 0.0:
                        ratio = 0.0
                    elif ratio > 1.0:
                        ratio = 1.0
                    local = int(round(ratio * self._steps_per_config))
                    if local < 0:
                        local = 0
                    elif local > self._steps_per_config:
                        local = self._steps_per_config
                
                overall_ratio = (float(self._config_idx) + (float(local) / float(self._steps_per_config))) / float(self._total_configs)
                if overall_ratio < 0.0:
                    overall_ratio = 0.0
                elif overall_ratio > 1.0:
                    overall_ratio = 1.0
                
                global_current = int(round(overall_ratio * 100.0))
                if global_current < 0:
                    global_current = 0
                elif global_current > 100:
                    global_current = 100
                
                self._pb.update_progress(global_current, 100)
            
            def mark_complete(self):
                overall_ratio = float(self._config_idx + 1) / float(self._total_configs)
                if overall_ratio < 0.0:
                    overall_ratio = 0.0
                elif overall_ratio > 1.0:
                    overall_ratio = 1.0
                global_current = int(round(overall_ratio * 100.0))
                if global_current < 0:
                    global_current = 0
                elif global_current > 100:
                    global_current = 100
                self._pb.update_progress(global_current, 100)
        
        # Single progress bar for the entire batch
        batch_title = "3D Zone: Writing ({} configuration(s))".format(len(configs))
        with forms.ProgressBar(title=batch_title) as pb:
            total_configs = len(configs) if configs else 1
            pb.update_progress(0, 100)
            
            for config_idx, zone_config in enumerate(configs):
                config_name = zone_config.get("name", "Unknown")
                config_order = zone_config.get("order", 0)
                
                logger.debug("[BATCH] Running configuration {}/{}: {}".format(
                    config_idx + 1, total_configs, config_name
                ))
                
                adapter = BatchProgressAdapter(pb, config_idx, total_configs, steps_per_config=100)
                
                # Execute configuration with adapted progress reporter
                # Note: view_id=None means process ALL elements (not filtered by view)
                # Use subtransaction=False and force_transaction=False since we're already in a transaction
                # Skip cache clear since we cleared it once at the start
                result = execute_configuration(
                    doc, zone_config, adapter,
                    view_id=None, force_transaction=False, use_subtransaction=False,
                    skip_cache_clear=True
                )
                adapter.mark_complete()
                
                result["config_name"] = config_name
                result["config_order"] = config_order
                
                summary["config_results"].append(result)
                summary["total_elements_updated"] += result.get("elements_updated", 0)
                summary["total_elements_already_correct"] += result.get("elements_already_correct", 0)
                summary["total_parameters_copied"] += result.get("parameters_copied", 0)
                summary["total_parameters_already_correct"] += result.get("parameters_already_correct", 0)
        
        # Commit single transaction for all configurations
        transaction.Commit()
    
    except Exception as e:
        # Rollback on error
        try:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
        except:
            pass
        raise
    finally:
        # Ensure transaction is disposed
        try:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
            transaction.Dispose()
        except:
            pass
    
    # Log summary
    logger.debug("3D Zone Write Complete: {} configs, {} elements updated ({} params), {} elements already correct ({} params)".format(
        summary["total_configs"],
        summary["total_elements_updated"],
        summary["total_parameters_copied"],
        summary["total_elements_already_correct"],
        summary["total_parameters_already_correct"]
    ))
    
    return summary

