# -*- coding: utf-8 -*-
"""Utility functions for MMI parameter operations."""

import re
from Autodesk.Revit.DB import FilteredElementCollector, ParameterElement, BuiltInParameter
from Autodesk.Revit.DB import ElementId, StorageType

# Try to import ParameterType, which might not be available in all Revit API versions
try:
    from Autodesk.Revit.DB import ParameterType
    HAS_PARAMETER_TYPE = True
except ImportError:
    # Fallback for older Revit versions where ParameterType might be defined differently
    HAS_PARAMETER_TYPE = False

from pyrevit import revit, script

# Initialize logger
logger = script.get_logger()

def find_mmi_parameters(doc):
    """Find all parameters named 'MMI' or containing 'MMI' in the document.
    
    Args:
        doc: The active Revit document
        
    Returns:
        list: List of parameter names that might be MMI parameters
    """
    mmi_params = []
    
    # Get all project parameters
    params = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()
    
    for param in params:
        try:
            # Get parameter name
            param_name = param.Name
            definition = param.GetDefinition()
            
            # Check if parameter name contains 'MMI'
            if 'MMI' in param_name:
                # Check storage type (we want string parameters)
                if HAS_PARAMETER_TYPE:
                    # Use ParameterType if available
                    if definition.ParameterType == ParameterType.Text:
                        mmi_params.append(param_name)
                else:
                    # Fallback: Check if it's a string parameter by storage type
                    if definition.StorageType == StorageType.String:
                        mmi_params.append(param_name)
        except Exception as e:
            logger.debug("Error checking parameter: {}".format(str(e)))
    
    return mmi_params

def validate_mmi_value(value_str):
    """Validate and correct MMI values if needed.
    
    Args:
        value_str: String value to validate
        
    Returns:
        tuple: (original, fixed_value) if correction needed, (None, None) if valid
    """
    if not value_str:
        return None, None
        
    # Extract digits from string
    digits = re.findall(r'\d+', value_str)
    if not digits:
        return None, None
        
    # Extract the numeric value
    extracted_value = int(digits[0])
    
    # Get the original raw value for comparison
    original = value_str.strip()
    
    # Apply validation rules
    # Rule: Should be a 3-digit number between 100-999
    if extracted_value < 100:
        # If less than 3 digits, e.g. "40" -> "400"
        fixed_value = str(extracted_value) + "0" * (3 - len(str(extracted_value)))
        return original, fixed_value
    elif extracted_value > 999:
        # If more than 3 digits, e.g. "2500" -> "250", "1000" -> "100"
        fixed_value = str(extracted_value)[:3]
        return original, fixed_value
    else:
        # Value is already valid, no correction needed
        return None, None

def get_elements_by_mmi_value(doc, mmi_value, param_name=None, comparison="equal"):
    """Get elements with a specific MMI value.
    
    Args:
        doc: The active Revit document
        mmi_value: The MMI value to search for
        param_name: Optional parameter name, will use 'MMI' if None
        comparison: Type of comparison ('equal', 'greater', 'less', 'greater_equal', 'less_equal')
        
    Returns:
        list: List of matching elements
    """
    if not param_name:
        from mmi.core import get_mmi_parameter_name
        param_name = get_mmi_parameter_name(doc)
    
    result_elements = []
    
    # Convert input to int for comparison
    try:
        target_value = int(mmi_value)
    except (ValueError, TypeError):
        # If conversion fails, use string comparison
        target_value = str(mmi_value)
    
    # Get all elements in the model
    all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    
    for element in all_elements:
        try:
            param = element.LookupParameter(param_name)
            
            # If parameter not found on element, try element type
            if not param:
                type_id = element.GetTypeId()
                if type_id and type_id != ElementId.InvalidElementId:
                    element_type = doc.GetElement(type_id)
                    if element_type:
                        param = element_type.LookupParameter(param_name)
            
            # Check if parameter exists and has a value
            if param and param.HasValue and param.StorageType == StorageType.String:
                value_str = param.AsString()
                
                # Try to extract numeric value from string
                match_found = False
                
                if isinstance(target_value, int):
                    # Numeric comparison
                    try:
                        numbers = re.findall(r'\d+', value_str)
                        if numbers:
                            param_value = int(numbers[0])
                            
                            # Compare based on specified comparison type
                            if comparison == "equal" and param_value == target_value:
                                match_found = True
                            elif comparison == "greater" and param_value > target_value:
                                match_found = True
                            elif comparison == "less" and param_value < target_value:
                                match_found = True
                            elif comparison == "greater_equal" and param_value >= target_value:
                                match_found = True
                            elif comparison == "less_equal" and param_value <= target_value:
                                match_found = True
                    except (ValueError, TypeError):
                        # Fall back to string comparison if numeric conversion fails
                        if value_str == str(target_value):
                            match_found = True
                else:
                    # String comparison
                    if value_str == target_value:
                        match_found = True
                
                if match_found:
                    result_elements.append(element)
                    
        except Exception as e:
            logger.debug("Error checking element MMI value: {}".format(str(e)))
    
    return result_elements

def get_mmi_statistics(doc, param_name=None):
    """Get statistics on MMI values in the model.
    
    Args:
        doc: The active Revit document
        param_name: Optional parameter name, will use 'MMI' if None
        
    Returns:
        dict: Statistics on MMI values
    """
    if not param_name:
        from mmi.core import get_mmi_parameter_name
        param_name = get_mmi_parameter_name(doc)
    
    # Initialize statistics
    stats = {
        "total_elements": 0,
        "elements_with_mmi": 0,
        "mmi_values": {},
        "invalid_values": [],
    }
    
    # Get all elements in the model
    all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    stats["total_elements"] = len(all_elements)
    
    for element in all_elements:
        try:
            param = element.LookupParameter(param_name)
            
            # If parameter not found on element, try element type
            if not param:
                type_id = element.GetTypeId()
                if type_id and type_id != ElementId.InvalidElementId:
                    element_type = doc.GetElement(type_id)
                    if element_type:
                        param = element_type.LookupParameter(param_name)
            
            # Check if parameter exists and has a value
            if param and param.HasValue and param.StorageType == StorageType.String:
                value_str = param.AsString()
                stats["elements_with_mmi"] += 1
                
                # Try to extract numeric value from string
                try:
                    numbers = re.findall(r'\d+', value_str)
                    if numbers:
                        mmi_value = int(numbers[0])
                        
                        # Add to statistics
                        if mmi_value in stats["mmi_values"]:
                            stats["mmi_values"][mmi_value] += 1
                        else:
                            stats["mmi_values"][mmi_value] = 1
                    else:
                        # No numeric value found
                        stats["invalid_values"].append({
                            "element_id": element.Id.IntegerValue,
                            "value": value_str,
                            "reason": "No numeric value found"
                        })
                except Exception as ex:
                    # Error extracting numeric value
                    stats["invalid_values"].append({
                        "element_id": element.Id.IntegerValue,
                        "value": value_str,
                        "reason": str(ex)
                    })
                    
        except Exception as e:
            logger.debug("Error checking element MMI value: {}".format(str(e)))
    
    return stats


def get_element_location(element):
    """Get the location point of an element.
    
    Args:
        element: The Revit element to get location from
        
    Returns:
        XYZ: The location point or None if not found
    """
    try:
        if hasattr(element, "Location"):
            location = element.Location
            if location:
                if isinstance(location, LocationPoint):
                    return location.Point
                elif isinstance(location, LocationCurve):
                    return location.Curve.Evaluate(0.5, True)  # Get midpoint
        return None
    except Exception as ex:
        logger.debug("Error getting element location: {}".format(ex))
        return None

def get_element_mmi_value(element, mmi_param_name, doc):
    """Get the MMI value from an element or its type.
    
    Args:
        element: The Revit element to check
        mmi_param_name: The name of the MMI parameter
        doc: The active Revit document
        
    Returns:
        tuple: (numeric_value, value_string, parameter) or (None, None, None) if parameter not found
        - If parameter exists but has no/invalid value: returns (None, value_string, param)
        - If parameter exists with valid MMI: returns (int_value, value_string, param)
        - If parameter doesn't exist: returns (None, None, None)
    """
    try:
        # Check if element has the MMI parameter
        param = element.LookupParameter(mmi_param_name)
        
        # If parameter not found on element, try element type
        if not param:
            try:
                type_id = element.GetTypeId()
                if type_id and type_id != ElementId.InvalidElementId:
                    element_type = doc.GetElement(type_id)
                    if element_type:
                        param = element_type.LookupParameter(mmi_param_name)
            except Exception as e:
                logger.debug("Error getting type parameter: {}".format(e))
        
        # If parameter doesn't exist at all, return None for everything
        if not param:
            return None, None, None
        
        # Parameter exists - check if it has a value
        if param.StorageType == StorageType.String:
            # Get the string value (might be empty or whitespace)
            value_str = param.AsString() if param.HasValue else ""
            
            # Try to extract numeric value from string
            if value_str:
                try:
                    # Extract numbers from string (e.g., "MMI-425" -> 425)
                    numbers = re.findall(r'\d+', value_str)
                    if numbers:
                        return int(numbers[0]), value_str, param
                except Exception as ex:
                    logger.warning("Couldn't parse MMI value from '{}': {}".format(
                        value_str, ex))
            
            # Parameter exists but has no valid numeric value
            return None, value_str, param
                
        return None, None, None
    except Exception as ex:
        logger.debug("Error getting MMI value: {}".format(ex))
        return None, None, None 