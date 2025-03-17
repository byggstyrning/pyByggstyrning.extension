import clr
from Autodesk.Revit.DB import *
from pyrevit import revit

def get_visible_elements():
    """Get all visible elements in the current view"""
    doc = revit.doc
    current_view = doc.ActiveView
    
    # Use FilteredElementCollector to get all elements visible in current view
    collector = FilteredElementCollector(doc, current_view.Id)
    elements = collector.WhereElementIsNotElementType().ToElements()
    
    return elements

def get_element_by_ifc_guid(ifc_guid):
    """Get Revit element by IFC GUID"""
    doc = revit.doc
    
    # Create a filtered element collector for the document
    collector = FilteredElementCollector(doc)
    elements = collector.WhereElementIsNotElementType().ToElements()
    
    # Find elements with matching IFC GUID
    for element in elements:
        try:
            # Try to get IFC GUID parameter - can be named different ways
            param = element.LookupParameter("IFCGuid")
            if not param:
                param = element.LookupParameter("IfcGUID")
            if not param:
                param = element.LookupParameter("IFC GUID")
                
            if param and param.AsString() == ifc_guid:
                return element
        except:
            continue
    
    return None

def get_available_parameters(element_types=None):
    """Get all available parameters that could be set on elements"""
    doc = revit.doc
    
    if not element_types:
        # Get a sample element to extract parameters
        sample_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        for element in sample_elements:
            if element.Category and element.Category.HasMaterialQuantities:
                break
    else:
        # Use the provided element types
        sample_elements = []
        for element_type in element_types:
            collector = FilteredElementCollector(doc).OfCategory(element_type)
            sample_elements.extend(collector.WhereElementIsNotElementType().ToElements())
    
    params = set()
    
    # Extract parameter names from sample elements
    for element in sample_elements:
        for param in element.Parameters:
            if param.StorageType in [StorageType.String, StorageType.Integer, StorageType.Double]:
                params.add(param.Definition.Name)
    
    return sorted(list(params))

def set_parameter_value(element, parameter_name, value):
    """Set parameter value on an element"""
    if not element:
        return False
    
    param = element.LookupParameter(parameter_name)
    if not param:
        return False
    
    # Check if parameter is read-only
    if param.IsReadOnly:
        return False
    
    # Set parameter value based on storage type
    try:
        with revit.Transaction("Set Parameter Value"):
            if param.StorageType == StorageType.String:
                param.Set(str(value))
            elif param.StorageType == StorageType.Integer:
                param.Set(int(value))
            elif param.StorageType == StorageType.Double:
                param.Set(float(value))
            else:
                return False
        return True
    except Exception as e:
        print("Error setting parameter: {}".format(str(e)))
        return False 