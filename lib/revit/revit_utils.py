import clr
clr.AddReference('RevitAPI')
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import *
from pyrevit import revit, HOST_APP
from System.Collections.Generic import List
import random
from unicodedata import normalize
from unicodedata import category as unicode_category

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

def get_available_parameters():
    """Get all available parameters in the document."""
    # Get all elements in active view
    collector = FilteredElementCollector(revit.doc, revit.doc.ActiveView.Id)
    elements = collector.WhereElementIsNotElementType().ToElements()
    
    # Collect unique instance parameters
    params = set()
    for element in elements:
        for param in element.Parameters:
            if not param.IsReadOnly and param.StorageType in [StorageType.String, StorageType.Double, StorageType.Integer]:
                params.add(param.Definition.Name)
    
    return sorted(list(params))

def set_parameter_value(element, param_name, value):
    """Set parameter value on an element."""
    param = element.LookupParameter(param_name)
    if not param:
        return False
        
    try:
        if param.StorageType == StorageType.String:
            param.Set(str(value))
        elif param.StorageType == StorageType.Double:
            param.Set(float(value))
        elif param.StorageType == StorageType.Integer:
            param.Set(int(value))
        return True
    except:
        return False

def isolate_elements(elements):
    """Isolate specific elements in the active view."""
    from pyrevit import revit
    
    doc = revit.doc
    current_view = doc.ActiveView
    
    try:
        # Create a list of element IDs
        element_ids = List[ElementId]([element.Id for element in elements])
        
        # Use transaction wrapper
        with revit.Transaction('Isolate Elements'):
            # Use the built-in temporary isolation mode
            current_view.IsolateElementsTemporary(element_ids)
        
        return True
    except Exception as e:
        print("Error isolating elements: {}".format(str(e)))
        return False

# New utility functions for the ColorElements script

def get_valid_view(doc, uidoc):
    """Get active view from document, checking if it's valid for visibility operations."""
    from Autodesk.Revit.UI import TaskDialog
    
    selected_view = doc.ActiveView
    
    # If we're in a browser view, get the first open UIView
    if selected_view.ViewType == ViewType.ProjectBrowser or selected_view.ViewType == ViewType.SystemBrowser:
        selected_view = doc.GetElement(uidoc.GetOpenUIViews()[0].ViewId)
    
    # Check if view supports visibility modes
    if not selected_view.CanUseTemporaryVisibilityModes():
        TaskDialog.Show(
            "Color Elements by Parameter", 
            "Visibility settings cannot be modified in {} views. Please change your current view.".format(selected_view.ViewType)
        )
        return None
    
    return selected_view

def strip_accents(text):
    """Remove accents from text."""
    return ''.join(char for char in normalize('NFKD', text) if unicode_category(char) != 'Mn')

def solid_fill_pattern_id(doc):
    """Get solid fill pattern ID from document."""
    solid_fill_id = None
    fillpatterns = FilteredElementCollector(doc).OfClass(FillPatternElement)
    for pat in fillpatterns:
        if pat.GetFillPattern().IsSolidFill:
            solid_fill_id = pat.Id
            break
    return solid_fill_id

def apply_color_to_elements(doc, view, element_ids, color, override_projection=True, override_surfaces=True):
    """Apply color override to elements in view.
    
    Args:
        doc: Revit document
        view: View to apply override in
        element_ids: List of ElementIds to apply override to
        color: Revit Color object
        override_projection: Whether to override projection/cut lines
        override_surfaces: Whether to override surface patterns
    """
    # Get solid fill pattern if we need it
    solid_fill_id = solid_fill_pattern_id(doc) if override_surfaces else None
    
    # Create override settings
    ogs = OverrideGraphicSettings()
    
    # Set color properties based on settings
    if override_projection:
        ogs.SetProjectionLineColor(color)
        ogs.SetCutLineColor(color)
        ogs.SetProjectionLinePatternId(ElementId(-1))  # Solid line
    
    # Always set surface pattern color
    if override_surfaces and solid_fill_id:
        ogs.SetSurfaceForegroundPatternColor(color)
        ogs.SetCutForegroundPatternColor(color)
        ogs.SetSurfaceForegroundPatternId(solid_fill_id)
        ogs.SetCutForegroundPatternId(solid_fill_id)
    
    # Apply override to each element
    for element_id in element_ids:
        view.SetElementOverrides(element_id, ogs)

def reset_element_overrides(doc, view, element_ids=None):
    """Reset graphic overrides for elements in a view.
    
    Args:
        doc: Revit document
        view: View to reset overrides in
        element_ids: Optional list of specific ElementIds to reset. If None, resets all elements.
    """
    # Create default override settings (resets to document defaults)
    ogs = OverrideGraphicSettings()
    
    if not element_ids:
        # Get all elements in view if no specific IDs provided
        collector = FilteredElementCollector(doc, view.Id) \
                    .WhereElementIsNotElementType() \
                    .WhereElementIsViewIndependent() \
                    .ToElementIds()
        element_ids = collector
    
    # Reset element overrides
    for element_id in element_ids:
        view.SetElementOverrides(element_id, ogs)
        
def get_parameter_value_string(para):
    """Get parameter value as string."""
    doc = revit.doc
    
    if not para.HasValue:
        return "None"
    
    if para.StorageType == StorageType.Double:
        return para.AsValueString()
    elif para.StorageType == StorageType.ElementId:
        id_val = para.AsElementId()
        # Use the ElementId.InvalidElementId comparison for safety
        if id_val != ElementId.InvalidElementId and id_val.IntegerValue >= 0:
            element = doc.GetElement(id_val)
            if element:
                return element.Name
        return "None"
    elif para.StorageType == StorageType.Integer:
        version = int(HOST_APP.version)
        if version > 2021:
            param_type = para.Definition.GetDataType()
            if SpecTypeId.Boolean.YesNo == param_type:
                return "True" if para.AsInteger() == 1 else "False"
            else:
                return para.AsValueString()
        else:
            # For older Revit versions
            try:
                param_type = para.Definition.ParameterType
                if ParameterType.YesNo == param_type:
                    return "True" if para.AsInteger() == 1 else "False"
                else:
                    return para.AsValueString()
            except:
                # If parameter type can't be determined, just return value string
                return para.AsValueString() or str(para.AsInteger())
    elif para.StorageType == StorageType.String:
        return para.AsString() or "None"
    else:
        return "None"

def generate_color_range(count):
    """Generate a range of visually distinct colors based on count.
    
    Args:
        count: Number of colors needed
        
    Returns:
        List of (r, g, b) tuples representing colors
    """
    # Define color palettes based on count ranges
    if count <= 5:
        # For small sets, use distinct colors
        distinct_colors = [
            (255, 0, 0),      # Red
            (0, 200, 0),      # Green
            (0, 0, 255),      # Blue
            (255, 215, 0),    # Gold/Yellow
            (148, 0, 211)     # Purple
        ]
        return distinct_colors[:count]
        
    elif count <= 20:
        # For medium sets, create a gradient between Red -> Green -> Blue
        colors = []
        if count > 0:
            # First half: Red to Green
            half_count = count // 2
            for i in range(half_count):
                factor = float(i) / (half_count - 1) if half_count > 1 else 0
                r = int(255 * (1 - factor))
                g = int(200 * factor)
                b = 0
                colors.append((r, g, b))
            
            # Second half: Green to Blue
            remaining = count - half_count
            for i in range(remaining):
                factor = float(i) / (remaining - 1) if remaining > 1 else 0
                r = 0
                g = int(200 * (1 - factor))
                b = int(255 * factor)
                colors.append((r, g, b))
        return colors
        
    else:
        # For large sets, use HSV color wheel for better distribution
        colors = []
        for i in range(count):
            # Use HSV with full saturation and value, varying hue
            h = float(i) / count
            # Convert HSV to RGB
            h_i = int(h * 6)
            f = h * 6 - h_i
            p = 0
            q = int(255 * (1 - f))
            t = int(255 * f)
            v = 255
            
            if h_i == 0:
                colors.append((v, t, p))
            elif h_i == 1:
                colors.append((q, v, p))
            elif h_i == 2:
                colors.append((p, v, t))
            elif h_i == 3:
                colors.append((p, q, v))
            elif h_i == 4:
                colors.append((t, p, v))
            else:
                colors.append((v, p, q))
        
        return colors

def get_categories_in_view(doc, view, excluded_cats=None):
    """Get used categories and parameters in the current view.
    
    Args:
        doc: Revit document
        view: View to get categories from
        excluded_cats: List of category IDs to exclude
        
    Returns:
        Dictionary with category information and parameters
    """
    
    if excluded_cats is None:
        excluded_cats = []
    
    # Get elements in the view
    collector = FilteredElementCollector(doc, view.Id) \
                .WhereElementIsNotElementType() \
                .WhereElementIsViewIndependent() \
                .ToElements()
    
    categories_dict = {}
    get_elementid_value = get_elementid_value_func()
    
    for element in collector:
        if element.Category is None:
            continue
            
        category = element.Category
        cat_id = get_elementid_value(category.Id)
        
        # Skip excluded categories
        if cat_id in excluded_cats:
            continue
            
        # If category not yet processed, add it
        if cat_id not in categories_dict:
            categories_dict[cat_id] = {
                'category': category,
                'instance_parameters': {},
                'type_parameters': {}
            }
        
        # Get instance parameters if not already collected
        if not categories_dict[cat_id]['instance_parameters']:
            for param in element.Parameters:
                # Skip category parameters
                if param.Definition.BuiltInParameter in [
                    BuiltInParameter.ELEM_CATEGORY_PARAM, 
                    BuiltInParameter.ELEM_CATEGORY_PARAM_MT
                ]:
                    continue
                
                param_name = param.Definition.Name
                categories_dict[cat_id]['instance_parameters'][param_name] = param
        
        # Get type parameters if not already collected
        if not categories_dict[cat_id]['type_parameters']:
            type_element = doc.GetElement(element.GetTypeId())
            if type_element:
                for param in type_element.Parameters:
                    # Skip category parameters
                    if param.Definition.BuiltInParameter in [
                        BuiltInParameter.ELEM_CATEGORY_PARAM, 
                        BuiltInParameter.ELEM_CATEGORY_PARAM_MT
                    ]:
                        continue
                    
                    param_name = param.Definition.Name
                    categories_dict[cat_id]['type_parameters'][param_name] = param
    
    return categories_dict 

def get_elementid_value_func():
    """Returns the ElementId value extraction function based on the Revit version.
    
    Follows API changes in Revit 2024.

    Returns:
        function: A function that returns the value of an ElementId.
    """    
    def get_value_post2024(item):
        return item.Value

    def get_value_pre2024(item):
        return item.IntegerValue

    version = int(HOST_APP.version)
    return get_value_post2024 if version > 2023 else get_value_pre2024 