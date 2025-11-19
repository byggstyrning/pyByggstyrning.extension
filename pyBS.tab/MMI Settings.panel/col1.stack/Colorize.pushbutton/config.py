# -*- coding: utf-8 -*-
"""Shift-click handler for Colorizer.

When shift-clicking, this creates view filters for MMI ranges and applies them to the active view.
This provides a more permanent solution compared to direct color overrides.
"""

__title__ = "Create MMI Filters"
__author__ = "Byggstyrning AB"
__doc__ = "Shift-click: Creates view filters for MMI ranges and applies them to the active view"

# Import standard libraries
import sys
import os
import os.path as op

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Add the extension directory to the path
script_path = __file__
pushbutton_dir = op.dirname(script_path)
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Try direct import from current directory's parent path
sys.path.append(op.dirname(op.dirname(panel_dir)))

# Initialize logger
logger = script.get_logger()

# Import MMI libraries
from mmi.core import get_mmi_parameter_name

# Official MMI Color codes from https://mmi-veilederen.no/?page_id=85
# Each filter matches exact MMI values (e.g., "000", "100", "125", etc.)
MMI_FILTER_RANGES = [
    {"value": "000", "name": "MMI_000_Tidligfase", "display_name": "000 - Tidligfase", 
     "color": Color(215, 50, 150)},
    {"value": "100", "name": "MMI_100_Grunnlagsinformasjon", "display_name": "100 - Grunnlagsinformasjon", 
     "color": Color(190, 40, 35)},
    {"value": "125", "name": "MMI_125_Etablert_konsept", "display_name": "125 - Etablert konsept", 
     "color": Color(210, 75, 70)},
    {"value": "150", "name": "MMI_150_Tverrfaglig_kontrollert_konsept", "display_name": "150 - Tverrfaglig kontrollert konsept", 
     "color": Color(225, 120, 115)},
    {"value": "175", "name": "MMI_175_Valgt_konsept", "display_name": "175 - Valgt konsept", 
     "color": Color(240, 170, 170)},
    {"value": "200", "name": "MMI_200_Ferdig_konsept", "display_name": "200 - Ferdig konsept", 
     "color": Color(230, 150, 55)},
    {"value": "225", "name": "MMI_225_Etablert_prinsipielle", "display_name": "225 - Etablert prinsipielle løsninger", 
     "color": Color(235, 175, 100)},
    {"value": "250", "name": "MMI_250_Tverrfaglig_kontrollert_prinsipielle", "display_name": "250 - Tverrfaglig kontrollert prinsipielle løsninger", 
     "color": Color(240, 200, 140)},
    {"value": "275", "name": "MMI_275_Valgt_prinsipielle", "display_name": "275 - Valgt prinsipielle løsninger", 
     "color": Color(245, 230, 215)},
    {"value": "300", "name": "MMI_300_Underlag_for_detaljering", "display_name": "300 - Underlag for detaljering", 
     "color": Color(250, 240, 80)},
    {"value": "325", "name": "MMI_325_Etablert_detaljerte", "display_name": "325 - Etablert detaljerte løsninger", 
     "color": Color(215, 205, 65)},
    {"value": "350", "name": "MMI_350_Tverrfaglig_kontrollert_detaljerte", "display_name": "350 - Tverrfaglig kontrollert detaljerte løsninger", 
     "color": Color(185, 175, 60)},
    {"value": "375", "name": "MMI_375_Detaljerte_anbud", "display_name": "375 - Detaljerte løsninger (anbud/bestilling)", 
     "color": Color(150, 150, 50)},
    {"value": "400", "name": "MMI_400_Arbeidsgrunnlag", "display_name": "400 - Arbeidsgrunnlag", 
     "color": Color(55, 130, 70)},
    {"value": "425", "name": "MMI_425_Etablert_utfort", "display_name": "425 - Etablert/utført", 
     "color": Color(75, 170, 90)},
    {"value": "450", "name": "MMI_450_Kontrollert_utforelse", "display_name": "450 - Kontrollert utførelse", 
     "color": Color(100, 195, 125)},
    {"value": "475", "name": "MMI_475_Godkjent_utforelse", "display_name": "475 - Godkjent utførelse", 
     "color": Color(155, 215, 165)},
    {"value": "500", "name": "MMI_500_Som_bygget", "display_name": "500 - Som bygget", 
     "color": Color(30, 70, 175)},
    {"value": "600", "name": "MMI_600_I_drift", "display_name": "600 - I drift", 
     "color": Color(175, 50, 205)},
]

def get_all_filterable_categories(doc):
    """Get all categories that can be used in filters."""
    filterable_categories = []
    
    # Get all categories
    categories = doc.Settings.Categories
    
    for category in categories:
        # Only include categories that allow bound parameters and are visible
        if category.AllowsBoundParameters and category.CategoryType == CategoryType.Model:
            try:
                # Check if category can be used in filters
                builtin_cat = category.Id.IntegerValue
                if builtin_cat < 0:  # Built-in categories have negative IDs
                    filterable_categories.append(category.Id)
            except:
                pass
    
    return filterable_categories

def create_mmi_filter(doc, filter_range, mmi_param_id, categories):
    """Create a single MMI filter for an exact MMI value.
    
    Args:
        doc: Revit document
        filter_range: Dictionary with filter definition (contains exact value)
        mmi_param_id: ElementId of the MMI parameter
        categories: List of category IDs to filter
        
    Returns:
        ParameterFilterElement or None if failed
    """
    try:
        filter_name = filter_range["name"]
        mmi_value = filter_range["value"]
        
        # Check if filter already exists
        existing_filters = FilteredElementCollector(doc) \
            .OfClass(ParameterFilterElement) \
            .ToElements()
        
        for existing_filter in existing_filters:
            if existing_filter.Name == filter_name:
                logger.debug("Filter '{}' already exists, will reuse".format(filter_name))
                return existing_filter
        
        # Create filter rule for exact MMI value match
        # MMI values are stored as strings like "000", "100", "125", etc.
        # Use Equals rule for exact matching
        
        rule = ParameterFilterRuleFactory.CreateEqualsRule(
            mmi_param_id,
            mmi_value,
            True  # Case sensitive
        )
        
        element_filter = ElementParameterFilter(rule)
        
        # Create the ParameterFilterElement
        param_filter = ParameterFilterElement.Create(
            doc,
            filter_name,
            categories,
            element_filter
        )
        
        logger.debug("Created filter: {} for MMI value '{}'".format(filter_name, mmi_value))
        return param_filter
        
    except Exception as ex:
        logger.error("Error creating filter '{}': {}".format(filter_range["name"], ex))
        import traceback
        logger.error(traceback.format_exc())
        return None

def apply_filter_to_view(doc, view, param_filter, filter_range):
    """Apply a filter to a view with graphics overrides.
    
    Args:
        doc: Revit document
        view: View to apply filter to
        param_filter: ParameterFilterElement to apply
        filter_range: Dictionary with filter definition (for color)
    """
    try:
        # Check if filter is already applied
        existing_filters = view.GetFilters()
        if param_filter.Id in existing_filters:
            logger.debug("Filter '{}' already applied to view".format(param_filter.Name))
            # Update the override anyway
            pass
        else:
            # Add filter to view
            view.AddFilter(param_filter.Id)
            logger.debug("Added filter '{}' to view".format(param_filter.Name))
        
        # Get solid fill pattern
        solid_fill_id = None
        patterns = FilteredElementCollector(doc).OfClass(FillPatternElement)
        for pat in patterns:
            fill_pattern = pat.GetFillPattern()
            if fill_pattern.IsSolidFill:
                solid_fill_id = pat.Id
                break
        
        # Create graphics override
        color = filter_range["color"]
        ogs = OverrideGraphicSettings()
        ogs.SetProjectionLineColor(color)
        ogs.SetCutLineColor(color)
        ogs.SetSurfaceForegroundPatternColor(color)
        ogs.SetCutForegroundPatternColor(color)
        
        if solid_fill_id:
            ogs.SetSurfaceForegroundPatternId(solid_fill_id)
            ogs.SetCutForegroundPatternId(solid_fill_id)
        
        # Set the filter override
        view.SetFilterOverrides(param_filter.Id, ogs)
        logger.debug("Applied graphics override to filter '{}'".format(param_filter.Name))
        
    except Exception as ex:
        logger.error("Error applying filter '{}' to view: {}".format(param_filter.Name, ex))
        import traceback
        logger.error(traceback.format_exc())

def create_and_apply_mmi_filters():
    """Main function to create MMI filters and apply them to the active view."""
    try:
        doc = revit.doc
        active_view = doc.ActiveView
        
        # Check if we have a valid view
        if not active_view:
            forms.alert("No active view found.", title="Error")
            return
        
        if active_view.ViewType == ViewType.DrawingSheet:
            forms.alert("View filters cannot be applied to sheets. Please open a model view.", 
                       title="Invalid View Type")
            return
        
        if active_view.IsTemplate:
            forms.alert("Cannot apply filters to view templates.", 
                       title="Invalid View Type")
            return
        
        # Get MMI parameter name
        mmi_param_name = get_mmi_parameter_name(doc)
        if not mmi_param_name:
            forms.alert("No MMI parameter configured. Please configure MMI settings first.", 
                       title="MMI Parameter Not Configured")
            return
        
        logger.debug("Using MMI parameter: {}".format(mmi_param_name))
        
        # Get the MMI parameter ID from a shared parameter or project parameter
        mmi_param_id = None
        param_definition = None
        
        # Try to find the parameter in project parameters first
        param_bindings = doc.ParameterBindings
        iterator = param_bindings.ForwardIterator()
        
        while iterator.MoveNext():
            definition = iterator.Key
            if definition.Name == mmi_param_name:
                param_definition = definition
                logger.debug("Found parameter definition in bindings: {}".format(mmi_param_name))
                break
        
        # If found in bindings, try to get parameter ID from an element
        if param_definition:
            collector = FilteredElementCollector(doc) \
                .WhereElementIsNotElementType() \
                .ToElements()
            
            # Try up to 100 elements to find the parameter
            checked_count = 0
            max_checks = 100
            
            for elem in collector:
                if checked_count >= max_checks:
                    break
                    
                try:
                    param = elem.LookupParameter(mmi_param_name)
                    if param and param.Id:
                        mmi_param_id = param.Id
                        logger.debug("Found parameter ID from element {}: {}".format(elem.Id, mmi_param_id))
                        break
                except:
                    pass
                
                checked_count += 1
        
        # If still not found, try checking element types as well
        if not mmi_param_id:
            logger.debug("Parameter ID not found on elements, checking element types...")
            type_collector = FilteredElementCollector(doc) \
                .WhereElementIsElementType() \
                .ToElements()
            
            checked_count = 0
            max_checks = 50
            
            for elem_type in type_collector:
                if checked_count >= max_checks:
                    break
                    
                try:
                    param = elem_type.LookupParameter(mmi_param_name)
                    if param and param.Id:
                        mmi_param_id = param.Id
                        logger.debug("Found parameter ID from element type {}: {}".format(elem_type.Id, mmi_param_id))
                        break
                except:
                    pass
                
                checked_count += 1
        
        if not mmi_param_id:
            error_msg = (
                "Could not find the MMI parameter '{}' in the model.\n\n"
                "The parameter may not be assigned to any elements, or it may be a built-in parameter "
                "that cannot be used for view filters.\n\n"
                "Please check:\n"
                "• Ensure elements have the '{}' parameter assigned\n"
                "• Verify the parameter is a project parameter or shared parameter\n"
                "• Consider reconfiguring the MMI parameter in Settings"
            ).format(mmi_param_name, mmi_param_name)
            
            forms.alert(error_msg, title="Parameter Not Found")
            logger.error("Failed to find MMI parameter '{}' for filter creation".format(mmi_param_name))
            return
        
        logger.debug("Found MMI parameter ID: {}".format(mmi_param_id))
        
        # Get all filterable categories
        categories = get_all_filterable_categories(doc)
        if not categories:
            forms.alert("No filterable categories found in the model.", 
                       title="No Categories")
            return
        
        logger.debug("Found {} filterable categories".format(len(categories)))
        
        # Create filters in a transaction
        created_filters = []
        
        with Transaction(doc, "Create MMI View Filters") as t:
            t.Start()
            
            for filter_range in MMI_FILTER_RANGES:
                param_filter = create_mmi_filter(doc, filter_range, mmi_param_id, categories)
                if param_filter:
                    created_filters.append((param_filter, filter_range))
            
            t.Commit()
        
        if not created_filters:
            forms.alert("Failed to create any view filters.", 
                       title="Error")
            return
        
        # Apply filters to the active view
        with Transaction(doc, "Apply MMI Filters to View") as t:
            t.Start()
            
            for param_filter, filter_range in created_filters:
                apply_filter_to_view(doc, active_view, param_filter, filter_range)
            
            t.Commit()
        
        # Show success message
        forms.show_balloon(
            header="MMI View Filters Created",
            text="Created {} view filters and applied them to '{}'.\n\nFilters:\n• {}".format(
                len(created_filters),
                active_view.Name,
                "\n• ".join([f["display_name"] for p, f in created_filters])
            ),
            is_new=True
        )
        
        logger.debug("Successfully created and applied {} MMI filters".format(len(created_filters)))
        
    except Exception as ex:
        logger.error("Error creating MMI filters: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
        forms.alert("Error creating MMI filters:\n\n{}".format(ex), 
                   title="Error")

# --- Main Execution --- 

if __name__ == '__main__':
    logger.debug("Shift-click: Creating MMI view filters...")
    create_and_apply_mmi_filters()

