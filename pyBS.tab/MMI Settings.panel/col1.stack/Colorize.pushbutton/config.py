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
from pyrevit import HOST_APP
from pyrevit.framework import List

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

# Import styles for theme support
from styles import ensure_styles_loaded

# Try direct import from current directory's parent path
sys.path.append(op.dirname(op.dirname(panel_dir)))

# Initialize logger
logger = script.get_logger()

# Debug logging setup
import json
import datetime
import codecs
DEBUG_LOG_PATH = op.join(extension_dir, '.cursor', 'debug.log')

def safe_str(obj):
    """Safely convert object to string, handling Unicode."""
    if obj is None:
        return None
    try:
        # Check for unicode type (IronPython 2.7)
        if isinstance(obj, unicode):
            return obj
    except NameError:
        # unicode doesn't exist (Python 3+)
        pass
    if isinstance(obj, str):
        try:
            return obj.decode('utf-8')
        except:
            return obj
    return str(obj)

def safe_json_encode(obj):
    """Safely encode object to JSON, handling Unicode issues."""
    if isinstance(obj, dict):
        return {safe_str(k): safe_json_encode(v) for k, v in obj.items()}
    elif isinstance(obj, (list, tuple)):
        return [safe_json_encode(item) for item in obj]
    else:
        return safe_str(obj)

def debug_log(location, message, data=None, hypothesis_id=None):
    """Write debug log entry."""
    try:
        # Clean all data to avoid Unicode issues
        safe_data = safe_json_encode(data) if data else {}
        
        log_entry = {
            "timestamp": int((datetime.datetime.now() - datetime.datetime(1970, 1, 1)).total_seconds() * 1000),
            "location": safe_str(location),
            "message": safe_str(message),
            "data": safe_data,
            "sessionId": "debug-session",
            "runId": "run1",
            "hypothesisId": safe_str(hypothesis_id) if hypothesis_id else None
        }
        # Ensure directory exists
        log_dir = op.dirname(DEBUG_LOG_PATH)
        if not op.exists(log_dir):
            try:
                os.makedirs(log_dir)
            except:
                pass
        # Write to file with UTF-8 encoding
        json_str = json.dumps(log_entry, ensure_ascii=False)
        with codecs.open(DEBUG_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(json_str + '\n')
        # Also log to pyRevit output for visibility (simplified)
        logger.debug("[DEBUG] {}: {}".format(safe_str(location), safe_str(message)))
    except Exception as e:
        # Log to pyRevit logger as fallback
        logger.error("Debug log failed: {} - Path: {}".format(str(e), DEBUG_LOG_PATH))
        import traceback
        logger.error(traceback.format_exc())

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

def get_categories_with_parameter(doc, mmi_param_id, mmi_param_name):
    """Get categories where the MMI parameter actually exists.
    
    This checks ParameterBindings first, then verifies by checking elements.
    Only returns categories where the parameter can be used in filters.
    """
    # Use .NET List[ElementId] like ColorSplasher does (line 489)
    valid_categories = List[ElementId]()
    
    
    # Method 1: Check ParameterBindings (works for project/shared parameters)
    try:
        param_bindings = doc.ParameterBindings
        iterator = param_bindings.ForwardIterator()
        iterator.Reset()
        
        while iterator.MoveNext():
            binding = iterator.Current
            param_def = binding.Key
            
            # Check by ID first
            if param_def.Id == mmi_param_id:
                # Found the parameter binding - get its categories
# Exclude certain categories that shouldn't be in filters
                excluded_category_ids = {
                    ElementId(-2000700),  # Materials
                    ElementId(-2003200),  # Areas
                    ElementId(-2000160),  # Rooms
                    ElementId(-2008107),  # HVAC Zones
                }
                
                categories = binding.Categories
                for cat in categories:
                    if cat.CategoryType == CategoryType.Model and cat.AllowsBoundParameters:
                        # Skip excluded categories
                        if cat.Id in excluded_category_ids:
continue
                        
                        valid_categories.Add(cat.Id)
if valid_categories.Count > 0:
                    return valid_categories
            
            # Also check by name (for built-in parameters that might not match by ID)
            if param_def.Name == mmi_param_name:
# Exclude certain categories that shouldn't be in filters
                excluded_category_ids = {
                    ElementId(-2000700),  # Materials
                    ElementId(-2003200),  # Areas
                    ElementId(-2000160),  # Rooms
                    ElementId(-2008107),  # HVAC Zones
                }
                
                categories = binding.Categories
                for cat in categories:
                    if cat.CategoryType == CategoryType.Model and cat.AllowsBoundParameters:
                        # Skip excluded categories
                        if cat.Id in excluded_category_ids:
                            continue
                        
                        # Avoid duplicates
                        if cat.Id not in [c.IntegerValue for c in valid_categories]:
                            valid_categories.Add(cat.Id)
if valid_categories.Count > 0:
                    return valid_categories
    except Exception as ex:
        pass
    
    # Method 2: For built-in parameters or if ParameterBindings didn't work,
    # find categories by checking elements that have the parameter
# Get unique categories from elements that have this parameter
    category_ids_found = set()
    
    # Check ALL elements, not just view-independent ones, to find all categories
    collector = FilteredElementCollector(doc) \
        .WhereElementIsNotElementType() \
        .ToElements()
    
    checked_count = 0
    max_checks = 1000  # Check more elements to find all categories
    
    for elem in collector:
        if checked_count >= max_checks:
            break
        
        try:
            # Skip elements without a valid category
            if not elem.Category or not elem.Category.Id:
                continue
            
            # Skip non-model categories (like Materials, which is -2000700)
            if elem.Category.CategoryType != CategoryType.Model:
                continue
            
            # Check if element has the parameter
            param = None
            try:
                param = elem.LookupParameter(mmi_param_name)
            except:
                pass
            
            param_found = False
            if param and param.Id == mmi_param_id:
                param_found = True
            
            # Also check by iterating Parameters (for built-in parameters)
            if not param_found:
                for pr in elem.Parameters:
                    if pr.Id == mmi_param_id:
                        param_found = True
                        break
            
            if param_found:
                # Element has the parameter - add its category
                cat_id = elem.Category.Id.IntegerValue
                category_ids_found.add(cat_id)
        except Exception as ex:
pass
        
        checked_count += 1
    
    # Convert found category IDs to ElementId list
    # Filter out invalid categories and categories that shouldn't be in filters
    excluded_category_ids = {
        -2000700,  # Materials (not an element category)
        -2003200,  # Areas (special category, not typically filtered)
        -2000160,  # Rooms (special category, not typically filtered)
        -2008107,  # HVAC Zones (special category)
    }
    
    for cat_id_int in category_ids_found:
        try:
            cat_id = ElementId(cat_id_int)
            category = Category.GetCategory(doc, cat_id)
            
            # Skip excluded categories
            if cat_id_int in excluded_category_ids:
continue
            
            # Only add model categories that allow bound parameters
            if category and category.CategoryType == CategoryType.Model and category.AllowsBoundParameters:
                # Additional validation: ensure category can be used in filters
                # Some categories like Areas, Rooms, HVAC Zones shouldn't be in element filters
                try:
                    # Test if category can be used in a filter by checking if it's filterable
                    # This is a heuristic - we'll include it if it passes the above checks
                    valid_categories.Add(cat_id)
except Exception as ex:
return valid_categories

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
        # Filter name should be just "MMI_xxx" (e.g., "MMI_000"), not "MMI_000_Tidligfase"
        filter_name = "MMI_{}".format(filter_range["value"])
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
        # Use CreateEqualsRule (same as ColorSplasher does for String parameters)
        # ColorSplasher uses CreateEqualsRule for all parameter types, including built-in parameters
        # The key is getting the parameter ID from an element's Parameter object, not from bindings
        
        # Get Revit version
        version = int(HOST_APP.version)
        if version > 2023:
            # Revit 2024+ uses different signature (no case_sensitive parameter)
            rule = ParameterFilterRuleFactory.CreateEqualsRule(
                mmi_param_id,
                mmi_value
            )
        else:
            # Revit 2023 and earlier requires case_sensitive parameter
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
        
        # Get parameter ID from an element (ColorSplasher method)
        # ColorSplasher gets parameter ID directly from element.Parameters, not from ParameterBindings
        # This works for BOTH project/shared parameters AND built-in parameters
        # See ColorSplasher line 1413-1414: iterates through elem.Parameters to find matching name
        mmi_param_id = None
        
        collector = FilteredElementCollector(doc) \
            .WhereElementIsNotElementType() \
            .WhereElementIsViewIndependent() \
            .ToElements()
        
        # Try up to 200 elements to find the parameter
        checked_count = 0
        max_checks = 200
        
        for elem in collector:
            if checked_count >= max_checks:
                break
            
            try:
                # Method 1: Try LookupParameter first (works for project/shared parameters)
                param = elem.LookupParameter(mmi_param_name)
                if param and param.Id:
                    mmi_param_id = param.Id
                    break
                
                # Method 2: Iterate through Parameters (works for built-in parameters too)
                # This is how ColorSplasher does it (see line 1413-1414)
                if not mmi_param_id:
                    for pr in elem.Parameters:
                        if pr.Definition.Name == mmi_param_name:
                            if pr.Id:
                                mmi_param_id = pr.Id
                                break
                    if mmi_param_id:
                        break
            except Exception as ex:
                pass
            
            checked_count += 1
        
        # Note: MMI is always an instance parameter, so we don't need to check element types
        
        if not mmi_param_id:
            # Build error message
            error_msg = (
                "Could not find the MMI parameter '{}' on any elements in the model.\n\n"
                "Please check:\n"
                "• Verify the parameter name matches exactly (case-sensitive)\n"
                "• Ensure there are elements in the view with this parameter\n"
                "• Try setting the parameter value on at least one element first\n"
                "• Verify the parameter is accessible on elements\n"
                "• Configure the MMI parameter in Settings if needed"
            ).format(mmi_param_name)
            
            forms.alert(error_msg, title="Parameter Not Found")
            logger.error("Failed to find MMI parameter '{}' for filter creation".format(mmi_param_name))
            return
        
        logger.debug("Found MMI parameter ID: {}".format(mmi_param_id))
        
        # Get categories where the parameter actually exists
        categories = get_categories_with_parameter(doc, mmi_param_id, mmi_param_name)
        if categories.Count == 0:
            error_msg = (
                "The MMI parameter '{}' was not found on any elements or categories in the model.\n\n"
                "Please ensure:\n"
                "• The parameter exists and is bound to at least one category\n"
                "• There are elements in the model with this parameter\n"
                "• The parameter is accessible on model elements"
            ).format(mmi_param_name)
            forms.alert(error_msg, title="No Categories Found")
            logger.error("No categories found with parameter '{}'".format(mmi_param_name))
            return
        
        logger.debug("Found {} categories with parameter".format(categories.Count))
        
        
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
        
        # Check if view has a template and ask user where to apply filters
        template_id = active_view.ViewTemplateId
        template_view = None
        template_name = None
        apply_to_template = False
        apply_to_view = False
        
        if template_id != ElementId.InvalidElementId:
            # View has a template - ask user where to apply filters
            template_view = doc.GetElement(template_id)
            if template_view:
                template_name = template_view.Name
# Ask user where to apply filters using custom WPF dialog
                xaml_file = op.join(pushbutton_dir, 'FilterSelectionDialog.xaml')
                
                class FilterSelectionDialog(forms.WPFWindow):
                    def __init__(self, xaml_file, view_name, template_name, filter_count):
                        # Load styles into Application.Resources BEFORE creating window (like MMISettingsWindow)
                        ensure_styles_loaded()
                        
                        forms.WPFWindow.__init__(self, xaml_file)
                        
                        # Ensure window Resources has the styles merged (for dark mode support)
                        try:
                            from System.Windows import Application, ResourceDictionary
                            if Application.Current is not None and Application.Current.Resources is not None:
                                # Merge Application resources into window resources if needed
                                if self.Resources is None:
                                    self.Resources = ResourceDictionary()
                                # Copy theme-aware brushes to window resources
                                for key in ['WindowBackgroundBrush', 'TextBrush', 'ControlBackgroundBrush', 
                                           'BorderBrush', 'PopupBackgroundBrush', 'BackgroundLightBrush',
                                           'BackgroundLighterBrush', 'AccentBrush', 'DisabledTextBrush']:
                                    try:
                                        if key in Application.Current.Resources.Keys:
                                            self.Resources[key] = Application.Current.Resources[key]
                                    except:
                                        pass
                        except Exception as ex:
                            logger.debug("Could not merge Application resources to window: {}".format(str(ex)))
                        
                        # Ensure options panel is visible and TextBrush is available
                        try:
                            from System.Windows import Visibility
                            from System.Windows.Media import SolidColorBrush, Colors
                            
                            # Ensure TextBrush is available
                            if self.Resources is None:
                                self.Resources = ResourceDictionary()
                            
                            if 'TextBrush' not in self.Resources.Keys:
                                if Application.Current is not None and Application.Current.Resources is not None:
                                    if 'TextBrush' in Application.Current.Resources.Keys:
                                        self.Resources['TextBrush'] = Application.Current.Resources['TextBrush']
                                    else:
                                        self.Resources['TextBrush'] = SolidColorBrush(Colors.Black)
                                else:
                                    self.Resources['TextBrush'] = SolidColorBrush(Colors.Black)
                            
                            # Ensure options panel is visible
                            options_panel = self.FindName('optionsPanel')
                            if options_panel:
                                options_panel.Visibility = Visibility.Visible
                                # Ensure all TextBlock children are visible
                                for i in range(options_panel.Children.Count):
                                    child = options_panel.Children[i]
                                    if hasattr(child, 'Visibility'):
                                        child.Visibility = Visibility.Visible
                        except Exception as ex:
                            logger.debug("Could not verify options panel: {}".format(str(ex)))
                        
                        self.viewNameRun.Text = view_name
                        self.templateNameRun.Text = template_name
                        self.filterCountRun.Text = str(filter_count)
                        self.selected_option = None
                    
                    def OkButton_Click(self, sender, e):
                        selected_item = self.optionComboBox.SelectedItem
                        if selected_item:
                            self.selected_option = selected_item.Content
                        self.Close()
                    
                
                try:
                    dialog = FilterSelectionDialog(xaml_file, active_view.Name, template_name, len(created_filters))
                    dialog.ShowDialog()
                    selected_option = dialog.selected_option
                except Exception as ex:
                    # Fallback to simple alert if custom dialog fails
                    logger.debug("Custom dialog failed, using fallback: {}".format(str(ex)))
                    import traceback
                    logger.debug(traceback.format_exc())
                    options = ["Add to Template", "Add to View", "Both"]
                    selected_option = forms.ask_for_one_item(
                        items=options,
                        default=options[0] if options else None,
                        prompt="Where would you like to apply the {} MMI filters?".format(len(created_filters)),
                        title="Apply Filters to Template or View?"
                    )
                
                if not selected_option:
                    # User cancelled
                    logger.info("User cancelled filter application")
                    return
                
                
                if selected_option == "Add to Template" or selected_option == "Both":
                    apply_to_template = True
                if selected_option == "Add to View" or selected_option == "Both":
                    apply_to_view = True
        else:
            # No template - apply to view only
            apply_to_view = True
        
        # Apply filters to template if requested
        if apply_to_template and template_view:
            try:
                with Transaction(doc, "Apply MMI Filters to View Template") as t:
                    t.Start()
                    
                    for param_filter, filter_range in created_filters:
                        apply_filter_to_view(doc, template_view, param_filter, filter_range)
                    
                    t.Commit()
                
                logger.info("Successfully applied {} filters to template '{}'".format(
                    len(created_filters), template_name))
            except Exception as ex:
                logger.error("Error applying filters to template: {}".format(ex))
                import traceback
                logger.error(traceback.format_exc())
                forms.alert(
                    "Error applying filters to template '{}':\n\n{}".format(template_name, ex),
                    title="Error"
                )
                # Continue to apply to view if that was also requested
        
        # Apply filters to the active view if requested
        if apply_to_view:
            try:
                with Transaction(doc, "Apply MMI Filters to View") as t:
                    t.Start()
                    
                    for param_filter, filter_range in created_filters:
                        apply_filter_to_view(doc, active_view, param_filter, filter_range)
                    
                    t.Commit()
                
                logger.info("Successfully applied {} filters to view '{}'".format(
                    len(created_filters), active_view.Name))
            except Exception as ex:
                logger.error("Error applying filters to view: {}".format(ex))
                import traceback
                logger.error(traceback.format_exc())
                forms.alert(
                    "Error applying filters to view '{}':\n\n{}".format(active_view.Name, ex),
                    title="Error"
                )
        
        # Show success message
        applied_locations = []
        if apply_to_template:
            applied_locations.append("template '{}'".format(template_name))
        if apply_to_view:
            applied_locations.append("view '{}'".format(active_view.Name))
        
        location_text = " and ".join(applied_locations)
        forms.show_balloon(
            header="MMI View Filters Created",
            text="Created {} view filters and applied them to {}.".format(
                len(created_filters),
                location_text
            ),
            is_new=True
        )
        
        logger.debug("Successfully created and applied {} MMI filters to {}".format(
            len(created_filters), location_text))
        
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

