# -*- coding: utf-8 -*-
__title__ = "PyRevit Colorizer"
__author__ = ""
__doc__ = """Dynamic element colorization tool for Revit.

This tool allows you to apply color overrides to elements
based on their parameter values. Features include:
- Color by instance or type parameters
- Save and load color configurations
- Predefined color schemas for common parameters
- Category-based filtering
- Exclude elements with no parameter value
"""

# Standard library imports
import sys
import os
import json
from collections import defaultdict
from re import split
from math import fabs
from random import randint
import clr
from traceback import extract_tb

from unicodedata import normalize
from unicodedata import category as unicode_category

# .NET imports
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Data')
clr.AddReference('System.Windows.Forms')
clr.AddReference('WindowsBase')
clr.AddReference('System')

# Import System namespace and its components
from System import Object, Uri, Action
import System
from System.Windows import Window, Application, FrameworkElement
from System.Windows.Controls import *
from System.Windows.Media import *
from System.Windows.Media.Imaging import BitmapImage
from System.Windows.Markup import XamlReader
from System.IO import StringReader, File
from System.Data import DataTable
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows.Data import Binding
from System.Collections.Generic import List, Dictionary
# pyRevit imports
from pyrevit import HOST_APP, revit, DB, UI
from pyrevit import forms   # By importing forms you also get references to WPF package! IT'S Very IMPORTANT !!!
import wpf         # wpf can be imported only after pyrevit.forms!
from pyrevit.script import get_logger
from pyrevit.framework import Forms, Drawing
from pyrevit.compat import get_elementid_value_func

# Revit imports
from Autodesk.Revit.DB import *

# Set up logger
logger = get_logger() # get logger and trigger debug mode using CTRL+click

# Global variable to track if a window is already running
_colorizer_window_instance = None

# Get the current document and application
doc = revit.doc
uidoc = HOST_APP.uidoc
app = revit.doc.Application
uiapp = HOST_APP.uiapp
version = int(HOST_APP.version)


# Helper function to safely get built-in category integer values
def safe_get_builtin_category_id(category_name):
    """Safely gets the integer ID of a built-in category by name"""
    try:
        category_enum = getattr(DB.BuiltInCategory, category_name, None)
        if category_enum is not None:
            return int(category_enum)
        return None
    except Exception as ex:
        logger.error("Error on safe get builtin category id: {}".format(ex))
        return None

# Categories to exclude from coloring
CAT_EXCLUDED = (
    int(DB.BuiltInCategory.OST_RoomSeparationLines),
    int(DB.BuiltInCategory.OST_Cameras),
    int(DB.BuiltInCategory.OST_CurtainGrids),
    int(DB.BuiltInCategory.OST_Elev),
    int(DB.BuiltInCategory.OST_Grids),
    int(DB.BuiltInCategory.OST_Views),
    int(DB.BuiltInCategory.OST_SitePropertyLineSegment),
    int(DB.BuiltInCategory.OST_SectionBox),
    int(DB.BuiltInCategory.OST_Lines),
    # The following categories may not exist in all Revit versions
    # int(DB.BuiltInCategory.OST_MatchLine),
    # int(DB.BuiltInCategory.OST_CenterLines),
    -2000278, 
    -1
)

# Predefined color schemas
PREDEFINED_SCHEMAS = {
    "Fire Rating": {
        "1hr": (255, 0, 0),     # Red
        "2hr": (0, 255, 0),     # Green
        "3hr": (0, 0, 255),     # Blue
        "None": (128, 128, 128) # Gray
    },
    "MMI": {
        "A": (255, 0, 0),       # Red
        "B": (255, 165, 0),     # Orange
        "C": (255, 255, 0),     # Yellow
        "D": (0, 128, 0),       # Green
        "None": (128, 128, 128) # Gray
    }
}

# Helper Classes 
class ValuesInfo():
    def __init__(self, para, val, idt, num1, num2, num3):
        self.par = para
        self.value = val
        self.name = strip_accents(para.Definition.Name)
        self.ele_id = List[DB.ElementId]()
        self.ele_id.Add(idt)
        self.n1 = num1
        self.n2 = num2
        self.n3 = num3
        self.colour = Drawing.Color.FromArgb(self.n1, self.n2, self.n3)
        self.values_double = []
        if para.StorageType == DB.StorageType.Double:
            self.values_double.append(para.AsDouble())
        elif para.StorageType == DB.StorageType.ElementId:
            self.values_double.append(para.AsElementId())

class ParameterInfo:
    """Stores information about a parameter"""
    def __init__(self, param_type, parameter):
        self.param_type = param_type  # 0 = Instance, 1 = Type
        self.parameter = parameter
        self.definition = parameter.Definition
        self.name = parameter.Definition.Name
        self.storage_type = parameter.StorageType
    
    def __str__(self):
        param_type_str = "Type" if self.param_type == 1 else "Instance"
        return "{} ({})".format(self.name, param_type_str)

class CategoryInfo:
    """Stores information about a category"""
    def __init__(self, category, parameters):
        self.name = category.Name
        self.category = category
        self.id = category.Id
        get_elementid_value = get_elementid_value_func()
        self.int_id = get_elementid_value(category.Id)
        self.parameters = parameters
    
    def __str__(self):
        return self.name

# UI Event Handlers
class ApplyColorsHandler(UI.IExternalEventHandler):
    """Handles applying colors to elements"""
    def __init__(self, colorizer_ui):
        self.ui = colorizer_ui
    
    def Execute(self, uiapp):
        try:
            active_doc = uiapp.ActiveUIDocument.Document
            view = self.ui.current_view
            if not view:
                self.ui.status_text.Text = "No active view available"
                self.logger.info("No active view available")
                return
                
            # Get the solid fill pattern for overrides
            solid_fill_id = self.get_solid_fill_pattern_id(active_doc)
            
            with revit.Transaction("Apply colors to elements"):
                # Get selected category
                selected_categories = [item for item in self.ui.category_items if item.IsSelected]
                if not selected_categories:
                    self.ui.status_text.Text = "No category selected"
                    self.logger.info("No category selected")
                    return
                
                for value_info in self.ui.values_list:
                    if not value_info.element_ids:
                        continue
                        
                    # Create override settings
                    ogs = DB.OverrideGraphicSettings()
                    color = DB.Color(value_info.r, value_info.g, value_info.b)
                    ogs.SetProjectionLineColor(color)
                    ogs.SetSurfaceForegroundPatternColor(color)
                    ogs.SetCutForegroundPatternColor(color)
                    
                    if solid_fill_id:
                        ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                        ogs.SetCutForegroundPatternId(solid_fill_id)
                    
                    # Set line pattern to solid
                    ogs.SetProjectionLinePatternId(DB.ElementId(-1))
                    
                    # Apply override to each element
                    for element_id in value_info.element_ids:
                        try:
                            view.SetElementOverrides(element_id, ogs)
                        except Exception as ex:
                            self.logger.error("Error on set element overrides: {}".format(ex))
                
                self.ui.status_text.Text = "Colors applied successfully"
        except Exception as ex:
            self.logger.error("Error on apply colors: {}".format(ex))
            
    def get_solid_fill_pattern_id(self, document):
        """Gets the solid fill pattern ID from the document"""
        patterns = DB.FilteredElementCollector(document).OfClass(DB.FillPatternElement)
        for pattern in patterns:
            if pattern.GetFillPattern().IsSolidFill:
                return pattern.Id
        return None
    
    def GetName(self):
        return "Apply Colors Handler"

class ResetColorsHandler(UI.IExternalEventHandler):
    """Handles resetting colors of elements"""
    def __init__(self, colorizer_ui):
        self.ui = colorizer_ui
        
    def Execute(self, uiapp):
        try:
            active_doc = uiapp.ActiveUIDocument.Document
            view = self.ui.current_view
            if not view:
                self.ui.status_text.Text = "No active view available"
                self.logger.info("No active view available")
                return
                
            # Default override settings (no overrides)
            ogs = DB.OverrideGraphicSettings()
            
            with revit.Transaction("Reset element colors"):
                # Reset all elements that have been colored
                all_colored_ids = []
                for value_info in self.ui.values_list:
                    all_colored_ids.extend(value_info.element_ids)
                
                for element_id in all_colored_ids:
                    try:
                        view.SetElementOverrides(element_id, ogs)
                    except Exception as ex:
                        self.logger.error("Error on reset colors: {}".format(ex))
                
                self.ui.status_text.Text = "Colors reset successfully"
        except Exception as ex:
            self.logger.error("Error on reset colors: {}".format(ex))
    
    def GetName(self):
        return "Reset Colors Handler"

def random_color():
    """Generates a random RGB color"""
    r = randint(0, 230)
    g = randint(0, 230)
    b = randint(0, 230)
    return r, g, b

def get_used_categories(active_view, excluded_cats=None):
    """Gets categories used in the active view"""
    if not excluded_cats:
        excluded_cats = CAT_EXCLUDED
        
    # Get all elements in view
    collector = DB.FilteredElementCollector(doc, active_view.Id) \
                  .WhereElementIsNotElementType() \
                  .WhereElementIsViewIndependent() \
                  .ToElements()
    
    categories = {}
    get_elementid_value = get_elementid_value_func()
    
    for element in collector:
        if not element.Category:
            continue
            
        # Skip excluded categories
        category_id = get_elementid_value(element.Category.Id)
        if category_id in excluded_cats or category_id >= -1:
            continue
            
        # If we've already processed this category, skip
        if category_id in categories:
            continue
            
        # Process instance parameters
        instance_params = []
        for param in element.Parameters:
            if param.Definition and param.StorageType != getattr(DB.StorageType, "None", None):
                instance_params.append(ParameterInfo(0, param))
        
        # Process type parameters
        type_params = []
        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            for param in element_type.Parameters:
                if param.Definition and param.StorageType != getattr(DB.StorageType, "None", None):
                    type_params.append(ParameterInfo(1, param))
        
        # Combine and sort parameters
        all_params = sorted(instance_params + type_params, key=lambda x: x.name)
        
        # Add category to dictionary
        categories[category_id] = CategoryInfo(element.Category, all_params)
    
    # Convert dictionary to sorted list
    category_list = sorted(categories.values(), key=lambda x: x.name)
    return category_list

# Add helper functions for parameter values
def get_parameter_values(category, param, new_view):
    """Gets all unique values for a parameter in the given category"""
    try:
        
        logger.info("Getting parameter values for category: {} (ID: {})".format(
            category.name, category.int_id))
        
        # Get the category ID as an ElementId
        category_id = DB.ElementId(category.int_id)
        logger.info("Created ElementId: {}".format(category_id))
        
        # Get elements of this category in the current view using OfCategoryId instead of OfCategory
        collector = (DB.FilteredElementCollector(doc, new_view.Id)
                    .OfCategoryId(category_id)
                    .WhereElementIsNotElementType()
                    .WhereElementIsViewIndependent())
        
        # Count elements found - do this before converting to elements
        element_count = collector.GetElementCount()
        logger.info("Found {} elements in category".format(element_count))
        
        # Convert to elements after getting count
        elements = collector.ToElements()
        
        # Initialize empty list and set for used colors
        list_values = []
        used_colors = set()
        
        # Store references to frequently used types/constants
        storage_double = DB.StorageType.Double
        
        for ele in elements:
            try:
                # Get element or its type based on parameter type
                ele_par = ele if param.param_type != 1 else doc.GetElement(ele.GetTypeId())
                
                for pr in ele_par.Parameters:
                    if pr.Definition.Name == param.name:
                        value = get_parameter_value(pr) or "None"
                        match = [x for x in list_values if x.value == value]
                        if match:
                            # Check if ele_id is a List[DB.ElementId] or a regular list
                            if hasattr(match[0].ele_id, 'Add'):
                                # It's a proper .NET List, use Add method
                                match[0].ele_id.Add(ele.Id)
                            else:
                                # It's a regular Python list, use append
                                match[0].ele_id.append(ele.Id)
                                
                            if pr.StorageType == storage_double:
                                match[0].values_double.append(pr.AsDouble())
                        else:
                            while True:
                                r, g, b = random_color()
                                if (r, g, b) not in used_colors:
                                    used_colors.add((r, g, b))
                                    try:
                                        # Create ValuesInfo with proper DB reference  
                                        val = ValuesInfo(pr, value, ele.Id, r, g, b)
                                        list_values.append(val)
                                        break
                                    except Exception as ex:
                                        logger.error("Error creating ValuesInfo: {}".format(ex))
                                        raise
                        break
            except Exception as ex:
                logger.error("Error processing element: {}".format(ex))
                continue
                    
        # Separate None values
        none_values = [x for x in list_values if x.value == "None"]
        logger.info("Found {} None values".format(len(none_values)))
        list_values = [x for x in list_values if x.value != "None"]
        logger.info("Found {} non-None values".format(len(list_values)))
        # Sort values
        list_values = sorted(list_values, key=lambda x: x.value, reverse=False)
        if len(list_values) > 1:
            try:
                first_value = list_values[0].value
                indx_del = get_index_units(first_value)
                if indx_del == 0:
                    list_values = sorted(list_values, key=lambda x: safe_float(x.value))
                elif 0 < indx_del < len(first_value):
                    list_values = sorted(list_values, key=lambda x: safe_float(x.value[:-indx_del]))
            except ValueError as ve:
                logger.error("ValueError during sorting: {}".format(ve))
            except Exception as ex:
                logger.error("Error on sort values: {}".format(ex))
                
        # Add None values at the end if they contain elements
        if none_values and any(len(x.ele_id) > 0 for x in none_values):
            list_values.extend(none_values)
                    
        logger.info("Returning {} unique values".format(len(list_values)))
        return list_values
    except Exception as ex:
        import traceback
        tb = traceback.format_exc()
        logger.error("Error in get_parameter_values: {} - Traceback: {}".format(ex, tb))
        raise

    
def random_color():
    """Generates a random RGB color"""
    r = randint(0, 230)
    g = randint(0, 230)
    b = randint(0, 230)
    return r, g, b

def get_parameter_value(para):
    """Extract value from a parameter based on its storage type"""
    if not para.HasValue:
        return "None"
    if para.StorageType == DB.StorageType.Double:
        return get_double_value(para)
    if para.StorageType == DB.StorageType.ElementId:
        return get_elementid_value(para)
    if para.StorageType == DB.StorageType.Integer:
        return get_integer_value(para)
    if para.StorageType == DB.StorageType.String:
        return para.AsString() or "None"
    else:
        return "None"

def get_double_value(para):
    """Extract value from a double parameter"""
    return para.AsValueString() or str(para.AsDouble())

def get_elementid_value(para):
    """Extract value from an ElementId parameter"""
    id_val = para.AsElementId()
    elementid_value = get_elementid_value_func()
    if elementid_value(id_val) >= 0:
        element = doc.GetElement(id_val)
        return element.Name if element else "None"
    else:
        return "None"

def get_integer_value(para):
    """Extract value from an integer parameter"""
    if version > 2021:
        param_type = para.Definition.GetDataType()
        if DB.SpecTypeId.Boolean.YesNo == param_type:
            return "Yes" if para.AsInteger() == 1 else "No"
        else:
            return para.AsValueString() or str(para.AsInteger())
    else:
        param_type = para.Definition.ParameterType
        if DB.ParameterType.YesNo == param_type:
            return "Yes" if para.AsInteger() == 1 else "No"
        else:
            return para.AsValueString() or str(para.AsInteger())

def safe_float(value):
    """Safely convert a string to float"""
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def get_index_units(value):
    """Get the index of units in a string value"""
    try:
        # Find where the numeric part ends and units begin
        if not value:
            return 0
            
        for i, c in enumerate(value):
            if i > 0 and not (c.isdigit() or c == '.' or c == ','):
                return len(value) - i
        return 0
    except:
        return 0
    
def strip_accents(text):
    return ''.join(char for char in normalize('NFKD', text) if unicode_category(char) != 'Mn')
  
# Main class for the Colorizer UI
class RevitColorizerWindow(Window):
    """Main WPF window for the Revit Colorizer tool"""
        
    def __init__(self):
        try:
            # Add initialization tracking
            logger.info("=============== WINDOW INITIALIZATION STARTED ===============")
            
            # Create XAML file for our UI
            xaml_file = os.path.join(
                os.path.dirname(__file__), 
                "RevitColorizerWindow.xaml"
            )

            wpf.LoadComponent(self, xaml_file)
            
            # Store logger reference
            self.logger = get_logger()
            
            # Store DB reference to ensure it's accessible in methods
            self.DB = DB
            self.logger.info("DB reference stored: {}".format(self.DB))

            # Store doc reference to ensure it's accessible in methods
            self.doc = doc
            self.logger.info("doc reference stored: {}".format(self.doc))
            
            # Store System reference to ensure it's accessible in methods
            self.System = System
            self.logger.info("System reference stored: {}".format(self.System))
            
            # store reference to helper functions
            self.get_parameter_value = get_parameter_value
            self.get_parameter_values = get_parameter_values
            self.get_index_units = get_index_units
            self.get_elementid_value = get_elementid_value
            self.get_double_value = get_double_value
            self.get_integer_value = get_integer_value
            self.safe_float = safe_float
            self.strip_accents = strip_accents
            
            # Store reference to randint from random module
            self.randint = randint
            
            # Initialize UI components
            self.setup_ui_components()
            
            # Set up event handlers
            self.apply_colors_handler = ApplyColorsHandler(self)
            self.reset_colors_handler = ResetColorsHandler(self)
            self.apply_colors_event = UI.ExternalEvent.Create(self.apply_colors_handler)
            self.reset_colors_event = UI.ExternalEvent.Create(self.reset_colors_handler)
            
            # Initialize state variables
            self.current_view = self.get_active_view()
            self.values_list = []
            
            # Load categories
            self.load_categories()

            self.Show()
        except Exception as ex:
            self.logger.error("Error initializing RevitColorizerWindow: {}".format(ex))
        
    def setup_ui_components(self):
        """Set up the UI components from XAML"""
        # Main components - these should match the names in the XAML file
        self.category_listbox = self.categoryListBox
        self.parameter_selector = self.parameterSelector
        self.values_listbox = self.valuesListBox
        self.settings_combo = self.settingsComboBox
        
        # Radio buttons for parameter type
        self.instance_radio = self.instanceRadioButton
        self.type_radio = self.typeRadioButton
        
        # Toggle switches (styled checkboxes)
        self.override_projection_checkbox = self.overrideProjectionCheckbox
        self.show_elements_checkbox = self.showElementsCheckbox
        self.keep_overrides_checkbox = self.keepOverridesCheckbox
        
        # Buttons
        self.apply_button = self.applyButton
        self.reset_button = self.resetButton
        self.save_button = self.saveButton
        self.load_button = self.loadButton
        self.add_filters_button = self.addFiltersButton
        self.remove_filters_button = self.removeFiltersButton
        self.close_button = self.closeButton
        
        # Status text
        self.status_text = self.statusText
        
        # Set up event handlers with strong references
        def create_handler(method_name):
            method = getattr(self, method_name)
            # Store critical references
            DB_ref = self.DB
            doc_ref = self.doc
            logger_ref = self.logger
            
            def handler(sender, args):
                try:
                    # Log event
                    logger_ref.info("Event triggered: {}".format(method_name))
                    
                    # Restore critical references if needed
                    if not hasattr(self, 'DB') or self.DB is None:
                        self.DB = DB_ref
                    if not hasattr(self, 'doc') or self.doc is None:
                        self.doc = doc_ref
                        
                    # Call the actual method
                    return method(sender, args)
                except Exception as ex:
                    logger_ref.error("Error in {}: {}".format(method_name, str(ex)))
            return handler
        
        # Store all handlers as instance variables to prevent garbage collection
        self.handlers = {}
        
        # Button handlers
        self.handlers['on_apply_colors'] = create_handler('on_apply_colors')
        self.handlers['on_reset_colors'] = create_handler('on_reset_colors')
        self.handlers['on_save_config'] = create_handler('on_save_config')
        self.handlers['on_load_config'] = create_handler('on_load_config')
        self.handlers['on_close'] = create_handler('on_close')
        self.handlers['on_add_filters'] = create_handler('on_add_filters')
        self.handlers['on_remove_filters'] = create_handler('on_remove_filters')
        
        # Other UI handlers
        self.handlers['on_category_selection_changed'] = create_handler('on_category_selection_changed')
        self.handlers['on_parameter_type_changed'] = create_handler('on_parameter_type_changed')
        self.handlers['on_radio_button_click'] = create_handler('on_radio_button_click')
        self.handlers['on_parameter_selected'] = create_handler('on_parameter_selected')
        self.handlers['on_parameter_dropdown_closed'] = create_handler('on_parameter_dropdown_closed')
        self.handlers['on_value_selected'] = create_handler('on_value_selected')
        self.handlers['on_value_double_click'] = create_handler('on_value_double_click')
        self.handlers['on_window_loaded'] = create_handler('on_window_loaded')
        
        # Checkbox handlers - create separate handlers for checked/unchecked
        self.handlers['on_option_changed_checked'] = create_handler('on_option_changed')
        self.handlers['on_option_changed_unchecked'] = create_handler('on_option_changed')
                
        # Connect all the handlers
        self.apply_button.Click += self.handlers['on_apply_colors']
        self.reset_button.Click += self.handlers['on_reset_colors']
        self.save_button.Click += self.handlers['on_save_config']
        self.load_button.Click += self.handlers['on_load_config']
        self.close_button.Click += self.handlers['on_close']
        
        # Set up ListBox selection change event
        self.category_listbox.SelectionChanged += self.handlers['on_category_selection_changed']
        
        # Set up radio button handlers
        self.instance_radio.Checked += self.handlers['on_parameter_type_changed']
        self.type_radio.Checked += self.handlers['on_parameter_type_changed']
        self.instance_radio.Click += self.handlers['on_radio_button_click']
        self.type_radio.Click += self.handlers['on_radio_button_click']
        
        # Set up toggle switch handlers
        self.override_projection_checkbox.Checked += self.handlers['on_option_changed_checked']
        self.override_projection_checkbox.Unchecked += self.handlers['on_option_changed_unchecked']
        self.show_elements_checkbox.Checked += self.handlers['on_option_changed_checked']
        self.show_elements_checkbox.Unchecked += self.handlers['on_option_changed_unchecked']
        self.keep_overrides_checkbox.Checked += self.handlers['on_option_changed_checked']
        self.keep_overrides_checkbox.Unchecked += self.handlers['on_option_changed_unchecked']
        
        # Set up filter button handlers
        self.add_filters_button.Click += self.handlers['on_add_filters']
        self.remove_filters_button.Click += self.handlers['on_remove_filters']
        
        # Set up parameter selector
        self.parameter_selector.SelectionChanged += self.handlers['on_parameter_selected']
        self.parameter_selector.DropDownClosed += self.handlers['on_parameter_dropdown_closed']
        
        # Set up values listbox for color selection
        self.values_listbox.SelectionChanged += self.handlers['on_value_selected']
        self.values_listbox.MouseDoubleClick += self.handlers['on_value_double_click']
        
        # Add window loaded event handler
        self.Loaded += self.handlers['on_window_loaded']
        
        # Set window title
        self.Title = "PyRevit Colorizer"
    
    def get_active_view(self):
        """Get the current active view"""
        selected_view = doc.ActiveView
        if selected_view.ViewType == DB.ViewType.ProjectBrowser or selected_view.ViewType == DB.ViewType.SystemBrowser:
            selected_view = doc.GetElement(uidoc.GetOpenUIViews()[0].ViewId)
        
        if not selected_view.CanUseTemporaryVisibilityModes():
            self.status_text.Text = "Visibility settings cannot be modified in {} views".format(selected_view.ViewType)
            self.logger.info("Visibility settings cannot be modified in {} views".format(selected_view.ViewType))
            return None
        
        return selected_view
    
    def load_categories(self):
        """Load categories from the current view"""
        if not self.current_view:
            return
            
        self.categories = get_used_categories(self.current_view)
        
        # Clear the category listbox
        self.category_listbox.Items.Clear()
        
        # Create observable collection with IsSelected property for binding
        from System.ComponentModel import INotifyPropertyChanged
        
        # Define a simple class for category items with selection state
        class CategoryItem(INotifyPropertyChanged):
            def __init__(self, category):
                self.category = category
                self._is_selected = False
                # Standard implementation of INotifyPropertyChanged
                self._propertyChangedHandlers = []
                self.owner = None  # Reference to the main window
            
            @property
            def name(self):
                return self.category.name
                
            @property
            def category_info(self):
                return self.category
                
            @property
            def IsSelected(self):
                return self._is_selected
                
            @IsSelected.setter
            def IsSelected(self, value):
                if self._is_selected != value:
                    self._is_selected = value
                    self.OnPropertyChanged("IsSelected")
                    # When selection changes, notify the main window to refresh parameters
                    if self.owner and value:  # Only on selection, not deselection
                        self.owner.refresh_parameter_selection()
            
            # INotifyPropertyChanged implementation
            def add_PropertyChanged(self, handler):
                self._propertyChangedHandlers.append(handler)
                
            def remove_PropertyChanged(self, handler):
                self._propertyChangedHandlers.remove(handler)
            
            def OnPropertyChanged(self, property_name):
                for handler in self._propertyChangedHandlers:
                    handler(self, PropertyChangedEventArgs(property_name))
                    
        # Add categories to listbox
        self.category_items = []
        for category in self.categories:
            item = CategoryItem(category)
            item.owner = self  # Set reference to main window
            self.category_items.append(item)
            self.category_listbox.Items.Add(item)
            
        # Update status
        self.status_text.Text = "Found {} categories in the current view".format(len(self.categories))
        self.logger.info("Found {} categories in the current view".format(len(self.categories)))
    
    def refresh_parameter_selection(self):
        """Refresh parameters based on current category selection"""
        try:
            # Get selected categories
            selected_categories = [item for item in self.category_items if item.IsSelected]
            selected_count = len(selected_categories)
            
            if selected_count == 0:
                self.status_text.Text = "Please select at least one category"
                self.logger.info("Please select at least one category")
                self.parameter_selector.Items.Clear()
                self.values_listbox.Items.Clear()
                return
                
            # Update status
            self.status_text.Text = "Selected {} categories".format(selected_count)
            self.logger.info("Selected {} categories".format(selected_count))
            
            # Set to instance parameter type if needed
            current_instance_state = self.instance_radio.IsChecked
            if not current_instance_state:
                self.instance_radio.IsChecked = True
                # We need to force the parameter type changed event
                # since simply setting IsChecked might not trigger it
                self.type_radio.IsChecked = False 
            
            # Load parameters for the first selected category
            if selected_count > 0:
                self.load_parameters_for_category(selected_categories[0].category_info)
        except Exception as ex:
            self.logger.error("Error refreshing parameter selection: {}".format(ex))
    
    def on_category_selection_changed(self, sender, args):
        """Handle category selection changes - this is now primarily a backup"""
        try:
            # We'll rely primarily on property changed events, but keep this as a backup
            pass
        except Exception as ex:
            self.logger.error("Error on category selection changed: {}".format(ex))
    
    def load_parameters_for_category(self, category):
        """Load parameters for the selected category"""
        try:
            # Clear existing parameters
            self.parameter_selector.Items.Clear()
            self.values_listbox.Items.Clear()
            
            # Get the parameter type (instance or type)
            param_type = 0 if self.instance_radio.IsChecked else 1
            param_type_name = "instance" if param_type == 0 else "type"
            
            # Filter parameters by type
            filtered_params = [p for p in category.parameters if p.param_type == param_type]
            
            # Add parameters to selector
            for param in filtered_params:
                self.parameter_selector.Items.Add(param)
                
            # Update status
            param_count = len(filtered_params)
            if param_count > 0:
                self.status_text.Text = "Found {} {} parameters".format(param_count, param_type_name)
                self.logger.info("Found {} {} parameters".format(param_count, param_type_name))
                # Select first parameter
                if self.parameter_selector.Items.Count > 0:
                    self.parameter_selector.SelectedIndex = 0
                    # Automatically load values for the first parameter
                    self.on_parameter_selected(None, None)
            else:
                self.status_text.Text = "No {} parameters found".format(param_type_name)
                self.logger.info("No {} parameters found".format(param_type_name))
                
        except Exception as ex:
            self.logger.error("Error loading parameters for category: {}".format(ex))
    
    def on_parameter_type_changed(self, sender, args):
        """Handle parameter type change (instance/type)"""
        try:
            # Add debugging to see which radio button changed
            radio = sender
            is_checked = radio.IsChecked
            
            # Only proceed if a radio button was checked (not unchecked)
            if not is_checked:
                return
                
            # Get selected categories
            selected_categories = [item for item in self.category_items if item.IsSelected]
            
            if selected_categories:
                # Reload parameters for the first selected category
                self.load_parameters_for_category(selected_categories[0].category_info)
        except Exception as ex:
            self.logger.error("Error on parameter type changed: {}".format(ex))
    
    def on_option_changed(self, sender, args):
        """Handle option checkbox changes"""
        try:
            # Get the checkbox that triggered the event
            checkbox = sender
            self.logger.info("Checkbox '{}' state changed to: {}".format(checkbox.Name, checkbox.IsChecked))
            
            # Force UI update for the checkbox
            checkbox.UpdateLayout()
            
            # Update status with current options
            options = []
            if self.override_projection_checkbox.IsChecked:
                options.append("Override Projection/Cut Lines")
            if self.show_elements_checkbox.IsChecked:
                options.append("Show Elements in Properties")
            if self.keep_overrides_checkbox.IsChecked:
                options.append("Keep Overrides")
                
            if options:
                self.status_text.Text = "Options enabled: {}".format(', '.join(options))
                self.logger.info("Options enabled: {}".format(', '.join(options)))
            else:
                self.status_text.Text = "All options disabled"
                self.logger.info("All options disabled")
            
            # If colors are already applied, update them with new settings
            if self.values_list and len(self.values_list) > 0:
                self.on_apply_colors(None, None)
                
        except Exception as ex:
            self.logger.error("Error on option changed: {}".format(ex))
    
    def on_add_filters(self, sender, args):
        """Handle adding view filters"""
        try:
            self.status_text.Text = "Adding view filters is not implemented yet"
            self.logger.info("Adding view filters is not implemented yet")
        except Exception as ex:
            self.logger.error("Error on add filters: {}".format(ex))
    
    def on_remove_filters(self, sender, args):
        """Handle removing view filters"""
        try:
            self.status_text.Text = "Removing view filters is not implemented yet"
            self.logger.info("Removing view filters is not implemented yet")
        except Exception as ex:
            self.logger.error("Error on remove filters: {}".format(ex))
    
    def on_close(self, sender, args):
        """Handle close button click"""
        try:
            self.Close()
        except Exception as ex:
            self.logger.error("Error on close: {}".format(ex))
    
    def on_window_loaded(self, sender, args):
        """Called when the window is fully loaded"""
        try:
            # Force UI to update
            self.parameter_selector.UpdateLayout()
            
            # Set default toggle states
            self.override_projection_checkbox.IsChecked = False
            self.show_elements_checkbox.IsChecked = False
            self.keep_overrides_checkbox.IsChecked = False
            
            # Load saved configurations into settings dropdown
            self.load_saved_configurations()
        except Exception as ex:
            self.logger.error("Error on window loaded: {}".format(ex))
            
    def load_saved_configurations(self):
        """Load saved configurations into settings dropdown"""
        try:
            app_data_path = os.path.expanduser("~\\AppData\\Roaming\\pyRevit\\PyRevitColorizer")
            if not os.path.exists(app_data_path):
                return
                
            # Get all .rvtcolor files
            config_files = [f for f in os.listdir(app_data_path) if f.endswith('.rvtcolor')]
            
            # Add each configuration to the dropdown
            for config_file in config_files:
                setting_name = os.path.splitext(config_file)[0]
                item = Forms.ComboBoxItem()
                item.Content = setting_name
                self.settings_combo.Items.Add(item)
                
            # Select Default
            self.settings_combo.SelectedIndex = 0
        except Exception as ex:
            self.logger.error("Error on load saved configurations: {}".format(ex))
    
    def on_apply_colors(self, sender, args):
        """Apply colors to elements"""
        if not self.values_list:
            self.status_text.Text = "No values to apply colors to"
            self.logger.info("No values to apply colors to")
            return
        
        # Check which options are enabled
        use_projection = self.override_projection_checkbox.IsChecked
        
        # Apply colors with appropriate options
        self.status_text.Text = "Applying colors..."
        self.logger.info("Applying colors...")
        self.apply_colors_event.Raise()
    
    def on_parameter_selected(self, sender, args):
        """Handle parameter selection"""
        try:
            # Debug state at the start of parameter selection
            self.debug_state("on_parameter_selected start")
            
            self.values_listbox.Items.Clear()
            self.values_list = []
            
            # Get selected categories
            selected_categories = [item for item in self.category_items if item.IsSelected]
            if not selected_categories:
                self.status_text.Text = "Please select at least one category"
                self.logger.info("Please select at least one category")
                return
                
            selected_parameter = self.parameter_selector.SelectedItem
            
            if not selected_parameter:
                self.status_text.Text = "Please select a parameter"
                self.logger.info("Please select a parameter")
                return
            
            # Get values for the selected parameter
            self.status_text.Text = "Loading values for parameter: {}".format(selected_parameter.name)
            self.logger.info("Loading values for parameter: {}".format(selected_parameter.name))
            
            try:
                
                # Use get_parameter_values to get values list
                self.values_list = get_parameter_values(selected_categories[0].category_info, 
                                                        selected_parameter, 
                                                        self.current_view)
                self.logger.info("Found {} values for parameter".format(len(self.values_list)))
                
            except Exception as ex:
                self.logger.error("Error getting parameter values: {}".format(ex))
                self.status_text.Text = "Error: {}".format(str(ex))
                return

            # Populate values listbox
            for value_info in self.values_list:
                self.values_listbox.Items.Add(value_info)
                
            if not self.values_list:
                self.status_text.Text = "No values found for parameter: {}".format(selected_parameter.name)
                self.logger.info("No values found for parameter: {}".format(selected_parameter.name))
            else:
                self.status_text.Text = "Found {} unique values".format(len(self.values_list))
                self.logger.info("Found {} unique values".format(len(self.values_list)))
                
                # Force a refresh to ensure the UI updates
                self.values_listbox.Items.Refresh()
                
                # Automatically apply colors if there are values
                # Always auto-apply when ticking a checkbox
                self.on_apply_colors(None, None)
        except Exception as ex:
            self.logger.error("Error on parameter selected: {}".format(ex))
            self.status_text.Text = "Error: {}".format(str(ex))

    def on_parameter_dropdown_closed(self, sender, args):
        """Handle parameter dropdown closing - another chance to catch selection"""
        try:
            selected_parameter = self.parameter_selector.SelectedItem
            if selected_parameter:
                # Manually call the parameter selection handler
                self.on_parameter_selected(sender, args)
        except Exception as ex:
            self.logger.error("Error on parameter dropdown closed: {}".format(ex))
            
    def on_reset_colors(self, sender, args):
        """Reset colors on elements"""
        self.status_text.Text = "Resetting colors..."
        self.logger.info("Resetting colors...")
        self.reset_colors_event.Raise()
    
    def on_save_config(self, sender, args):
        """Save configuration to file"""
        if not self.values_list:
            self.status_text.Text = "No configuration to save"
            self.logger.info("No configuration to save")
            return
            
        try:
            # Get selected categories and parameter
            selected_categories = [item for item in self.category_items if item.IsSelected]
            if not selected_categories:
                self.status_text.Text = "Please select at least one category"
                self.logger.info("Please select at least one category")
                return
                
            selected_parameter = self.parameter_selector.SelectedItem
            if not selected_parameter:
                self.status_text.Text = "Please select a parameter"
                self.logger.info("Please select a parameter")
                return
                
            # Get the name for the config (from settings combobox or prompt for new name)
            setting_name = self.settings_combo.Text
            if not setting_name or setting_name == "<Default>":
                # Prompt for a name
                input_dialog = Forms.InputBox("Enter a name for this configuration:", "Save Configuration", "")
                if not input_dialog:
                    self.status_text.Text = "Save cancelled"
                    self.logger.info("Save cancelled")
                    return
                setting_name = input_dialog
                
            # Create configuration
            config = {
                "name": setting_name,
                "category": selected_categories[0].name,
                "parameter": selected_parameter.name,
                "instance_parameter": self.instance_radio.IsChecked,
                "options": {
                    "override_projection": self.override_projection_checkbox.IsChecked,
                    "show_elements": self.show_elements_checkbox.IsChecked,
                    "keep_overrides": self.keep_overrides_checkbox.IsChecked
                },
                "values": []
            }
            
            # Add values
            for value_info in self.values_list:
                config["values"].append({
                    "value": value_info.value,
                    "color": [value_info.r, value_info.g, value_info.b]
                })
            
            # Get the settings directory
            app_data_path = os.path.expanduser("~\\AppData\\Roaming\\pyRevit\\PyRevitColorizer")
            if not os.path.exists(app_data_path):
                os.makedirs(app_data_path)
                
            # Save to file
            config_file = os.path.join(app_data_path, "{}.rvtcolor".format(setting_name))
            with open(config_file, 'w') as f:
                json.dump(config, f, indent=4)
                
            # Add to settings dropdown if it's not already there
            found = False
            for i in range(self.settings_combo.Items.Count):
                if self.settings_combo.Items[i].Content == setting_name:
                    found = True
                    break
                    
            if not found:
                item = Forms.ComboBoxItem()
                item.Content = setting_name
                self.settings_combo.Items.Add(item)
                self.settings_combo.SelectedItem = item
                
            self.status_text.Text = "Configuration '{}' saved successfully".format(setting_name)
            self.logger.info("Configuration '{}' saved successfully".format(setting_name))
        except Exception as ex:
            self.logger.error("Error on save config: {}".format(ex))
    
    def on_load_config(self, sender, args):
        """Load configuration from file"""
        try:
            # Check if a setting is selected
            selected_item = self.settings_combo.SelectedItem
            if selected_item and selected_item.Content != "<Default>":
                # Load the selected configuration
                setting_name = selected_item.Content
            else:
                # Open dialog to select a file
                open_dialog = Forms.OpenFileDialog()
                open_dialog.Title = "Load Color Configuration"
                open_dialog.Filter = "Color Configuration (*.rvtcolor)|*.rvtcolor"
                open_dialog.InitialDirectory = os.path.expanduser("~\\AppData\\Roaming\\pyRevit\\PyRevitColorizer")
                
                if open_dialog.ShowDialog() != Forms.DialogResult.OK:
                    return
                    
                setting_name = os.path.splitext(os.path.basename(open_dialog.FileName))[0]
                config_file = open_dialog.FileName
            
            # If loading from settings combobox
            if not 'config_file' in locals():
                app_data_path = os.path.expanduser("~\\AppData\\Roaming\\pyRevit\\PyRevitColorizer")
                config_file = os.path.join(app_data_path, "{}.rvtcolor".format(setting_name))
                if not os.path.exists(config_file):
                    self.status_text.Text = "Configuration file for '{}' not found".format(setting_name)
                    self.logger.info("Configuration file for '{}' not found".format(setting_name))
                    return
            
            # Load configuration
            with open(config_file, 'r') as f:
                config = json.load(f)
            
            # Set options from config
            if "options" in config:
                if "override_projection" in config["options"]:
                    self.override_projection_checkbox.IsChecked = config["options"]["override_projection"]
                if "show_elements" in config["options"]:
                    self.show_elements_checkbox.IsChecked = config["options"]["show_elements"]
                if "keep_overrides" in config["options"]:
                    self.keep_overrides_checkbox.IsChecked = config["options"]["keep_overrides"]
            
            # Set parameter type
            instance_param = config.get("instance_parameter", True)
            self.instance_radio.IsChecked = instance_param
            self.type_radio.IsChecked = not instance_param
            
            # Find and select the category
            for cat_item in self.category_items:
                if cat_item.name == config["category"]:
                    cat_item.IsSelected = True
                    # Trigger parameter loading
                    self.load_parameters_for_category(cat_item.category_info)
                    break
            else:
                self.status_text.Text = "Category '{}' not found in current view".format(config["category"])
                self.logger.info("Category '{}' not found in current view".format(config["category"]))
                return
            
            # Find parameter in the loaded parameters
            for i in range(self.parameter_selector.Items.Count):
                param = self.parameter_selector.Items[i]
                if param.name == config["parameter"]:
                    self.parameter_selector.SelectedIndex = i
                    break
            else:
                self.status_text.Text = "Parameter '{}' not found in selected category".format(config["parameter"])
                self.logger.info("Parameter '{}' not found in selected category".format(config["parameter"]))
                return
            
            # Apply colors to values if we have them
            if self.values_list and "values" in config:
                value_to_color = {val["value"]: val["color"] for val in config["values"]}
                
                for value_info in self.values_list:
                    if value_info.value in value_to_color:
                        color = value_to_color[value_info.value]
                        value_info.r = color[0]
                        value_info.g = color[1]
                        value_info.b = color[2]
                        value_info.color = Color.FromRgb(color[0], color[1], color[2])
                
                # Refresh listbox
                self.values_listbox.Items.Refresh()
            
            self.status_text.Text = "Configuration '{}' loaded successfully".format(setting_name)
            self.logger.info("Configuration '{}' loaded successfully".format(setting_name))
        except Exception as ex:
            self.logger.error("Error on load config: {}".format(ex))

    def on_value_selected(self, sender, args):
        """Handle value selection in listbox"""
        try:
            selected_value = self.values_listbox.SelectedItem
            if not selected_value:
                return
                
            # Just update status text with selection info
            self.status_text.Text = "Selected value: {}".format(selected_value.value)
            self.logger.info("Selected value: {}".format(selected_value.value))
        except Exception as ex:
            self.logger.error("Error on value selected: {}".format(ex))
            
    def on_value_double_click(self, sender, args):
        """Handle double-click on a value to change its color"""
        try:
            selected_value = self.values_listbox.SelectedItem
            if not selected_value:
                return
                
            # Open color picker dialog
            color_dialog = Forms.ColorDialog()
            color_dialog.Color = Forms.Color.FromArgb(selected_value.r, selected_value.g, selected_value.b)
            
            if color_dialog.ShowDialog() == Forms.DialogResult.OK:
                # Update color
                selected_value.r = color_dialog.Color.R
                selected_value.g = color_dialog.Color.G
                selected_value.b = color_dialog.Color.B
                selected_value.color = Color.FromRgb(color_dialog.Color.R, color_dialog.Color.G, color_dialog.Color.B)
                
                # Refresh listbox
                self.values_listbox.Items.Refresh()
                self.status_text.Text = "Updated color for value: {}".format(selected_value.value)
                self.logger.info("Updated color for value: {}".format(selected_value.value))
        except Exception as ex:
            self.logger.error("Error on value double click: {}".format(ex))

    def on_radio_button_click(self, sender, args):
        """Handle direct clicks on radio buttons"""
        try:
            # Force reload parameters for the selected category
            selected_categories = [item for item in self.category_items if item.IsSelected]
            if selected_categories:
                self.load_parameters_for_category(selected_categories[0].category_info)
        except Exception as ex:
            self.logger.error("Error on radio button click: {}".format(ex))


    def debug_state(self, context=""):
        """Log the current state of critical variables for debugging"""
        try:
            self.logger.info("-------- DEBUG STATE [{}] --------".format(context))
            self.logger.info("DB reference: {}".format(self.DB if hasattr(self, 'DB') else "NOT FOUND"))
            self.logger.info("doc reference: {}".format(self.doc if hasattr(self, 'doc') else "NOT FOUND"))
            self.logger.info("System reference: {}".format(self.System if hasattr(self, 'System') else "NOT FOUND"))
            self.logger.info("values_list count: {}".format(len(self.values_list) if hasattr(self, 'values_list') and self.values_list else "NOT FOUND"))
            self.logger.info("--------------------------------")
        except Exception as ex:
            self.logger.error("Error in debug_state: {}".format(ex))


colorizer_window = RevitColorizerWindow()