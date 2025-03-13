# -*- coding: utf-8 -*-
"""PyRevit Colorizer for elements based on parameter values.
Allows users to colorize elements in Revit views based on parameter values.
"""
# pylint: disable=import-error,unused-argument,missing-docstring,invalid-name,broad-except
# pyright: reportMissingImports=false

import os
import sys
import clr
from random import randint
from unicodedata import normalize
from unicodedata import category as unicode_category
from traceback import extract_tb

# .NET imports
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Data')
clr.AddReference('System.Windows.Forms')
clr.AddReference('WindowsBase')
clr.AddReference('System')

import System
from System import Object
from System.Collections.Generic import Dictionary
from System.Windows import (
    FrameworkElement, Window, Visibility, Controls, 
    Media, Shapes, Input, Data
)
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Collections.Generic import List

# Import Revit API
from pyrevit import HOST_APP, revit, DB, UI
from pyrevit.forms import WPFWindow
from pyrevit.framework import List
from pyrevit.compat import get_elementid_value_func
from pyrevit.script import get_logger
import wpf

# Get logger
logger = get_logger()

# Document references
doc = revit.DOCS.doc
uidoc = HOST_APP.uiapp.ActiveUIDocument
version = int(HOST_APP.version)
uiapp = HOST_APP.uiapp

# Categories to exclude
CAT_EXCLUDED = (int(DB.BuiltInCategory.OST_RoomSeparationLines), 
                int(DB.BuiltInCategory.OST_Cameras), 
                int(DB.BuiltInCategory.OST_CurtainGrids), 
                int(DB.BuiltInCategory.OST_Elev), 
                int(DB.BuiltInCategory.OST_Grids), 
                int(DB.BuiltInCategory.OST_IOSModelGroups), 
                int(DB.BuiltInCategory.OST_Views), 
                int(DB.BuiltInCategory.OST_SitePropertyLineSegment), 
                int(DB.BuiltInCategory.OST_SectionBox), 
                int(DB.BuiltInCategory.OST_ShaftOpening), 
                int(DB.BuiltInCategory.OST_BeamAnalytical), 
                int(DB.BuiltInCategory.OST_StructuralFramingOpening), 
                int(DB.BuiltInCategory.OST_MEPSpaceSeparationLines), 
                int(DB.BuiltInCategory.OST_DuctSystem), 
                int(DB.BuiltInCategory.OST_Lines), 
                int(DB.BuiltInCategory.OST_PipingSystem), 
                int(DB.BuiltInCategory.OST_Matchline), 
                int(DB.BuiltInCategory.OST_CenterLines), 
                int(DB.BuiltInCategory.OST_CurtainGridsRoof), 
                int(DB.BuiltInCategory.OST_SWallRectOpening), 
                -2000278, -1)

# Helper Classes and Methods

    
def get_active_view(active_doc):
    """Get active view from document."""
    selected_view = active_doc.ActiveView
    if selected_view.ViewType == DB.ViewType.ProjectBrowser or selected_view.ViewType == DB.ViewType.SystemBrowser:
        selected_view = active_doc.GetElement(uidoc.GetOpenUIViews()[0].ViewId)
    
    if not selected_view.CanUseTemporaryVisibilityModes():
        UI.TaskDialog.Show(
            "Color Elements by Parameter", 
            "Visibility settings cannot be modified in {} views. Please change your current view.".format(selected_view.ViewType)
        )
        return None
    
    return selected_view


def strip_accents(text):
    """Remove accents from text."""
    return ''.join(char for char in normalize('NFKD', text) if unicode_category(char) != 'Mn')

def solid_fill_pattern_id():
    """Get solid fill pattern ID."""
    solid_fill_id = None
    fillpatterns = DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement)
    for pat in fillpatterns:
        if pat.GetFillPattern().IsSolidFill:
            solid_fill_id = pat.Id
            break
    return solid_fill_id

# Value Classes
class ValuesInfo(object):
    """Class to store parameter value information."""
    def __init__(self, para, val, idt, r, g, b):
        self.par = para
        self.value = val
        self.name = strip_accents(para.Definition.Name)
        self.ele_id = List[DB.ElementId]()
        self.ele_id.Add(idt)
        self.n1 = r
        self.n2 = g
        self.n3 = b
        self.color = Media.Color.FromRgb(r, g, b)
        self.values_double = []
        if para.StorageType == DB.StorageType.Double:
            self.values_double.append(para.AsDouble())
        elif para.StorageType == DB.StorageType.ElementId:
            self.values_double.append(para.AsElementId())

class ParameterInfo(object):
    """Class to store parameter information."""
    def __init__(self, param_type, para):
        self.param_type = param_type  # 0 for instance, 1 for type
        self.rl_par = para
        self.par = para.Definition
        self.name = strip_accents(para.Definition.Name)

    def __str__(self):
        return self.name

class CategoryInfo(object):
    """Class to store category information."""
    def __init__(self, category, parameters):
        self.name = strip_accents(category.Name)
        self.cat = category
        get_elementid_value = get_elementid_value_func()
        self.int_id = get_elementid_value(category.Id)
        self.par = parameters

    def __str__(self):
        return self.name

# Category View Model for XAML binding
class CategoryItem(Object, INotifyPropertyChanged):
    """View model for a category with property change notifications."""
    
    def __init__(self, category_info):
        self._category_info = category_info
        self._is_selected = False
        # Initialize event handlers
        self._propertyChanged = None
        self._window_reference = None
    
    @property
    def name(self):
        return self._category_info.name
    
    @property
    def category_info(self):
        return self._category_info
    
    @property 
    def IsSelected(self):
        return self._is_selected
        
    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            old_value = self._is_selected
            self._is_selected = value
            self.OnPropertyChanged("IsSelected")
            # Notify the window when IsSelected changes (checkbox clicked)
            if self._window_reference:
                self._window_reference.on_category_checkbox_changed(self)
            else:
                old_value, value, self.name
    
    def set_window_reference(self, window):
        """Set reference to the main window for callbacks."""
        self._window_reference = window
    
    # INotifyPropertyChanged implementation
    def add_PropertyChanged(self, handler):
        if self._propertyChanged is None:
            self._propertyChanged = handler
        else:
            self._propertyChanged += handler
        
    def remove_PropertyChanged(self, handler):
        if self._propertyChanged is not None:
            self._propertyChanged -= handler
    
    def OnPropertyChanged(self, property_name):
        if self._propertyChanged is not None:
            self._propertyChanged(self, PropertyChangedEventArgs(property_name))

# Parameter display wrapper for WPF items
class ParameterDisplayItem(object):
    """Wrapper for parameter items to ensure proper display in WPF controls."""
    def __init__(self, parameter_info):
        self.parameter_info = parameter_info
        self.display_name = parameter_info.name
    
    def __str__(self):
        return self.display_name
    
    def ToString(self):
        """Explicit ToString implementation for WPF binding."""
        return self.display_name

# External Event Handlers
class ApplyColorsHandler(UI.IExternalEventHandler):
    """Handler for applying colors to elements."""
    
    def __init__(self, colorizer_ui):
        self.ui = colorizer_ui
    
    def Execute(self, uiapp):
        try:
            active_doc = uiapp.ActiveUIDocument.Document
            # Use the active view from the UI instance
            view = self.ui.active_view
            if not view:
                return
                
            # Use UI instance references to access DB and doc
            DB = self.ui.DB
            doc = self.ui.doc
            revit = self.ui.revit
            logger = self.ui.logger
            
            # Get solid fill pattern ID
            solid_fill_id = None
            fillpatterns = DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement)
            for pat in fillpatterns:
                if pat.GetFillPattern().IsSolidFill:
                    solid_fill_id = pat.Id
                    break
            logger.info("Solid fill ID: %s", solid_fill_id)

            logger.info("Applying colors to elements")
            with revit.Transaction("Apply colors to elements"):
                selected_cat = self.ui.get_selected_category()
                if not selected_cat:
                    logger.info("No category selected")
                    return
                    
                selected_param = self.ui.get_selected_parameter()
                if not selected_param:
                    logger.info("No parameter selected")
                    return
                    
                value_items = self.ui.get_value_items()
                logger.info("Value items: %s", value_items)
                # Check if we're dealing with rooms, spaces, areas that need color schemes
                # Use the stored function from the UI
                get_elementid_value = self.ui.get_elementid_value
                if get_elementid_value(selected_cat.cat.Id) in (
                    int(DB.BuiltInCategory.OST_Rooms), 
                    int(DB.BuiltInCategory.OST_MEPSpaces), 
                    int(DB.BuiltInCategory.OST_Areas)
                ):
                    # Handle rooms/spaces/areas that might need color schemes
                    if self.ui.version > 2021:
                        if str(view.GetColorFillSchemeId(selected_cat.cat.Id)) == "-1":
                            schemes = DB.FilteredElementCollector(active_doc).OfClass(DB.BuiltInCategoryFillScheme).ToElements()
                            for scheme in schemes:
                                if scheme.CategoryId == selected_cat.cat.Id and len(scheme.GetEntries()) > 0:
                                    view.SetColorFillSchemeId(selected_cat.cat.Id, scheme.Id)
                                    break
                    self.ui.statusText.Text = "Note: Rooms, spaces and areas may require a color scheme in the view."
                else:
                    self.ui.statusText.Text = ""
                logger.info("Applying colors to elements")
                # Apply colors to elements
                for value_item in value_items:
                    ogs = DB.OverrideGraphicSettings()
                    color = DB.Color(value_item.n1, value_item.n2, value_item.n3)
                    
                    # Set color properties based on UI settings
                    if self.ui.overrideProjectionCheckbox.IsChecked:
                        ogs.SetProjectionLineColor(color)
                        ogs.SetCutLineColor(color)
                        ogs.SetProjectionLinePatternId(DB.ElementId(-1))
                    
                    # Always set surface pattern color
                    ogs.SetSurfaceForegroundPatternColor(color)
                    ogs.SetCutForegroundPatternColor(color)
                    
                    if solid_fill_id is not None:
                        ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                        ogs.SetCutForegroundPatternId(solid_fill_id)
                    
                    # Apply override to each element
                    for element_id in value_item.ele_id:
                        view.SetElementOverrides(element_id, ogs)
                logger.info("Colors applied successfully to elements.")
                self.ui.statusText.Text = "Colors applied successfully to elements."
                
        except Exception as ex:
            self.ui.statusText.Text = "Error applying colors: " + str(ex)
            logger.error("Error applying colors: %s", ex)
            self.log_exception()
    
    def GetName(self):
        return "Apply Colors to Elements"
    
    def log_exception(self):
        exc_type, exc_value, exc_traceback = sys.exc_info()
        logger = self.ui.logger
        logger.debug("Exception type: %s", exc_type)
        logger.debug("Exception value: %s", exc_value)
        logger.debug("Traceback details:")
        for tb in extract_tb(exc_traceback):
            logger.debug("File: %s, Line: %s, Function: %s, Code: %s", tb[0], tb[1], tb[2], tb[3])

class ResetColorsHandler(UI.IExternalEventHandler):
    """Handler for resetting element colors."""
    
    def __init__(self, colorizer_ui):
        self.ui = colorizer_ui
    
    def Execute(self, uiapp):
        try:
            active_doc = uiapp.ActiveUIDocument.Document
            # Use the active view from the UI instance
            view = self.ui.active_view
            if not view:
                return
            
            # Use UI instance references to access DB, UI, and other modules
            DB = self.ui.DB
            UI = self.ui.UI
            revit = self.ui.revit
            logger = self.ui.logger
                
            with revit.Transaction("Reset element colors"):
                # Reset all element overrides in the view
                ogs = DB.OverrideGraphicSettings()
                
                # Get all elements in view
                collector = DB.FilteredElementCollector(active_doc, view.Id) \
                             .WhereElementIsNotElementType() \
                             .WhereElementIsViewIndependent() \
                             .ToElementIds()
                                
                # Reset element overrides
                for element_id in collector:
                    view.SetElementOverrides(element_id, ogs)
                
                self.ui.statusText.Text = "Colors reset successfully."
                
        except Exception as ex:
            self.ui.statusText.Text = "Error resetting colors: " + str(ex)
            logger.error("Error resetting colors: %s", ex)
            self.log_exception()
    
    def GetName(self):
        return "Reset Element Colors"
    
    def log_exception(self):
        exc_type, exc_value, exc_traceback = sys.exc_info()
        logger = self.ui.logger
        logger.debug("Exception type: %s", exc_type)
        logger.debug("Exception value: %s", exc_value)
        logger.debug("Traceback details:")
        for tb in extract_tb(exc_traceback):
            logger.debug("File: %s, Line: %s, Function: %s, Code: %s", tb[0], tb[1], tb[2], tb[3])

def get_used_categories(active_view, excluded_cats=None):
    """Get all used categories and their parameters in the active view."""
    if excluded_cats is None:
        excluded_cats = CAT_EXCLUDED
        
    # Get all elements in view
    collector = DB.FilteredElementCollector(doc, active_view.Id) \
                 .WhereElementIsNotElementType() \
                 .WhereElementIsViewIndependent() \
                 .ToElements()
                 
    categories = []
    
    # Get the function once instead of calling it for each element
    get_elementid_value = get_elementid_value_func()
    
    for element in collector:
        if element.Category is None:
            continue
            
        current_cat_id = get_elementid_value(element.Category.Id)
        
        # Skip excluded categories and already processed categories
        if current_cat_id in excluded_cats or any(x.int_id == current_cat_id for x in categories):
            continue
            
        # Get instance parameters
        instance_parameters = []
        for param in element.Parameters:
            if param.Definition.BuiltInParameter not in (DB.BuiltInParameter.ELEM_CATEGORY_PARAM, 
                                                        DB.BuiltInParameter.ELEM_CATEGORY_PARAM_MT):
                instance_parameters.append(ParameterInfo(0, param))
        
        # Get type parameters
        type_element = element.Document.GetElement(element.GetTypeId())
        if type_element is None:
            continue
            
        type_parameters = []
        for param in type_element.Parameters:
            if param.Definition.BuiltInParameter not in (DB.BuiltInParameter.ELEM_CATEGORY_PARAM, 
                                                        DB.BuiltInParameter.ELEM_CATEGORY_PARAM_MT):
                type_parameters.append(ParameterInfo(1, param))
        
        # Combine all parameters
        all_parameters = instance_parameters + type_parameters
        all_parameters.sort(key=lambda x: x.name.upper())
        
        # Add category to the list
        categories.append(CategoryInfo(element.Category, all_parameters))
    
    # Sort categories by name
    categories.sort(key=lambda x: x.name)
    return categories

# Main UI Class
class RevitColorizerWindow(WPFWindow):
    """Main WPF window for the Revit Colorizer tool."""
    
    def __init__(self):
        try:
            # Add initialization tracking
            logger.info("=============== WINDOW INITIALIZATION STARTED ===============")
            
            # Create XAML file path
            xaml_file = os.path.join(
                os.path.dirname(__file__), 
                "RevitColorizerWindow.xaml"
            )
            
            # Load XAML
            WPFWindow.__init__(self, xaml_file)
            
            # Store references
            self.logger = logger
            self.DB = DB
            self.doc = doc
            self.uidoc = uidoc
            self.uiapp = uiapp  # Store reference to uiapp for closing handler
            self.System = System
            self.randint = randint
            self.version = version
            self.Media = Media  # Store Media reference
            self.ParameterDisplayItem = ParameterDisplayItem  # Store reference to the wrapper class
            self.revit = revit  # Store reference to the revit module
            self.UI = UI  # Store reference to the UI module
            
            # Store both the function and its result for use in different contexts
            self.get_elementid_value_func_ref = get_elementid_value_func  # Store the function reference
            self.get_elementid_value = get_elementid_value_func()  # Store the result of calling the function
            
            # Create a custom ValuesInfo class
            self.ValuesInfo = self.create_custom_values_info_class()
            
            # Create external event handlers
            self.apply_colors_handler = ApplyColorsHandler(self)
            self.apply_colors_event = UI.ExternalEvent.Create(self.apply_colors_handler)
            
            self.reset_colors_handler = ResetColorsHandler(self)
            self.reset_colors_event = UI.ExternalEvent.Create(self.reset_colors_handler)
            
            # Get active view
            self.active_view = get_active_view(doc)
            if not self.active_view:
                self.Close()
                return
                
            # Set up UI event handlers
            self.setup_ui_components()
            
            # Load data
            self.load_categories()
            
            # Show startup message
            self.statusText.Text = "Ready. Select a category to begin."
            
            logger.info("=============== WINDOW INITIALIZATION COMPLETED ===============")
            
        except Exception as ex:
            UI.TaskDialog.Show("Error", "Failed to initialize Revit Colorizer: " + str(ex))
            logger.error("Initialization error: %s", str(ex))
            self.Close()
    
    def create_custom_values_info_class(self):
        """Create a completely custom version of ValuesInfo that doesn't use strip_accents."""
        outer_self = self
        # Store reference to Media namespace for the inner class
        Media_ref = self.Media
        
        class CustomValuesInfo(object):
            """Simplified class to store parameter value information without accent stripping."""
            def __init__(self, para, val, idt, r, g, b):
                self.par = para
                self.value = val
                # Just use the parameter name directly without accent stripping
                self.name = para.Definition.Name
                # Use a simple Python list instead of System.Collections.Generic.List
                self.ele_id = []
                self.ele_id.append(idt)
                self.n1 = r
                self.n2 = g
                self.n3 = b
                # Explicitly create a proper WPF color object for binding
                self._color = Media_ref.Color.FromRgb(r, g, b)
                self._brush = Media_ref.SolidColorBrush(self._color)
                self.values_double = []
                if para.StorageType == outer_self.DB.StorageType.Double:
                    self.values_double.append(para.AsDouble())
                elif para.StorageType == outer_self.DB.StorageType.ElementId:
                    self.values_double.append(para.AsElementId())
                
                # Add a method that mimics the Add method of List
                self.ele_id_add = self.ele_id.append
            
            # Color property with proper getter
            @property
            def color(self):
                return self._brush
            
            # Setter to update color
            @color.setter
            def color(self, wpf_color):
                self._color = wpf_color
                self._brush = Media_ref.SolidColorBrush(wpf_color)
        
        return CustomValuesInfo
    
    def setup_ui_components(self):
        """Set up UI components and event handlers."""
        # Close button
        self.closeButton.Click += self.on_close
        
        # Apply and Reset buttons
        self.applyButton.Click += self.on_apply_colors
        self.resetButton.Click += self.on_reset_colors
        
        # Filter buttons
        self.addFiltersButton.Click += self.on_add_filters
        self.removeFiltersButton.Click += self.on_remove_filters
        
        # Instead of relying on ListBox selection, we'll manually check the 
        # CheckBox controls after the form loads and ListBox is populated
        
        # Parameter type selection
        self.instanceRadioButton.Checked += self.on_parameter_type_changed
        self.typeRadioButton.Checked += self.on_parameter_type_changed
        
        # Parameter selection
        self.parameterSelector.SelectionChanged += self.on_parameter_selected
        self.parameterSelector.DropDownClosed += self.on_parameter_dropdown_closed

        # Refresh parameters button
        self.refreshParametersButton.Click += self.on_refresh_parameters
        
        # Values list
        self.valuesListBox.MouseDoubleClick += self.on_value_double_click
        self.valuesListBox.SelectionChanged += self.on_value_click

        # Checkbox events
        self.showElementsCheckbox.Checked += self.on_show_elements_changed
        self.showElementsCheckbox.Unchecked += self.on_show_elements_changed
        
        # Setting initial values for checkboxes
        self.overrideProjectionCheckbox.IsChecked = True
        self.showElementsCheckbox.IsChecked = False
        
        # Default to instance parameters being selected
        self.instanceRadioButton.IsChecked = True
        
        # Add a handler for mouse click events on the ListBox to detect checkbox clicks
        self.categoryListBox.PreviewMouseLeftButtonUp += self.on_category_listbox_clicked
        
        # Add window closing event handler to reset colors when window closes
        self.Closing += self.on_window_closing

    def load_categories(self):
        """Load categories from the current view."""
        try:
            # Clear existing items
            self.categoryListBox.Items.Clear()
            
            # Get categories used in the current view
            categories = get_used_categories(self.active_view, CAT_EXCLUDED)
            
            # Create category items for the list
            for category in categories:
                category_item = CategoryItem(category)
                # Set reference to this window BEFORE adding to listbox
                category_item.set_window_reference(self)
                self.categoryListBox.Items.Add(category_item)
                
            # Process UI updates
            self.categoryListBox.UpdateLayout()
            
            # Select the first category by default and directly call the handler to load parameters
            if self.categoryListBox.Items.Count > 0:
                # Important: Set the first category as selected explicitly
                
                # Use a slight delay to ensure the UI is ready
                def select_first_category():
                    self.categoryListBox.Items[0].IsSelected = True
                    self.process_category_selection()
                
                # Use Dispatcher to ensure UI is updated properly
                self.System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                    self.System.Action(select_first_category),
                    self.System.Windows.Threading.DispatcherPriority.Background)
                
            self.statusText.Text = "Loaded {} categories.".format(len(categories))
            self.logger.info("Loaded {} categories with checkbox bindings.".format(len(categories)))
                
        except Exception as ex:
            self.statusText.Text = "Error loading categories: " + str(ex)
            self.logger.error("Error loading categories: %s", str(ex))
    
    def on_category_listbox_clicked(self, sender, args):
        """Handle mouse clicks on the category list box to detect checkbox clicks."""
        try:
            def delayed_process():
                self.process_category_selection()

            # Use a short delay to ensure the checkbox state has been updated
            self.System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                self.System.Action(delayed_process),
                self.System.Windows.Threading.DispatcherPriority.Background)
            
        except Exception as ex:
            self.statusText.Text = "Error handling category click: " + str(ex)
            self.logger.error("Category click error: %s", str(ex))
    
    def get_selected_categories(self):
        """Get all selected categories."""
        selected_categories = []
        for item in self.categoryListBox.Items:
            if item.IsSelected:
                selected_categories.append(item.category_info)
        return selected_categories
    
    def process_category_selection(self):
        """Process the current category selection state."""
        try:
            # Store currently selected parameter name to try to preserve it
            current_param_name = None
            if self.parameterSelector.SelectedItem:
                current_param_name = self.parameterSelector.SelectedItem.parameter_info.name
            
            # Get current parameter type (instance/type)
            current_param_type_is_instance = self.instanceRadioButton.IsChecked
            
            # Load parameters for all selected categories while preserving selection
            self.load_parameters_for_categories(current_param_name, current_param_type_is_instance)
            
            # Count the number of selected categories
            selected_count = sum(1 for item in self.categoryListBox.Items if item.IsSelected)
            
            if selected_count == 0:
                self.statusText.Text = "No categories selected."
                self.valuesListBox.Items.Clear()
            
            # Force an update to the UI
            self.categoryListBox.UpdateLayout()
            
        except Exception as ex:
            self.statusText.Text = "Error processing category selection: " + str(ex)
            self.logger.error("Process category selection error: %s", str(ex))
    
    def load_parameters_for_categories(self, preserve_param_name=None, preserve_param_type_is_instance=None):
        """Load common parameters for all selected categories.
        
        Args:
            preserve_param_name: Name of parameter to preserve selection for
            preserve_param_type_is_instance: Whether the parameter to preserve is an instance parameter
        """
        try:
            # Clear parameter list
            self.parameterSelector.Items.Clear()
            
            # Get all selected categories
            selected_categories = self.get_selected_categories()
            
            if not selected_categories:
                self.statusText.Text = "No categories selected."
                self.valuesListBox.Items.Clear()
                return
            
            # If parameter type was not specified to preserve, use current selection
            if preserve_param_type_is_instance is None:
                is_instance = self.instanceRadioButton.IsChecked
            else:
                # Use the specified parameter type and set the radio buttons
                is_instance = preserve_param_type_is_instance
                self.instanceRadioButton.IsChecked = is_instance
                self.typeRadioButton.IsChecked = not is_instance
            
            param_type_code = 0 if is_instance else 1
            
            # Start with parameters from the first category
            first_category = selected_categories[0]
            common_params = [p for p in first_category.par if p.param_type == param_type_code]
            
            # Find common parameters across all selected categories
            for category in selected_categories[1:]:
                # Get filtered parameters for this category
                category_params = [p for p in category.par if p.param_type == param_type_code]
                
                # Filter down to parameters that exist in both lists (by name)
                category_param_names = [p.name for p in category_params]
                common_params = [p for p in common_params if p.name in category_param_names]
            
            # Add common parameters to dropdown with display wrapper
            selected_index = -1
            for i, param in enumerate(common_params):
                self.parameterSelector.Items.Add(self.ParameterDisplayItem(param))
                # Check if this is the parameter to preserve
                if preserve_param_name and param.name == preserve_param_name:
                    selected_index = i
            
            # Restore selection if possible, otherwise select first parameter
            if selected_index >= 0:
                self.parameterSelector.SelectedIndex = selected_index
                self.statusText.Text = "Loaded {} common parameters, maintained selection of '{}'.".format(
                    len(common_params), preserve_param_name)
            elif self.parameterSelector.Items.Count > 0:
                self.parameterSelector.SelectedIndex = 0
                param_type_name = "instance" if is_instance else "type"
                if preserve_param_name:
                    self.statusText.Text = "Loaded {} common {} parameters. Parameter '{}' is not common to selected categories.".format(
                        len(common_params), param_type_name, preserve_param_name)
                else:
                    self.statusText.Text = "Loaded {} common {} parameters for {} selected categories.".format(
                        len(common_params), param_type_name, len(selected_categories))
            else:
                self.statusText.Text = "No common {} parameters found for the selected categories.".format(
                    "instance" if is_instance else "type")
                self.valuesListBox.Items.Clear()
                
        except Exception as ex:
            self.statusText.Text = "Error loading parameters: " + str(ex)
            self.logger.error("Parameter loading error: %s", str(ex))
    
    def on_parameter_type_changed(self, sender, args):
        """Handle parameter type (instance/type) changed event."""
        try:
            # Reload parameters for all selected categories with the new parameter type
            self.load_parameters_for_categories()
        except Exception as ex:
            self.statusText.Text = "Error changing parameter type: " + str(ex)
            self.logger.error("Parameter type change error: %s", str(ex))
    
    def on_parameter_selected(self, sender, args):
        """Handle parameter selection changed event."""
        try:
            selected_parameter = self.get_selected_parameter()
            
            if selected_parameter:
                # Load parameter values for all selected categories
                values = self.get_parameter_values(selected_parameter, self.active_view)
                
                # Clear and populate values list
                self.valuesListBox.Items.Clear()
                
                for value_info in values:
                    self.valuesListBox.Items.Add(value_info)
                    
                self.statusText.Text = "Loaded {} values for parameter '{}'.".format(
                    len(values), selected_parameter.name)
            else:
                self.valuesListBox.Items.Clear()
                self.statusText.Text = "No parameter selected."
        except Exception as ex:
            self.statusText.Text = "Error selecting parameter: " + str(ex)
            self.logger.error("Parameter selection error: %s", str(ex))
    
    def on_parameter_dropdown_closed(self, sender, args):
        """Handle parameter dropdown closed event."""
        try:
            self.on_parameter_selected(sender, args)
        except Exception as ex:
            self.statusText.Text = "Error processing parameter selection: " + str(ex)
            self.logger.error("Parameter dropdown closed error: %s", str(ex))
    
    def on_value_click(self, sender, args):
        """Handle value item click to select Revit elements based on value."""
        try:
            # Check if we have a selected value
            if sender.SelectedItem is None or not self.showElementsCheckbox.IsChecked:
                return
                
            value_item = sender.SelectedItem
            
            # Create element ID collection for selection
            element_ids = self.System.Collections.Generic.List[self.DB.ElementId]()
            
            # Add all element IDs for this value to the collection
            for element_id in value_item.ele_id:
                element_ids.Add(element_id)
            # Set selection in the active document
            self.uidoc.Selection.SetElementIds(element_ids)
            
            # Update status
            elements_count = len(value_item.ele_id)
            self.statusText.Text = "Selected {} elements with value '{}'.".format(
                elements_count, value_item.value)
            
        except Exception as ex:
            self.logger.error("Element selection error: %s", str(ex))
    
    def on_value_double_click(self, sender, args):
        """Handle value item double click to change color."""
        try:
            if sender.SelectedItem is None:
                return
                
            value_item = sender.SelectedItem
            
            # Create color picker dialog
            color_dialog = self.System.Windows.Forms.ColorDialog()
            color_dialog.Color = self.System.Drawing.Color.FromArgb(value_item.n1, value_item.n2, value_item.n3)
            color_dialog.AllowFullOpen = True
            
            # Show dialog and get result
            if color_dialog.ShowDialog() == self.System.Windows.Forms.DialogResult.OK:
                # Update color values
                value_item.n1 = color_dialog.Color.R
                value_item.n2 = color_dialog.Color.G
                value_item.n3 = color_dialog.Color.B
                # Use the property setter to update the brush
                value_item.color = self.Media.Color.FromRgb(
                    color_dialog.Color.R, color_dialog.Color.G, color_dialog.Color.B)
                
                # Refresh the list view
                self.valuesListBox.Items.Refresh()
                self.statusText.Text = "Color updated for value '{}'.".format(value_item.value)
        except Exception as ex:
            self.statusText.Text = "Error changing color: " + str(ex)
            self.logger.error("Value double click error: %s", str(ex))
    
    def on_apply_colors(self, sender, args):
        """Handle apply colors button click."""
        try:
            if not self.valuesListBox.Items.Count:
                self.statusText.Text = "No values to apply colors to."
                return
                
            self.statusText.Text = "Applying colors..."
            self.apply_colors_event.Raise()
        except Exception as ex:
            self.statusText.Text = "Error applying colors: " + str(ex)
            self.logger.error("Apply colors error: %s", str(ex))
    
    def on_reset_colors(self, sender, args):
        """Handle reset colors button click."""
        try:
            self.statusText.Text = "Resetting colors..."
            self.reset_colors_event.Raise()
        except Exception as ex:
            self.statusText.Text = "Error resetting colors: " + str(ex)
            self.logger.error("Reset colors error: %s", str(ex))
    
    def on_add_filters(self, sender, args):
        """Handle add filters button click."""
        try:
            self.statusText.Text = "Creating view filters is not implemented yet."
            # This would be implemented similar to the CreateFilters class in ColorSplasher
        except Exception as ex:
            self.statusText.Text = "Error adding filters: " + str(ex)
            self.logger.error("Add filters error: %s", str(ex))
    
    def on_remove_filters(self, sender, args):
        """Handle remove filters button click."""
        try:
            self.statusText.Text = "Removing filters is not implemented yet."
            # This would be implemented similar to the reset colors but focused on filters
        except Exception as ex:
            self.statusText.Text = "Error removing filters: " + str(ex)
            self.logger.error("Remove filters error: %s", str(ex))
    
    def on_save_config(self, sender, args):
        """Handle save configuration button click."""
        try:
            self.statusText.Text = "Saving configuration is not implemented yet."
            # This would save the current configuration to a file
        except Exception as ex:
            self.statusText.Text = "Error saving configuration: " + str(ex)
            self.logger.error("Save config error: %s", str(ex))
    
    def on_load_config(self, sender, args):
        """Handle load configuration button click."""
        try:
            self.statusText.Text = "Loading configuration is not implemented yet."
            # This would load a configuration from a file
        except Exception as ex:
            self.statusText.Text = "Error loading configuration: " + str(ex)
            self.logger.error("Load config error: %s", str(ex))
    
    def on_close(self, sender, args):
        """Handle close button click."""
        try:
            self.Close()
        except Exception as ex:
            self.logger.error("Close error: %s", str(ex))
            self.Close()
            
    def on_window_closing(self, sender, e):
        """Handle window closing event to reset colors when window closes."""
        try:
            self.logger.info("Window closing event triggered - resetting element colors")
            # Only reset colors if we actually have a valid view
            if self.active_view:
                # Reset colors by raising the reset colors event
                # Note: we need to raise the event because we want the reset to happen
                # in the Revit main thread, not in the UI thread.
                self.reset_colors_event.Raise()
                self.logger.info("Reset colors event raised on window close")
        except Exception as ex:
            # Just log the error, don't show to user since window is closing
            self.logger.error("Error resetting colors on window close: %s", str(ex))
    
    def OnClosing(self, e):
        """Override WPF window's OnClosing method to ensure colors are reset when window closes.
        This catches ALL closing scenarios including the Windows X button."""
        try:
            self.logger.info("OnClosing override triggered - resetting element colors")
            # Only reset colors if we actually have a valid view
            if self.active_view:
                # Raise the external event instead of directly executing the handler
                # This ensures the handler runs in the proper Revit API context
                self.reset_colors_event.Raise()
                self.logger.info("Reset colors event raised in OnClosing")
        except Exception as ex:
            # Just log the error, don't show to user since window is closing
            self.logger.error("Error in OnClosing: %s", str(ex))
        
        # Always call the base class implementation
        super(RevitColorizerWindow, self).OnClosing(e)
    
    # Helper methods to get selected items
    def get_selected_category(self):
        """Get the first checked category."""
        # Check all category items and return the first one that is checked
        for item in self.categoryListBox.Items:
            if item.IsSelected:
                return item.category_info
        return None
    
    
    def get_parameter_values(self, param, view):
        """Get all values for a parameter across all selected categories."""
        values = []
        used_colors = set()
        
        # Get all selected categories
        selected_categories = self.get_selected_categories()
        
        for category in selected_categories:
            # Try to find BuiltInCategory if possible
            bic = None
            for sample_bic in self.System.Enum.GetValues(self.DB.BuiltInCategory):
                if category.int_id == int(sample_bic):
                    bic = sample_bic
                    break
            
            if not bic:
                continue
                
            # Get all elements of this category in view
            collector = self.DB.FilteredElementCollector(self.doc, view.Id) \
                        .OfCategory(bic) \
                        .WhereElementIsNotElementType() \
                        .WhereElementIsViewIndependent() \
                        .ToElements()
                        
            for element in collector:
                # Get parameter host (instance or type)
                param_host = element if param.param_type == 0 else self.doc.GetElement(element.GetTypeId())
                
                if not param_host:
                    continue
                
                # Try to find the parameter
                for parameter in param_host.Parameters:
                    if parameter.Definition.Name == param.par.Name:
                        value = self.get_parameter_value(parameter)
                        
                        # Check if this value already exists
                        matching_items = [x for x in values if x.value == value]
                        
                        if matching_items:
                            # Add element ID to existing value
                            matching_items[0].ele_id.append(element.Id)
                            if parameter.StorageType == self.DB.StorageType.Double:
                                matching_items[0].values_double.append(parameter.AsDouble())
                        else:
                            # Instead of picking a random color, create a placeholder for now
                            # We'll set proper colors after collecting all unique values
                            values.append(self.ValuesInfo(parameter, value, element.Id, 0, 0, 0))
                        break
        
        # Sort values, with "None" at the end
        none_values = [x for x in values if x.value == "None"]
        values = [x for x in values if x.value != "None"]
        
        # Try to sort numerically if possible
        if values:
            try:
                # Extract numeric value from strings like "136 m²"
                def extract_number(value):
                    # Remove any non-numeric characters except decimal point and negative sign
                    num_str = ''.join(c for c in value if c.isdigit() or c in '.-')
                    try:
                        return float(num_str) if num_str else float('inf')
                    except ValueError:
                        return float('inf')
                
                values.sort(key=lambda x: extract_number(x.value))
            except:
                # Fall back to string sorting if numeric sort fails
                values.sort(key=lambda x: x.value)
        
        # Add None values at the end
        if none_values and any(len(x.ele_id) > 0 for x in none_values):
            values.extend(none_values)
        
        # Now assign colors from our color range
        if values:
            # Generate color range for the number of values we have
            color_range = self.generate_color_range(len(values))
            
            # Assign colors to each value
            for i, value_item in enumerate(values):
                r, g, b = color_range[i]
                value_item.n1 = r
                value_item.n2 = g
                value_item.n3 = b
                # Update the color property with a new WPF color object
                value_item.color = self.Media.Color.FromRgb(r, g, b)
                
        # Always use gray for None values
        for value_item in none_values:
            value_item.n1 = 192
            value_item.n2 = 192
            value_item.n3 = 192
            # Update the color property with a new WPF color object
            value_item.color = self.Media.Color.FromRgb(192, 192, 192)
                
        return values


    def get_selected_parameter(self):
        """Get the selected parameter."""
        if self.parameterSelector.SelectedItem:
            # Unwrap the parameter from its display wrapper
            return self.parameterSelector.SelectedItem.parameter_info
        return None
    
    def get_value_items(self):
        """Get all value items from the list."""
        return [item for item in self.valuesListBox.Items]
    

    def random_color(self):
        """Generate a random color."""
        r = self.randint(0, 230)
        g = self.randint(0, 230)
        b = self.randint(0, 230)
        return r, g, b
    
    def generate_color_range(self, count):
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
    
    def get_parameter_value(self, para):
        """Get parameter value as string."""
        if not para.HasValue:
            return "None"
        
        if para.StorageType == self.DB.StorageType.Double:
            return para.AsValueString()
        elif para.StorageType == self.DB.StorageType.ElementId:
            id_val = para.AsElementId()
            if self.get_elementid_value(id_val) >= 0:
                return self.DB.Element.Name.GetValue(self.doc.GetElement(id_val))
            else:
                return "None"
        elif para.StorageType == self.DB.StorageType.Integer:
            if self.version > 2021:
                param_type = para.Definition.GetDataType()
                if self.DB.SpecTypeId.Boolean.YesNo == param_type:
                    return "True" if para.AsInteger() == 1 else "False"
                else:
                    return para.AsValueString()
            else:
                param_type = para.Definition.ParameterType
                if self.DB.ParameterType.YesNo == param_type:
                    return "True" if para.AsInteger() == 1 else "False"
                else:
                    return para.AsValueString()
        elif para.StorageType == self.DB.StorageType.String:
            return para.AsString()
        else:
            return "None"

    def on_show_elements_changed(self, sender, args):
        """Handle show elements checkbox state change."""
        try:
            # If checkbox is unchecked, clear Revit selection
            if not self.showElementsCheckbox.IsChecked:
                # Clear current selection
                self.uidoc.Selection.SetElementIds(self.System.Collections.Generic.List[self.DB.ElementId]())
                self.statusText.Text = "Element selection disabled."
            else:
                # If a value is selected, apply selection immediately
                if self.valuesListBox.SelectedItem is not None:
                    self.on_value_click(self.valuesListBox, None)
                else:
                    self.statusText.Text = "Element selection enabled. Click a value to select elements."
        except Exception as ex:
            self.statusText.Text = "Error changing selection mode: " + str(ex)
            self.logger.error("Show elements checkbox error: %s", str(ex))

    def on_category_checkbox_changed(self, category_item):
        """Handle category checkbox selection changed event."""
        try:
            self.logger.info("Category checkbox changed: %s, IsSelected: %s", 
                             category_item.name, category_item.IsSelected)
            
            # Directly process the selection
            self.process_category_selection()
                
        except Exception as ex:
            self.statusText.Text = "Error handling category checkbox: " + str(ex)
            self.logger.error("Category checkbox error: %s", str(ex))

    def on_refresh_parameters(self, sender, args):
        """Handle refresh parameters button click."""
        try:
            # Store current selections
            current_param_type_is_instance = self.instanceRadioButton.IsChecked
            current_param_name = None
            if self.parameterSelector.SelectedItem:
                current_param_name = self.parameterSelector.SelectedItem.parameter_info.name
                
            # Clear and reload parameters
            self.parameterSelector.Items.Clear()
            selected_categories = self.get_selected_categories()
            
            if not selected_categories:
                self.statusText.Text = "No categories selected."
                self.valuesListBox.Items.Clear()
                return
                
            # Get common parameters
            param_type_code = 0 if current_param_type_is_instance else 1
            common_params = [p for p in selected_categories[0].par if p.param_type == param_type_code]
            
            for category in selected_categories[1:]:
                category_params = [p for p in category.par if p.param_type == param_type_code]
                category_param_names = [p.name for p in category_params]
                common_params = [p for p in common_params if p.name in category_param_names]
                
            # Add parameters to dropdown
            selected_index = -1
            for i, param in enumerate(common_params):
                self.parameterSelector.Items.Add(self.ParameterDisplayItem(param))
                if current_param_name and param.name == current_param_name:
                    selected_index = i
                    
            # Restore selection
            if selected_index >= 0:
                self.parameterSelector.SelectedIndex = selected_index
                self.statusText.Text = "Refreshed parameters, maintained selection of '{0}'.".format(current_param_name)
            elif self.parameterSelector.Items.Count > 0:
                self.parameterSelector.SelectedIndex = 0
                self.statusText.Text = "Refreshed {0} common parameters.".format(len(common_params))
            else:
                param_type = "instance" if current_param_type_is_instance else "type"
                self.statusText.Text = "No common {0} parameters found.".format(param_type)
                self.valuesListBox.Items.Clear()
                return
                
            # Reload values
            if self.parameterSelector.SelectedItem:
                self.on_parameter_selected(self.parameterSelector, None)
                
        except Exception as ex:
            self.statusText.Text = "Error refreshing parameters: {0}".format(str(ex))
            self.logger.error("Parameter refresh error: %s", str(ex))

# Main script execution
if __name__ == "__main__":
    # Check if we have an active view
    active_view = get_active_view(doc)
    
    if active_view:
        # Start the UI
        colorizer_ui = RevitColorizerWindow()
        colorizer_ui.Show()
    else:
        UI.TaskDialog.Show(
            "PyRevit Colorizer", 
            "Please open a view where visibility settings can be modified."
        )