# -*- coding: utf-8 -*-
__title__ = "Color\nElements"
__author__ = "Byggstyrning AB"
__doc__ = """Color Elements and selected elements by Parameter values.
Allows you to colorize elements in Revit views based on parameter values.
"""
__highlight__ = 'updated'
# pylint: disable=import-error,unused-argument,missing-docstring,invalid-name,broad-except
# pyright: reportMissingImports=false

import os
import sys
import clr
from random import randint
from unicodedata import normalize
from unicodedata import category as unicode_category
from traceback import extract_tb
import re

# Create coloringschemas directory if it doesn't exist
script_path = __file__
panel_dir = os.path.dirname(script_path)
tab_dir = os.path.dirname(panel_dir)
extension_dir = os.path.dirname(os.path.dirname(tab_dir))
script_dir = extension_dir
schemas_dir = os.path.join(script_dir, "coloringschemas")
if not os.path.exists(schemas_dir):
    os.makedirs(schemas_dir)

# Add lib path to sys.path for importing lib modules
lib_path = os.path.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# .NET imports
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Data')
clr.AddReference('System.Windows.Forms')
clr.AddReference('WindowsBase')
clr.AddReference('System')

import System
from System import Object
from System.Windows.Threading import DispatcherPriority
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
from pyrevit.script import get_logger
import wpf
from revit import revit_utils

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
        # Add IsChecked property for checkbox binding
        self._is_checked = False
        
        # Add a method that mimics the Add method of List
        self.ele_id_add = self.ele_id.append
    
    # IsChecked property with getter and setter
    @property
    def IsChecked(self):
        return self._is_checked
    
    @IsChecked.setter
    def IsChecked(self, value):
        self._is_checked = value
    
    # Color property with proper getter
    @property
    def color(self):
        return self.color

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
        get_elementid_value = revit_utils.get_elementid_value_func()
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
            
            # Get all views to apply colors to (multi-view support)
            views_to_color = self.ui.get_target_views()
            if not views_to_color:
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

            with revit.Transaction("Apply colors to elements"):
                selected_cat = self.ui.get_selected_category()
                if not selected_cat:
                    return
                    
                selected_param = self.ui.get_selected_parameter()
                if not selected_param:
                    return
                    
                value_items = self.ui.get_value_items()
                # Check if we're dealing with rooms, spaces, areas that need color schemes
                # Use the stored function from the UI
                get_elementid_value = self.ui.get_elementid_value
                is_room_space_area = get_elementid_value(selected_cat.cat.Id) in (
                    int(DB.BuiltInCategory.OST_Rooms), 
                    int(DB.BuiltInCategory.OST_MEPSpaces), 
                    int(DB.BuiltInCategory.OST_Areas)
                )
                
                # Apply colors to each selected view
                for view in views_to_color:
                    try:
                        if is_room_space_area:
                            # Handle rooms/spaces/areas that might need color schemes
                            if self.ui.version > 2021:
                                if str(view.GetColorFillSchemeId(selected_cat.cat.Id)) == "-1":
                                    schemes = DB.FilteredElementCollector(active_doc).OfClass(DB.BuiltInCategoryFillScheme).ToElements()
                                    for scheme in schemes:
                                        if scheme.CategoryId == selected_cat.cat.Id and len(scheme.GetEntries()) > 0:
                                            view.SetColorFillSchemeId(selected_cat.cat.Id, scheme.Id)
                                            break
                            
                        # Apply colors to elements in this view
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
                            
                            # Apply override to each element in this view
                            for element_id in value_item.ele_id:
                                view.SetElementOverrides(element_id, ogs)
                                
                    except Exception as view_ex:
                        logger.debug("Error applying colors to view {}: {}".format(view.Name, str(view_ex)))

                # Update status text
                view_count = len(views_to_color)
                if is_room_space_area:
                    self.ui.statusText.Text = "Colors applied to {} view(s). Note: Rooms/spaces/areas may require color schemes.".format(view_count)
                else:
                    self.ui.statusText.Text = "Colors applied successfully to {} view(s).".format(view_count)
                
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
        self.specific_view = None
        self.specific_views = None  # For resetting multiple views at once
    
    def Execute(self, uiapp):
        try:
            # Check if window is still active
            if not self.ui.is_window_active:
                return
            
            # Don't reset colors during initialization
            if getattr(self.ui, '_is_initializing', False):
                return
                
            active_doc = uiapp.ActiveUIDocument.Document
            
            # Determine which views to reset
            views_to_reset = []
            
            if self.specific_views:
                # Multiple specific views provided (e.g., when closing window)
                views_to_reset = self.specific_views
                self.specific_views = None
            elif self.specific_view:
                # Single specific view provided (e.g., when changing views)
                views_to_reset = [self.specific_view]
                self.specific_view = None
            else:
                # Use all selected views (multi-view mode) or active view
                views_to_reset = self.ui.get_target_views()
            
            if not views_to_reset:
                return
            
            # Use UI instance references to access DB, UI, and other modules
            DB = self.ui.DB
            UI = self.ui.UI
            revit = self.ui.revit
            logger = self.ui.logger
                
            with revit.Transaction("Reset element colors"):
                # Reset all element overrides in each view
                ogs = DB.OverrideGraphicSettings()
                
                for view in views_to_reset:
                    try:
                        # Get all elements in this view
                        collector = DB.FilteredElementCollector(active_doc, view.Id) \
                                     .WhereElementIsNotElementType() \
                                     .WhereElementIsViewIndependent() \
                                     .ToElementIds()
                                        
                        # Reset element overrides
                        for element_id in collector:
                            view.SetElementOverrides(element_id, ogs)
                            
                    except Exception as view_ex:
                        logger.debug("Error resetting colors in view {}: {}".format(view.Name, str(view_ex)))
                
                view_count = len(views_to_reset)
                self.ui.statusText.Text = "Colors reset in {} view(s).".format(view_count)
                
        except Exception as ex:
            self.ui.statusText.Text = "Error resetting colors: " + str(ex)
            logger.error("Error resetting colors: %s", ex)
            self.log_exception()
    
    def set_specific_view(self, view):
        """Set a specific view to reset colors in."""
        self.specific_view = view
        self.specific_views = None
    
    def set_specific_views(self, views):
        """Set multiple specific views to reset colors in."""
        self.specific_views = views
        self.specific_view = None
        
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



# Main UI Class
class RevitColorizerWindow(WPFWindow):
    """Main WPF window for the Revit Colorizer tool."""
    
    def __init__(self):
        # Initialize active_view to None to prevent AttributeError in OnClosing
        self.active_view = None
        try:
# Create XAML file path
            xaml_file = os.path.join(
                os.path.dirname(__file__), 
                "ColorElementsWindow.xaml"
            )
            
            
            # Load XAML
            try:
                WPFWindow.__init__(self, xaml_file)
                
                # Load styles AFTER window initialization (window-scoped, does not affect Revit UI)
                from styles import load_styles_to_window
                load_styles_to_window(self)
# Create dummy statusText object since status bar was removed from XAML
                # This prevents errors when code tries to set statusText.Text
                class DummyStatusText(object):
                    def __init__(self):
                        self._text = ""
                    @property
                    def Text(self):
                        return self._text
                    @Text.setter
                    def Text(self, value):
                        # Do nothing - status bar is hidden
                        self._text = value
                
                # Only create dummy if statusText doesn't exist (it was removed from XAML)
                if not hasattr(self, 'statusText'):
                    self.statusText = DummyStatusText()
            except Exception as ex:
                raise
            
            # Load styles ResourceDictionary (for window-specific resources if needed)
            self.load_styles()
            
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
            
            # Flag indicating if the window is active
            self.is_window_active = True
            
            # Flag to prevent color resets during initialization
            self._is_initializing = True
            
            # Flag to prevent selection storage during restore process
            self._processing_restored_selection = False
            
            # Flag to prevent recursive updates when handling select all checkbox
            self._updating_select_all = False
            
            # Track previous checkbox state to handle indeterminate transitions correctly
            self._previous_select_all_state = None
            
            # Flag to prevent duplicate processing when both checkbox changed and listbox clicked fire
            self._processing_category_selection = False
            
            # Flag to prevent on_parameter_selected from processing during programmatic loading
            self._loading_parameters = False
            
            # Flag to prevent clearing values during selection restoration
            self._restoring_selection = False
            
            # Flag to prevent recursive updates when handling multiple selections
            self._is_updating_selection = False
            
            # Variable to store selected category names for persistence between views
            self.selected_category_names = []
            
            # Store both the function and its result for use in different contexts
            self.get_elementid_value_func_ref = revit_utils.get_elementid_value_func  # Store the function reference
            self.get_elementid_value = revit_utils.get_elementid_value_func()  # Store the result of calling the function
            
            # Create a custom ValuesInfo class
            self.ValuesInfo = self.create_custom_values_info_class()
            
            # Store references to required classes and functions
            # We need to initialize these BEFORE creating any event handlers or methods
            # that might use them
            self.ParameterInfo = ParameterInfo
            self.CategoryInfo = CategoryInfo
            self.PropertyChangedEventArgs = PropertyChangedEventArgs  # Store reference to PropertyChangedEventArgs
            self.strip_accents = strip_accents
            self.normalize = normalize
            self.unicode_category = unicode_category
            self.CAT_EXCLUDED = CAT_EXCLUDED
            
            # Create a custom CategoryItem class
            self.CategoryItem = self.create_custom_category_item_class()
            
            # Multi-view support: list of views to color (auto-detected)
            self.target_views = []  # List of View objects to color
            self.current_sheet = None  # Sheet if we're on one, None otherwise
            
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
            
            # Auto-detect target views (sheet views or single view)
            self.detect_target_views()
                
            # Register for view activation events
            self.uiapp.ViewActivating += self.on_view_activating
            self.uiapp.ViewActivated += self.on_view_activated
                
            # Set up UI event handlers
            self.setup_ui_components()
            
            # Apply SearchableComboBoxStyle programmatically after styles are loaded
            if "SearchableComboBoxStyle" in self.Resources if self.Resources else False:
                try:
                    self.parameterSelector.Style = self.Resources["SearchableComboBoxStyle"]
                    # Set DisplayMemberPath so the ComboBox displays the parameter name
                    self.parameterSelector.DisplayMemberPath = "display_name"
                except Exception as ex:
                    self.logger.error("Error applying SearchableComboBoxStyle: {}".format(str(ex)))
            
            # Load data
            self.load_categories()
            
            # Mark initialization as complete (allow color resets now)
            self._is_initializing = False
            
        except Exception as ex:
            UI.TaskDialog.Show("Error", "Failed to initialize Revit Colorizer: " + str(ex))
            logger.error("Initialization error: %s", str(ex))
            self.Close()
    
    def load_styles(self):
        """Load the common styles ResourceDictionary with theme support."""
        try:
            # Use the proper theme-aware function from styles (lib is in sys.path)
            from styles import load_styles_to_window
            result = load_styles_to_window(self)
            
            if result:
                logger.debug("Loaded styles with theme support")
            else:
                logger.warning("Could not load styles with theme support")
        except ImportError as e:
            logger.error("Failed to import styles: {}".format(str(e)))
            logger.error("lib_path: {}, in sys.path: {}".format(lib_path, lib_path in sys.path))
            raise
        except Exception as e:
            logger.warning("Could not load styles: {}. Using default styles.".format(str(e)))
            import traceback
            logger.debug("Style loading error details: {}".format(traceback.format_exc()))
    
    def set_busy(self, is_busy, message="Loading..."):
        """Show or hide the busy overlay indicator."""
        try:
            if is_busy:
                self.busyOverlay.Visibility = Visibility.Visible
                self.busyTextBlock.Text = message
            else:
                self.busyOverlay.Visibility = Visibility.Collapsed
        except Exception as e:
            logger.debug("Error setting busy indicator: {}".format(str(e)))
    
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
                # Add IsChecked property for checkbox binding
                self._is_checked = False
                
                if para.StorageType == DB.StorageType.Double:
                    self.values_double.append(para.AsDouble())
                elif para.StorageType == DB.StorageType.ElementId:
                    self.values_double.append(para.AsElementId())
                
                # Add a method that mimics the Add method of List
                self.ele_id_add = self.ele_id.append
            
            # IsChecked property with getter and setter
            @property
            def IsChecked(self):
                return self._is_checked
                
            @IsChecked.setter
            def IsChecked(self, value):
                self._is_checked = value
            
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
    
    def create_custom_category_item_class(self):
        """Create a custom version of CategoryItem that uses our stored PropertyChangedEventArgs reference."""
        outer_self = self
        # Store reference to System.ComponentModel.PropertyChangedEventArgs for the inner class
        PropertyChangedEventArgs_ref = self.PropertyChangedEventArgs
        Object_ref = Object
        INotifyPropertyChanged_ref = INotifyPropertyChanged
        
        class CustomCategoryItem(Object_ref, INotifyPropertyChanged_ref):
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
                    # Use the stored reference to PropertyChangedEventArgs
                    self._propertyChanged(self, PropertyChangedEventArgs_ref(property_name))
        
        return CustomCategoryItem
    
    def setup_ui_components(self):
        """Set up UI components and event handlers."""        
        # Apply and Reset buttons
        self.applyButton.Click += self.on_apply_colors
        self.resetButton.Click += self.on_reset_colors
                        
        # Parameter type selection
        self.instanceRadioButton.Checked += self.on_parameter_type_changed
        self.typeRadioButton.Checked += self.on_parameter_type_changed
        
        # Parameter selection
        self.parameterSelector.SelectionChanged += self.on_parameter_selected
        self.parameterSelector.DropDownClosed += self.on_parameter_dropdown_closed
        self.parameterSelector.DropDownOpened += self.on_parameter_dropdown_opened
        
        # Store unfiltered parameter list for search filtering
        self.all_parameters = []

        # Refresh parameters button
        self.refreshParametersButton.Click += self.on_refresh_parameters
        # Load refresh icon image
        try:
            script_dir = os.path.dirname(__file__)
            panel_dir = os.path.dirname(script_dir)
            tab_dir = os.path.dirname(panel_dir)
            extension_dir = os.path.dirname(tab_dir)
            refresh_icon_path = os.path.join(extension_dir, 'lib', 'styles', 'icons', 'refresh.png')
            if os.path.exists(refresh_icon_path):
                from System.Windows.Media.Imaging import BitmapImage
                from System import Uri
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(refresh_icon_path)
                bitmap.EndInit()
                self.refreshParametersImage.Source = bitmap
            else:
                logger.warning("Refresh icon not found at: {}".format(refresh_icon_path))
        except Exception as e:
            logger.warning("Could not load refresh icon: {}".format(str(e)))
        
        # Values list
        self.valuesListBox.MouseDoubleClick += self.on_value_double_click
        self.valuesListBox.SelectionChanged += self.on_value_click
        self.valuesListBox.PreviewMouseLeftButtonUp += self.on_values_list_mouse_up
        
        # Connect the checkbox events in the values list
        # These are defined in the XAML file
        self.ValueCheckbox_Changed = self.ValueCheckbox_Changed
        self.HeaderCheckBox_Changed = self.HeaderCheckBox_Changed 

        # Checkbox events
        self.showElementsCheckbox.Checked += self.on_show_elements_changed
        self.showElementsCheckbox.Unchecked += self.on_show_elements_changed
        self.overrideProjectionCheckbox.Checked += self.on_override_projection_changed
        self.overrideProjectionCheckbox.Unchecked += self.on_override_projection_changed
        
        # Select All Categories checkbox events
        self.SelectAllCategories_Changed = self.SelectAllCategories_Changed
        
        # Setting initial values for checkboxes
        self.overrideProjectionCheckbox.IsChecked = False
        self.showElementsCheckbox.IsChecked = True
        
        # Default to instance parameters being selected
        self.instanceRadioButton.IsChecked = True
        
        # Add a handler for mouse click events on the ListBox to detect checkbox clicks
        self.categoryListBox.PreviewMouseLeftButtonUp += self.on_category_listbox_clicked

    def load_categories(self):
        """Load categories from the current view."""
        try:
            if not self.selected_category_names:
                self.store_selected_categories()
            
            # Clear existing items
            self.categoryListBox.Items.Clear()
            
            # Get categories using our wrapper method that provides proper context
            categories = self.get_used_categories(self.active_view)
            
            # This prevents checkbox event handlers from running during restoration
            self._processing_restored_selection = True
            
            # Create category items for the list
            category_items_added = []
            for category in categories:
                category_item = self.CategoryItem(category)
                # Set reference to this window BEFORE adding to listbox
                category_item.set_window_reference(self)
                self.categoryListBox.Items.Add(category_item)
                category_items_added.append(category_item)
                
            # Process UI updates
            self.categoryListBox.UpdateLayout()
            
            # Try to restore previously selected categories
            categories_selected = False
            
            # If we have stored category names, try to select them
            if self.selected_category_names:
                
                # Select categories by name - all at once before processing anything
                selected_count = 0
                for item in category_items_added:
                    if item.name in self.selected_category_names:
                        item.IsSelected = True
                        categories_selected = True
                        selected_count += 1
                                    
            
            # Process the selection to load parameters etc.
            if self.categoryListBox.Items.Count > 0:
                # Now process the selection (flag still prevents store_selected_categories from running)
                self.process_category_selection()
                
                # After processing, turn off the flag
                self._processing_restored_selection = False
                
                # Now that all categories are restored, store them again to ensure consistency
                self.store_selected_categories()
                
                # Update the select all checkbox state to reflect the current selection
                self.update_select_all_checkbox_state()
                
                # Removed final selections verification logging
                
        except Exception as ex:
            self._processing_restored_selection = False  # Make sure to reset flag on error
            self.logger.error("Error loading categories: %s", str(ex))
    

    def get_used_categories(self, active_view, excluded_cats=None):
        """Get all used categories and their parameters in the selected view(s).
        
        When multi-view mode is enabled, this aggregates categories from all selected views.
        """
        try:
            if excluded_cats is None:
                excluded_cats = self.CAT_EXCLUDED
            
            # Get elements from all selected views (or just active view if single-view mode)
            views_to_query = self.get_target_views()
            
            elements = []
            for view in views_to_query:
                try:
                    collector = self.DB.FilteredElementCollector(self.doc, view.Id) \
                                .WhereElementIsNotElementType() \
                                .WhereElementIsViewIndependent()
                    elements.extend(collector.ToElements())
                except Exception as ex:
                    self.logger.debug("Error collecting elements from view {}: {}".format(view.Name, str(ex)))
                        
            categories = []
                    
            # Store references to avoid scope issues
            get_elementid_value = self.get_elementid_value
            
            # Create a simple, self-contained version of strip_accents that doesn't rely on external functions
            def simple_strip_accents(text):
                """Simplified function to strip accents from text without external dependencies."""
                try:
                    # For non-ASCII characters, just keep the ASCII ones
                    result = ""
                    for char in text:
                        if ord(char) < 128:  # ASCII range
                            result += char
                        else:
                            # Try to replace common accented characters
                            if char in 'áàâäãåā':
                                result += 'a'
                            elif char in 'éèêëēė':
                                result += 'e'
                            elif char in 'íìîïī':
                                result += 'i'
                            elif char in 'óòôöõøō':
                                result += 'o'
                            elif char in 'úùûüū':
                                result += 'u'
                            elif char in 'ñń':
                                result += 'n'
                            elif char in 'çć':
                                result += 'c'
                            elif char in 'ÿ':
                                result += 'y'
                            elif char in 'žźż':
                                result += 'z'
                            elif char in 'šś':
                                result += 's'
                            else:
                                # Keep other characters as is
                                result += char
                    return result
                except:
                    # Failsafe - if anything goes wrong, just return the original text
                    return text
            
            # Create custom parameter and category wrapper classes that use our local strip_accents function
            class LocalParameterInfo(object):
                """Class to store parameter information with local strip_accents."""
                def __init__(self, param_type, para):
                    self.param_type = param_type  # 0 for instance, 1 for type
                    self.rl_par = para
                    self.par = para.Definition
                    self.name = simple_strip_accents(para.Definition.Name)

                def __str__(self):
                    return self.name
            
            class LocalCategoryInfo(object):
                """Class to store category information with local strip_accents."""
                def __init__(self, category, parameters):
                    self.name = simple_strip_accents(category.Name)
                    self.cat = category
                    self.int_id = get_elementid_value(category.Id)
                    self.par = parameters

                def __str__(self):
                    return self.name
            
            # Use our local classes instead of the original ones
            for element in elements:
                if element.Category is None:
                    continue
                    
                current_cat_id = get_elementid_value(element.Category.Id)
                
                # Skip excluded categories and already processed categories
                if current_cat_id in excluded_cats:
                    continue
                if any(x.int_id == current_cat_id for x in categories):
                    continue
                    
                # Get instance parameters
                instance_parameters = []
                for param in element.Parameters:
                    if param.Definition.BuiltInParameter not in (self.DB.BuiltInParameter.ELEM_CATEGORY_PARAM, 
                                                                self.DB.BuiltInParameter.ELEM_CATEGORY_PARAM_MT):
                        instance_parameters.append(LocalParameterInfo(0, param))
                
                # Get type parameters
                type_element = element.Document.GetElement(element.GetTypeId())
                if type_element is None:
                    continue
                    
                type_parameters = []
                for param in type_element.Parameters:
                    if param.Definition.BuiltInParameter not in (self.DB.BuiltInParameter.ELEM_CATEGORY_PARAM, 
                                                                self.DB.BuiltInParameter.ELEM_CATEGORY_PARAM_MT):
                        type_parameters.append(LocalParameterInfo(1, param))
                
                # Combine all parameters
                all_parameters = instance_parameters + type_parameters
                all_parameters.sort(key=lambda x: x.name.upper())
                
                # Add category to the list
                categories.append(LocalCategoryInfo(element.Category, all_parameters))
            
            # Sort categories by name
            categories.sort(key=lambda x: x.name)
            return categories
        except Exception as ex:
            self.logger.error("Error in get_used_categories: %s", str(ex))
            return []

    def on_category_listbox_clicked(self, sender, args):
        """Handle mouse clicks on the category list box to detect checkbox clicks."""
        try:
            
            # Don't schedule delayed processing - on_category_checkbox_changed will handle it synchronously
            # This prevents duplicate processing when both events fire
            # The checkbox changed handler will process the selection immediately
            
            self.apply_colors_event.Raise()
            
            
        except Exception as ex:
            self.logger.error("Category click error: %s", str(ex))
    
    def get_parameter_count(self):
        """Safely get the count of parameters, handling both ItemsSource and Items."""
        if self.parameterSelector.ItemsSource is not None:
            return self.parameterSelector.ItemsSource.Count
        else:
            return self.parameterSelector.Items.Count
    
    
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
            
            # Prevent duplicate processing when both checkbox changed and listbox clicked fire
            if getattr(self, '_processing_category_selection', False):
                return
            
            self._processing_category_selection = True
            
            try:
                # Store currently selected parameter name to try to preserve it
                # Handle both ItemsSource and Items scenarios
                current_param_name = None
                try:
                    if self.parameterSelector.SelectedItem:
                        current_param_name = self.parameterSelector.SelectedItem.parameter_info.name
                except:
                    # If accessing SelectedItem fails (e.g., when ItemsSource is being changed), skip preservation
                    current_param_name = None
                
                # Get current parameter type (instance/type)
                current_param_type_is_instance = self.instanceRadioButton.IsChecked
                
                # Load parameters for all selected categories while preserving selection
                self.load_parameters_for_categories(current_param_name, current_param_type_is_instance)
                
                
                # Count the number of selected categories
                selected_count = sum(1 for item in self.categoryListBox.Items if item.IsSelected)
                
                if selected_count == 0:
                    # Clear parameter selector and show placeholder
                    # Explicitly clear selection first
                    self.parameterSelector.SelectedItem = None
                    self.parameterSelector.SelectedIndex = -1
                    # Then clear ItemsSource and Items
                    if self.parameterSelector.ItemsSource is not None:
                        self.parameterSelector.ItemsSource = None
                    self.parameterSelector.Items.Clear()
                    # Force UI update to ensure placeholder shows
                    self.parameterSelector.UpdateLayout()
                    # Clear values listbox
                    if self.valuesListBox.ItemsSource is not None:
                        self.valuesListBox.ItemsSource = None
                    self.valuesListBox.Items.Clear()
                    # Don't reset colors during initialization
                    if not getattr(self, '_is_initializing', False):
                        self.reset_colors_event.Raise()
                
                # Force an update to the UI
                self.categoryListBox.UpdateLayout()
                
                # Update the stored selected categories after all processing is done
                # But only if we're not in the middle of a restore operation
                if not self._processing_restored_selection:
                    self.store_selected_categories()
                    # Also update the select all checkbox state when not restoring
                    self.update_select_all_checkbox_state()
            finally:
                # Always clear the flag when done
                self._processing_category_selection = False
            
        except Exception as ex:
            self._processing_category_selection = False  # Make sure to clear flag on error
            self.logger.error("Process category selection error: %s", str(ex))
    
    def load_parameters_for_categories(self, preserve_param_name=None, preserve_param_type_is_instance=None):
        """Load common parameters for all selected categories.
        
        Args:
            preserve_param_name: Name of parameter to preserve selection for
            preserve_param_type_is_instance: Whether the parameter to preserve is an instance parameter
        """
        try:
            
            # Set flag to prevent on_parameter_selected from processing during programmatic loading
            self._loading_parameters = True
            
            # Clear parameter list - handle both ItemsSource and Items
            
            # Clear ItemsSource first if it's set, then clear Items
            if self.parameterSelector.ItemsSource is not None:
                self.parameterSelector.ItemsSource = None
            self.parameterSelector.Items.Clear()
            
            
            # Get all selected categories
            selected_categories = self.get_selected_categories()
            
            if not selected_categories:
                self.statusText.Text = "No categories selected."
                # Clear parameter selector and show placeholder
                # Explicitly clear selection first
                self.parameterSelector.SelectedItem = None
                self.parameterSelector.SelectedIndex = -1
                # Then clear ItemsSource and Items
                if self.parameterSelector.ItemsSource is not None:
                    self.parameterSelector.ItemsSource = None
                self.parameterSelector.Items.Clear()
                # Force UI update to ensure placeholder shows
                self.parameterSelector.UpdateLayout()
                # Clear values listbox
                if self.valuesListBox.ItemsSource is not None:
                    self.valuesListBox.ItemsSource = None
                self.valuesListBox.Items.Clear()
                # Don't reset colors during initialization
                if not getattr(self, '_is_initializing', False):
                    self.reset_colors_event.Raise()
                # Clear loading flag before returning
                self._loading_parameters = False
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
            
            # Check if all categories are selected
            total_categories = self.categoryListBox.Items.Count
            all_categories_selected = len(selected_categories) == total_categories
            
            if all_categories_selected:
                # When all categories are selected, show unique parameters from all categories
                unique_params = {}  # Use dict to avoid duplicates by name
                
                for category in selected_categories:
                    category_params = [p for p in category.par if p.param_type == param_type_code]
                    for param in category_params:
                        # Use parameter name as key to avoid duplicates
                        if param.name not in unique_params:
                            unique_params[param.name] = param
                
                # Convert back to list and sort
                common_params = list(unique_params.values())
                common_params.sort(key=lambda x: x.name.upper())
                
            else:
                # When not all categories are selected, use common parameters logic
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
            
            # Store all parameters for search filtering
            self.all_parameters = []
            for param in common_params:
                self.all_parameters.append(self.ParameterDisplayItem(param))
            
            # Add common parameters to dropdown with display wrapper
            # Check if ItemsSource is set (SearchableComboBox uses ItemsSource)
            selected_index = -1
            params_added = 0
            
            # If ItemsSource was set, we need to use ItemsSource instead of Items
            # Create a new collection and set it as ItemsSource
            from System.Collections.ObjectModel import ObservableCollection
            items_collection = ObservableCollection[Object]()
            
            for i, param_item in enumerate(self.all_parameters):
                items_collection.Add(param_item)
                params_added += 1
                # Check if this is the parameter to preserve
                if preserve_param_name and param_item.parameter_info.name == preserve_param_name:
                    selected_index = i
            
            # Set ItemsSource (this will work whether ItemsSource was previously set or not)
            self.parameterSelector.ItemsSource = items_collection
            
            
            # Clear the loading flag before setting selection (which may trigger SelectionChanged)
            self._loading_parameters = False
            
            # Restore selection if possible, otherwise leave unselected (show placeholder)
            if selected_index >= 0:
                self.parameterSelector.SelectedIndex = selected_index
            elif self.get_parameter_count() > 0:
                # Don't auto-select - explicitly set to -1 to show placeholder
                self.parameterSelector.SelectedIndex = -1
                self.parameterSelector.UpdateLayout()
                param_type_name = "instance" if is_instance else "type"
                
                pass
            else:
                # No parameters found - ensure placeholder is visible
                self.parameterSelector.SelectedIndex = -1
                self.parameterSelector.UpdateLayout()
                if self.valuesListBox.ItemsSource is not None:
                    self.valuesListBox.ItemsSource = None
                self.valuesListBox.Items.Clear()
                
        except Exception as ex:
            # Make sure to clear the loading flag on error
            self._loading_parameters = False
            self.logger.error("Parameter loading error: %s", str(ex))
    
    def on_parameter_type_changed(self, sender, args):
        """Handle parameter type (instance/type) changed event."""
        try:
            # Reload parameters for all selected categories with the new parameter type
            self.load_parameters_for_categories()
        except Exception as ex:
            self.logger.error("Parameter type change error: %s", str(ex))
    
    def on_parameter_selected(self, sender, args):
        """Handle parameter selection change."""
        try:
            # Skip processing if we're in the middle of programmatically loading parameters
            if getattr(self, '_loading_parameters', False):
                return
            
            # Skip clearing values if we're restoring selection (to prevent clearing during ItemsSource restoration)
            restoring_selection = getattr(self, '_restoring_selection', False)
            
            # Store the selected parameter name for restoration after dropdown closes (if filtering occurred)
            # This ensures we can restore the selection even if ItemsSource is cleared
            if self.parameterSelector.SelectedItem:
                try:
                    self._selected_param_name_before_search = self.parameterSelector.SelectedItem.parameter_info.name
                except:
                    pass
            
            # Only clear values if we're not restoring (to prevent clearing during ItemsSource restoration)
            if not restoring_selection:
                # Clear current values list
                if hasattr(self, 'valuesListBox') and self.valuesListBox:
                    if self.valuesListBox.ItemsSource is not None:
                        self.valuesListBox.ItemsSource = None

                # Clear listbox checkboxes and header checkbox
                self.valuesListBox.Items.Clear()
                self.headerCheckBox.IsChecked = False

                # Set selection to none
                self.uidoc.Selection.SetElementIds(self.System.Collections.Generic.List[self.DB.ElementId]())
            
            # Get selected parameter and category
            selected_param = self.get_selected_parameter()
            if not selected_param:
                # If restoring and no selection, don't clear values - they should already be set
                if restoring_selection:
                    return
                return
            
            # Fetch values for the parameter
            values = self.get_parameter_values(selected_param, self.active_view)
            
            # Set values to list
            if values:
                self.valuesListBox.ItemsSource = values
                # Force immediate UI update
                self.valuesListBox.UpdateLayout()
                self.valuesListBox.InvalidateVisual()
                # Also update the window to ensure visibility
                self.UpdateLayout()
                
                
                # Check for matching schema file
                schema_path = self.check_for_matching_schema(selected_param.name)
                if schema_path:
                    self.load_color_schema_from_file(schema_path)
            
            # Update UI status
            param_source_txt = "Instance" if self.instanceRadioButton.IsChecked else "Type"
            
            # Check if all categories are selected to add helpful context
            selected_categories = self.get_selected_categories()
            total_categories = self.categoryListBox.Items.Count
            all_categories_selected = len(selected_categories) == total_categories
            
            self.apply_colors_event.Raise()
            
        except Exception as ex:
            self.logger.error("Parameter selection error: %s", str(ex))
    
    def on_parameter_dropdown_closed(self, sender, args):
        """Handle parameter dropdown closed event."""
        try:
            
            # Only clear and restore if items have been filtered (current count < all_parameters count)
            # This prevents clearing items when dropdown closes due to programmatic selection changes
            if hasattr(self, 'all_parameters') and self.all_parameters:
                current_count = self.get_parameter_count()
                all_count = len(self.all_parameters)
                
                # Only restore if items have been filtered (fewer items than all_parameters)
                if current_count < all_count:
                    
                    # Store current selection name BEFORE clearing (for reliable restoration)
                    current_selection_name = None
                    if self.parameterSelector.SelectedItem:
                        try:
                            current_selection_name = self.parameterSelector.SelectedItem.parameter_info.name
                        except:
                            pass
                    
                    # Also check if we have a stored name from dropdown opened or previous selection
                    if hasattr(self, '_selected_param_name_before_search') and self._selected_param_name_before_search:
                        current_selection_name = self._selected_param_name_before_search
                    
                    # Set flag to prevent on_parameter_selected from clearing values during restoration
                    # MUST be set BEFORE clearing ItemsSource, as clearing triggers SelectionChanged
                    self._restoring_selection = True
                    
                    # Temporarily hide values listbox to prevent showing wrong values during restoration
                    values_listbox_was_visible = True
                    if hasattr(self, 'valuesListBox') and self.valuesListBox:
                        values_listbox_was_visible = self.valuesListBox.Visibility == System.Windows.Visibility.Visible
                        if values_listbox_was_visible:
                            self.valuesListBox.Visibility = System.Windows.Visibility.Hidden
                    
                    # Clear search filter and restore all parameters
                    # Handle both ItemsSource and Items
                    if self.parameterSelector.ItemsSource is not None:
                        self.parameterSelector.ItemsSource = None
                    self.parameterSelector.Items.Clear()
                    
                    
                    # Temporarily remove SelectionChanged handler to prevent auto-selection of first item
                    # when ItemsSource is set
                    try:
                        self.parameterSelector.SelectionChanged -= self.on_parameter_selected
                    except:
                        pass
                    
                    # Set SelectedIndex to -1 BEFORE setting ItemsSource to prevent auto-selection
                    self.parameterSelector.SelectedIndex = -1
                    
                    # Restore using ItemsSource (SearchableComboBox uses ItemsSource)
                    from System.Collections.ObjectModel import ObservableCollection
                    items_collection = ObservableCollection[Object]()
                    for param_item in self.all_parameters:
                        items_collection.Add(param_item)
                    self.parameterSelector.ItemsSource = items_collection
                    
                    
                    # Restore selection after ItemsSource is set (match by parameter name)
                    if current_selection_name:
                        try:
                            # Find the matching item by parameter name
                            target_item = None
                            for item in items_collection:
                                if hasattr(item, 'parameter_info') and item.parameter_info.name == current_selection_name:
                                    target_item = item
                                    break
                            
                            if target_item:
                                # Set selection IMMEDIATELY with handler still removed to prevent event firing
                                self.parameterSelector.SelectedItem = target_item
                                
                                # Re-add SelectionChanged handler AFTER selection is set
                                try:
                                    self.parameterSelector.SelectionChanged += self.on_parameter_selected
                                except:
                                    pass
                                
                                # Explicitly call on_parameter_selected to ensure values listbox is populated
                                # Now that handler is re-added, this will work correctly
                                self.on_parameter_selected(self.parameterSelector, None)
                                
                                # Force UI refresh to ensure values listbox is visible
                                if hasattr(self, 'valuesListBox') and self.valuesListBox:
                                    self.valuesListBox.UpdateLayout()
                                    self.valuesListBox.InvalidateVisual()
                                    # Restore visibility if it was hidden
                                    if values_listbox_was_visible:
                                        self.valuesListBox.Visibility = Visibility.Visible
                        except Exception as ex:
                            self.logger.debug("Error restoring selection: {}".format(str(ex)))
                            # Re-add handler even on error to prevent losing it
                            try:
                                self.parameterSelector.SelectionChanged += self.on_parameter_selected
                            except:
                                pass
                            # Restore visibility even on error
                            if hasattr(self, 'valuesListBox') and self.valuesListBox and values_listbox_was_visible:
                                self.valuesListBox.Visibility = Visibility.Visible
                    else:
                        # Re-add handler even if no selection to restore
                        try:
                            self.parameterSelector.SelectionChanged += self.on_parameter_selected
                        except:
                            pass
                        # Restore visibility
                        if hasattr(self, 'valuesListBox') and self.valuesListBox and values_listbox_was_visible:
                            self.valuesListBox.Visibility = Visibility.Visible
                    
                    # Clear restoration flag AFTER all restoration is complete
                    self._restoring_selection = False
                    
                    
                    # Clear stored selection references
                    if hasattr(self, '_selected_item_before_search'):
                        self._selected_item_before_search = None
                    if hasattr(self, '_selected_param_name_before_search'):
                        self._selected_param_name_before_search = None
        except Exception as ex:
            self.logger.error("Parameter dropdown closed error: %s", str(ex))
    
    def on_parameter_dropdown_opened(self, sender, args):
        """Handle parameter dropdown opened event - initialize search filter."""
        try:
            # Store current selection before filtering (store by name for reliable restoration)
            if self.parameterSelector.SelectedItem:
                self._selected_item_before_search = self.parameterSelector.SelectedItem
                # Also store the parameter name for reliable matching
                try:
                    self._selected_param_name_before_search = self.parameterSelector.SelectedItem.parameter_info.name
                except:
                    self._selected_param_name_before_search = None
            
            # Use Dispatcher to wait for popup to be fully rendered before finding SearchTextBox
            from System.Windows.Threading import DispatcherPriority
            self.parameterSelector.Dispatcher.BeginInvoke(
                DispatcherPriority.Loaded,
                System.Action(self._initialize_search_textbox)
            )
        except Exception as ex:
            self.logger.error("Error initializing search filter: %s", str(ex))
            import traceback
            self.logger.debug("Traceback: %s", traceback.format_exc())
    
    def _initialize_search_textbox(self):
        """Initialize the search textbox after popup is loaded."""
        try:
            
            # Check if template exists
            if not self.parameterSelector.Template:
                self.logger.warning("ComboBox has no template")
                return
            
            # Try to find Popup first
            popup = self.parameterSelector.Template.FindName("Popup", self.parameterSelector)
            
            # Try to find SearchTextBox directly
            search_textbox = self.parameterSelector.Template.FindName("SearchTextBox", self.parameterSelector)
            
            # If not found directly, try traversing visual tree from Popup
            if not search_textbox and popup:
                try:
                    # Try to find SearchTextBox in Popup's visual tree
                    from System.Windows.Media import VisualTreeHelper
                    if popup.Child:
                        # Try to find by name in visual tree
                        def find_child_by_name(parent, name):
                            if parent is None:
                                return None
                            if hasattr(parent, 'Name') and parent.Name == name:
                                return parent
                            for i in range(VisualTreeHelper.GetChildrenCount(parent)):
                                child = VisualTreeHelper.GetChild(parent, i)
                                result = find_child_by_name(child, name)
                                if result:
                                    return result
                            return None
                        
                        search_textbox = find_child_by_name(popup.Child, "SearchTextBox")
                except Exception as ex:
                    self.logger.debug("Error traversing visual tree: {}".format(str(ex)))
            
            if search_textbox:
                self.logger.debug("Found SearchTextBox successfully")
                # Clear search text
                search_textbox.Text = ""
                # Wire up TextChanged event if not already done
                if not hasattr(self, '_search_textbox_wired'):
                    search_textbox.TextChanged += self.on_search_text_changed
                    search_textbox.KeyDown += self.on_search_textbox_keydown
                    search_textbox.PreviewKeyDown += self.on_search_textbox_preview_keydown
                    self._search_textbox_wired = True
                    self.logger.debug("Wired up SearchTextBox events")
                # Focus the search textbox
                search_textbox.Focus()
                self.logger.debug("Focused SearchTextBox")
            else:
                self.logger.warning("SearchTextBox not found in ComboBox template")
                if not popup:
                    self.logger.warning("Popup not found in ComboBox template")
        except Exception as ex:
            self.logger.error("Error finding search textbox: %s", str(ex))
            import traceback
            self.logger.debug("Traceback: %s", traceback.format_exc())
    
    def on_search_text_changed(self, sender, args):
        """Handle search text changed event - filter parameters."""
        try:
            
            search_text = sender.Text.lower().strip()
            
            # Clear current items - handle ItemsSource
            if self.parameterSelector.ItemsSource is not None:
                self.parameterSelector.ItemsSource = None
            self.parameterSelector.Items.Clear()
            
            
            if not hasattr(self, 'all_parameters') or not self.all_parameters:
                return
            
            # Use ItemsSource for SearchableComboBox
            from System.Collections.ObjectModel import ObservableCollection
            filtered_collection = ObservableCollection[Object]()
            
            if not search_text:
                # Show all parameters if search is empty
                for param_item in self.all_parameters:
                    filtered_collection.Add(param_item)
            else:
                # Filter parameters based on search text
                for param_item in self.all_parameters:
                    param_name = param_item.display_name.lower()
                    if search_text in param_name:
                        filtered_collection.Add(param_item)
            
            self.parameterSelector.ItemsSource = filtered_collection
            
            # Try to restore selection if it matches the filter
            if hasattr(self, '_selected_item_before_search') and self._selected_item_before_search:
                try:
                    # Use ItemsSource since we just set it
                    for item in filtered_collection:
                        if item == self._selected_item_before_search:
                            self.parameterSelector.SelectedItem = item
                            break
                except Exception as ex:
                    self.logger.debug("Error restoring selection in filter: {}".format(str(ex)))
        except Exception as ex:
            self.logger.error("Error filtering parameters: %s", str(ex))
    
    def on_search_textbox_preview_keydown(self, sender, args):
        """Handle PreviewKeyDown events - intercept arrow keys before they're handled."""
        try:
            from System.Windows.Input import Key
            
            # Handle Down arrow key - move focus to first item in results
            if args.Key == Key.Down:
                args.Handled = True
                # Use Dispatcher to set focus after current event completes
                from System.Windows.Threading import DispatcherPriority
                self.parameterSelector.Dispatcher.BeginInvoke(
                    DispatcherPriority.Input,
                    System.Action(lambda: self._focus_first_item())
                )
                return
            
            # Handle Up arrow key
            elif args.Key == Key.Up:
                if self.parameterSelector.Items.Count > 0:
                    args.Handled = True
                    from System.Windows.Threading import DispatcherPriority
                    self.parameterSelector.Dispatcher.BeginInvoke(
                        DispatcherPriority.Input,
                        System.Action(lambda: self._focus_last_item())
                    )
                return
            
            # Handle Enter key
            elif args.Key == Key.Enter:
                if self.parameterSelector.SelectedIndex >= 0:
                    args.Handled = True
                    self.parameterSelector.IsDropDownOpen = False
                return
            
            # Handle Escape key
            elif args.Key == Key.Escape:
                args.Handled = True
                self.parameterSelector.IsDropDownOpen = False
                return
                
        except Exception as ex:
            self.logger.error("Error handling search textbox preview keydown: %s", str(ex))
    
    def _focus_first_item(self):
        """Focus the first ComboBoxItem."""
        try:
            first_item = self._find_first_combobox_item()
            if first_item:
                first_item.Focus()
                if self.parameterSelector.Items.Count > 0:
                    self.parameterSelector.SelectedIndex = 0
        except Exception as ex:
            self.logger.debug("Error focusing first item: {}".format(str(ex)))
    
    def _focus_last_item(self):
        """Focus the last ComboBoxItem."""
        try:
            last_item = self._find_last_combobox_item()
            if last_item:
                last_item.Focus()
                self.parameterSelector.SelectedIndex = self.parameterSelector.Items.Count - 1
        except Exception as ex:
            self.logger.debug("Error focusing last item: {}".format(str(ex)))
    
    def on_search_textbox_keydown(self, sender, args):
        """Handle keyboard events in search textbox - enable arrow key navigation."""
        # Note: PreviewKeyDown handles arrow keys, this handler is kept for compatibility
        # but arrow keys are intercepted in PreviewKeyDown before reaching here
        pass
    
    def _find_first_combobox_item(self):
        """Find the first ComboBoxItem in the dropdown."""
        try:
            from System.Windows.Controls import ComboBoxItem
            from System.Windows.Media import VisualTreeHelper
            
            if not self.parameterSelector.Template:
                return None
            
            popup = self.parameterSelector.Template.FindName("Popup", self.parameterSelector)
            if not popup or not popup.Child:
                return None
            
            # Traverse visual tree to find first ComboBoxItem
            def find_first_combobox_item(parent):
                if parent is None:
                    return None
                if isinstance(parent, ComboBoxItem):
                    return parent
                child_count = VisualTreeHelper.GetChildrenCount(parent)
                for i in range(child_count):
                    child = VisualTreeHelper.GetChild(parent, i)
                    result = find_first_combobox_item(child)
                    if result:
                        return result
                return None
            
            return find_first_combobox_item(popup.Child)
        except Exception as ex:
            self.logger.debug("Error finding first ComboBoxItem: {}".format(str(ex)))
            return None
    
    def _find_last_combobox_item(self):
        """Find the last ComboBoxItem in the dropdown."""
        try:
            from System.Windows.Controls import ComboBoxItem
            from System.Windows.Media import VisualTreeHelper
            
            if not self.parameterSelector.Template:
                return None
            
            popup = self.parameterSelector.Template.FindName("Popup", self.parameterSelector)
            if not popup or not popup.Child:
                return None
            
            # Traverse visual tree to find all ComboBoxItems, return last one
            items = []
            def find_all_combobox_items(parent):
                if parent is None:
                    return
                if isinstance(parent, ComboBoxItem):
                    items.append(parent)
                child_count = VisualTreeHelper.GetChildrenCount(parent)
                for i in range(child_count):
                    child = VisualTreeHelper.GetChild(parent, i)
                    find_all_combobox_items(child)
            
            find_all_combobox_items(popup.Child)
            return items[-1] if items else None
        except Exception as ex:
            self.logger.debug("Error finding last ComboBoxItem: {}".format(str(ex)))
            return None
    
    def on_values_list_mouse_up(self, sender, args):
        """Handle row click for toggling, while allowing selection to proceed."""
        if getattr(self, '_is_updating_selection', False):
            return
            
        try:
            # Check if we clicked on a checkbox or its child - if so, let the checkbox handle it
            from System.Windows.Media import VisualTreeHelper
            point = args.GetPosition(self.valuesListBox)
            hit = VisualTreeHelper.HitTest(self.valuesListBox, point)
            if hit and hit.VisualHit:
                parent = hit.VisualHit
                while parent and parent != self.valuesListBox:
                    if "CheckBox" in parent.GetType().Name:
                        return # Let the checkbox handle its own toggle
                    parent = VisualTreeHelper.GetParent(parent)
            
            # Find the ListBoxItem that was clicked
            item_container = None
            hit_result = VisualTreeHelper.HitTest(self.valuesListBox, point)
            if hit_result and hit_result.VisualHit:
                parent = hit_result.VisualHit
                while parent and parent != self.valuesListBox:
                    if "ListBoxItem" in parent.GetType().Name:
                        item_container = parent
                        break
                    parent = VisualTreeHelper.GetParent(parent)
                
            if item_container and item_container.DataContext:
                item = item_container.DataContext
                target_state = not item.IsChecked
                
                self._is_updating_selection = True
                try:
                    # Sync multi-selection if the clicked item is part of it
                    if self.valuesListBox.SelectedItems.Count > 1 and self.valuesListBox.SelectedItems.Contains(item):
                        for selected_item in self.valuesListBox.SelectedItems:
                            selected_item.IsChecked = target_state
                    else:
                        item.IsChecked = target_state
                    
                    self.valuesListBox.Items.Refresh()
                    self.select_checked_elements()
                finally:
                    self._is_updating_selection = False
                
        except Exception as ex:
            self.logger.debug("Mouse up error: %s" % str(ex))

    def on_value_click(self, sender, args):
        """Handle value item click to select Revit elements based on value."""
        if getattr(self, '_is_updating_selection', False):
            return
            
        try:
# Only select elements in Revit, don't toggle here anymore
            # Toggling is now handled in on_values_list_mouse_up
            if self.showElementsCheckbox.IsChecked:
                self.select_checked_elements()
            
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
                self.apply_colors_event.Raise()
        except Exception as ex:
            self.logger.error("Value double click error: %s", str(ex))
    
    def on_apply_colors(self, sender, args):
        """Handle apply colors button click."""
        try:
            if not self.valuesListBox.Items.Count:
                return
                
            self.apply_colors_event.Raise()
        except Exception as ex:
            self.logger.error("Apply colors error: %s", str(ex))
    
    def on_reset_colors(self, sender, args):
        """Handle reset colors button click."""
        try:
            self.reset_colors_event.Raise()
        except Exception as ex:
            self.logger.error("Reset colors error: %s", str(ex))
    
    def on_add_filters(self, sender, args):
        """Handle add filters button click."""
        try:
            # This would be implemented similar to the CreateFilters class in ColorSplasher
            pass
        except Exception as ex:
            self.logger.error("Add filters error: %s", str(ex))
    
    def on_remove_filters(self, sender, args):
        """Handle remove filters button click."""
        try:
            # This would be implemented similar to the reset colors but focused on filters
            pass
        except Exception as ex:
            self.logger.error("Remove filters error: %s", str(ex))


    def load_color_schema_from_file(self, file_path):
        """Load a color scheme from a file."""
        try:
            if not os.path.exists(file_path):
                return
            
            # Load the schema
            schema_data = []
            with open(file_path, "r") as file:
                for line in file:
                    line = line.strip()
                    if not line:  # Skip empty lines
                        continue
                    
                    # Parse line in format "value::RrGgBb"
                    parts = line.split("::")
                    if len(parts) != 2:
                        continue
                    
                    value = parts[0]
                    rgb_part = parts[1]
                    
                    # Parse RGB values using regex
                    rgb_values = re.split(r'[RGB]', rgb_part)
                    if len(rgb_values) >= 4:  # [0] is empty before the first R
                        r = int(rgb_values[1])
                        g = int(rgb_values[2])
                        b = int(rgb_values[3])
                        
                        schema_data.append({
                            'value': value,
                            'color': Media.Color.FromRgb(r, g, b)
                        })
            
            # Apply to current values if any
            if schema_data and self.valuesListBox.Items:
                # First try matching by value
                for value_item in self.valuesListBox.Items:
                    for schema_item in schema_data:
                        if value_item.value == schema_item['value']:
                            wpf_color = schema_item['color']
                            # Update RGB values
                            value_item.n1 = wpf_color.R
                            value_item.n2 = wpf_color.G
                            value_item.n3 = wpf_color.B
                            # Update color property using the setter
                            value_item.color = wpf_color
                                
                # Refresh the list
                self.valuesListBox.Items.Refresh()
        except Exception as ex:
            self.logger.error("Error loading color schema: %s", str(ex))

    def get_available_schemas(self):
        """Get list of available color schemas in the coloringschemas folder."""
        schemas = []
        script_dir = os.path.dirname(__file__)
        schemas_dir = os.path.join(script_dir, "coloringschemas")
        
        if os.path.exists(schemas_dir):
            for file in os.listdir(schemas_dir):
                if file.endswith(".cschn"):
                    schema_name = os.path.splitext(file)[0]
                    schemas.append({
                        'name': schema_name,
                        'path': os.path.join(schemas_dir, file)
                    })
        
        return schemas

    def check_for_matching_schema(self, parameter_name):
        """Check if there's a color schema matching the parameter name."""
        if not parameter_name:
            return None
        
        schemas = self.get_available_schemas()
        for schema in schemas:
            # Case-insensitive comparison
            if schema['name'].lower() == parameter_name.lower():
                return schema['path']
        
        return None

    def on_show_elements_changed(self, sender, args):
        """Handle show elements checkbox state change."""
        try:
            # If checkbox is unchecked, clear Revit selection
            if not self.showElementsCheckbox.IsChecked:
                # Clear current selection
                self.uidoc.Selection.SetElementIds(self.System.Collections.Generic.List[self.DB.ElementId]())
            else:
                # Select based on checked checkboxes
                self.select_checked_elements()
        except Exception as ex:
            self.logger.error("Show elements checkbox error: %s", str(ex))

    def on_override_projection_changed(self, sender, args):
        """Handle override projection checkbox state change."""
        try:
            self.apply_colors_event.Raise()
        except Exception as ex:
            self.logger.error("Show elements checkbox error: %s", str(ex))

    def SelectAllCategories_Changed(self, sender, args):
        """Handle select all categories checkbox state change."""
        try:
            # Skip processing if we're in the middle of restoring selection or processing category changes
            if self._processing_restored_selection or getattr(self, '_updating_select_all', False):
                return
            
            # Set flag to prevent recursive updates
            self._updating_select_all = True
            
            # Get the state of the select all checkbox
            checkbox_state = self.selectAllCategoriesCheckBox.IsChecked
            
            # Determine the action based on checkbox state
            # Handle indeterminate state based on previous state:
            # - If previous was True and now None: transitioning to unchecked (False)
            # - If previous was False and now None: transitioning to checked (True)
            # - If previous was None: treat as checked (True)
            if checkbox_state is None:  # Indeterminate state
                if self._previous_select_all_state == True:
                    # Was checked, now indeterminate - user is unchecking
                    is_checked = False
                    self.selectAllCategoriesCheckBox.IsChecked = False
                elif self._previous_select_all_state == False:
                    # Was unchecked, now indeterminate - user is checking
                    is_checked = True
                    self.selectAllCategoriesCheckBox.IsChecked = True
                else:
                    # Previous state unknown or None - default to checked
                    is_checked = True
                    self.selectAllCategoriesCheckBox.IsChecked = True
            elif checkbox_state == True:  # Checked state - keep all selected
                is_checked = True
            elif checkbox_state == False:  # Unchecked state - deselect all
                is_checked = False
            else:
                # Fallback: if state is unexpected, deselect all
                is_checked = False
                self.selectAllCategoriesCheckBox.IsChecked = False
            
            # Store current state for next time
            self._previous_select_all_state = checkbox_state if checkbox_state is not None else (True if is_checked else False)
            
            # Update all category items
            if self.categoryListBox.Items:
                # Update all category checkboxes
                for category_item in self.categoryListBox.Items:
                    category_item.IsSelected = is_checked
                
                # Force refresh of the list box
                self.categoryListBox.UpdateLayout()
                
                # Process the category selection after all are updated
                self.process_category_selection()
                
                # Update status
                # Clear the flag
            self._updating_select_all = False
                
        except Exception as ex:
            self._updating_select_all = False  # Make sure to clear flag on error
            self.logger.error("Select all categories error: %s", str(ex))

    def update_select_all_checkbox_state(self):
        """Update the select all checkbox state based on individual category selections."""
        try:
            if not self.categoryListBox.Items.Count:
                return
            
            # Count selected categories
            selected_count = sum(1 for item in self.categoryListBox.Items if item.IsSelected)
            total_count = self.categoryListBox.Items.Count
            
            # Temporarily disable event handling to prevent recursion
            self._updating_select_all = True
            
            # Update select all checkbox state
            if selected_count == 0:
                self.selectAllCategoriesCheckBox.IsChecked = False
            elif selected_count == total_count:
                self.selectAllCategoriesCheckBox.IsChecked = True
            else:
                # WPF checkboxes support indeterminate state (three-state)
                # This shows a filled square when some but not all items are selected
                self.selectAllCategoriesCheckBox.IsThreeState = True
                self.selectAllCategoriesCheckBox.IsChecked = None  # None represents indeterminate state
            
            # Re-enable event handling
            self._updating_select_all = False
            
        except Exception as ex:
            self._updating_select_all = False  # Make sure to clear flag on error
            self.logger.error("Error updating select all checkbox state: %s", str(ex))

    def on_category_checkbox_changed(self, category_item):
        """Handle category checkbox selection changed event."""
        try:
            
            # Skip processing if we're in the middle of restoring selection or updating select all
            if self._processing_restored_selection or getattr(self, '_updating_select_all', False):
                return
                
            # Update the select all checkbox state based on individual category selections
            self.update_select_all_checkbox_state()
                
            # Directly process the selection which will handle the category storage if needed
            self.process_category_selection()
            
                
        except Exception as ex:
            self.logger.error("Category checkbox error: %s", str(ex))

    def on_refresh_parameters(self, sender, args):
        """Handle refresh parameters button click."""
        try:
            # Store current selections
            current_param_type_is_instance = self.instanceRadioButton.IsChecked
            current_param_name = None
            try:
                if self.parameterSelector.SelectedItem:
                    current_param_name = self.parameterSelector.SelectedItem.parameter_info.name
            except:
                # If accessing SelectedItem fails (e.g., when ItemsSource is being changed), skip preservation
                current_param_name = None
                
            # Clear and reload parameters - handle ItemsSource
            if self.parameterSelector.ItemsSource is not None:
                self.parameterSelector.ItemsSource = None
            self.parameterSelector.Items.Clear()
            selected_categories = self.get_selected_categories()
            
            if not selected_categories:
                self.statusText.Text = "No categories selected."
                # Clear parameter selector and show placeholder
                # Explicitly clear selection first
                self.parameterSelector.SelectedItem = None
                self.parameterSelector.SelectedIndex = -1
                # Force UI update to ensure placeholder shows
                self.parameterSelector.UpdateLayout()
                # Clear values listbox
                if self.valuesListBox.ItemsSource is not None:
                    self.valuesListBox.ItemsSource = None
                self.valuesListBox.Items.Clear()
                # Don't reset colors during initialization
                if not getattr(self, '_is_initializing', False):
                    self.reset_colors_event.Raise()
                return
                
            # Get parameters (common or unique based on selection)
            param_type_code = 0 if current_param_type_is_instance else 1
            total_categories = self.categoryListBox.Items.Count
            all_categories_selected = len(selected_categories) == total_categories
            
            if all_categories_selected:
                # When all categories are selected, show unique parameters from all categories
                unique_params = {}  # Use dict to avoid duplicates by name
                
                for category in selected_categories:
                    category_params = [p for p in category.par if p.param_type == param_type_code]
                    for param in category_params:
                        # Use parameter name as key to avoid duplicates
                        if param.name not in unique_params:
                            unique_params[param.name] = param
                
                # Convert back to list and sort
                common_params = list(unique_params.values())
                common_params.sort(key=lambda x: x.name.upper())
            else:
                # When not all categories are selected, use common parameters logic
                common_params = [p for p in selected_categories[0].par if p.param_type == param_type_code]
                
                for category in selected_categories[1:]:
                    category_params = [p for p in category.par if p.param_type == param_type_code]
                    category_param_names = [p.name for p in category_params]
                    common_params = [p for p in common_params if p.name in category_param_names]
                
            # Store all parameters for search filtering
            self.all_parameters = []
            for param in common_params:
                self.all_parameters.append(self.ParameterDisplayItem(param))
            
            # Add parameters to dropdown using ItemsSource
            from System.Collections.ObjectModel import ObservableCollection
            items_collection = ObservableCollection[Object]()
            selected_index = -1
            for i, param_item in enumerate(self.all_parameters):
                items_collection.Add(param_item)
                if current_param_name and param_item.parameter_info.name == current_param_name:
                    selected_index = i
            
            self.parameterSelector.ItemsSource = items_collection
                    
            # Restore selection if possible, otherwise leave unselected (show placeholder)
            if selected_index >= 0:
                self.parameterSelector.SelectedIndex = selected_index
            elif self.get_parameter_count() > 0:
                # Don't auto-select - explicitly set to -1 to show placeholder
                self.parameterSelector.SelectedIndex = -1
                self.parameterSelector.UpdateLayout()
            else:
                # No parameters found - ensure placeholder is visible
                self.parameterSelector.SelectedIndex = -1
                self.parameterSelector.UpdateLayout()
                if self.valuesListBox.ItemsSource is not None:
                    self.valuesListBox.ItemsSource = None
                self.valuesListBox.Items.Clear()
                return
                
            # Reload values
            if self.parameterSelector.SelectedItem:
                self.on_parameter_selected(self.parameterSelector, None)
                
            self.apply_colors_event.Raise()
                
        except Exception as ex:
            self.logger.error("Parameter refresh error: %s", str(ex))

    # ============================================================
    # Multi-View Support Methods (Automatic - no UI)
    # ============================================================
    
    def detect_target_views(self):
        """Auto-detect target views based on context.
        
        If we're on a sheet, target all views on that sheet.
        Otherwise, target just the active view.
        """
        try:
            self.target_views = []
            self.current_sheet = None
            
            # Check if active view is a sheet
            if self.active_view.ViewType == self.DB.ViewType.DrawingSheet:
                # We're on a sheet - get all views on it
                self.current_sheet = self.active_view
                self.target_views = self.get_views_from_sheet(self.current_sheet)
            else:
                # Check if active view is placed on a sheet
                sheet = self.find_sheet_containing_view(self.active_view)
                if sheet:
                    self.current_sheet = sheet
                    self.target_views = self.get_views_from_sheet(sheet)
                else:
                    # Regular view - just use the active view
                    self.target_views = [self.active_view]
            
        except Exception as ex:
            self.logger.error("Error detecting target views: {}".format(str(ex)))
            # Fallback to active view
            self.target_views = [self.active_view]
    
    def get_views_from_sheet(self, sheet):
        """Get all colorable views placed on a sheet."""
        try:
            placed_view_ids = sheet.GetAllPlacedViews()
            
            views = []
            for view_id in placed_view_ids:
                view = self.doc.GetElement(view_id)
                if view and self.can_color_view(view):
                    views.append(view)
            
            return views if views else [self.active_view]
            
        except Exception as ex:
            self.logger.error("Error getting views from sheet: {}".format(str(ex)))
            return [self.active_view]
    
    def find_sheet_containing_view(self, view):
        """Find the sheet that contains the given view."""
        try:
            # Get all sheets in the document
            sheets = self.DB.FilteredElementCollector(self.doc) \
                        .OfClass(self.DB.ViewSheet) \
                        .ToElements()
            
            for sheet in sheets:
                placed_view_ids = sheet.GetAllPlacedViews()
                if view.Id in placed_view_ids:
                    return sheet
            
            return None
            
        except Exception as ex:
            self.logger.error("Error finding sheet for view: {}".format(str(ex)))
            return None
    
    def can_color_view(self, view):
        """Check if a view can have element colors applied."""
        try:
            # Check if view supports temporary visibility modes
            if not view.CanUseTemporaryVisibilityModes():
                return False
            
            # Exclude certain view types
            excluded_types = [
                self.DB.ViewType.ProjectBrowser,
                self.DB.ViewType.SystemBrowser,
                self.DB.ViewType.Schedule,
                self.DB.ViewType.Legend,
                self.DB.ViewType.Report,
                self.DB.ViewType.DraftingView,
                self.DB.ViewType.DrawingSheet,  # The sheet itself, not the views on it
            ]
            
            if view.ViewType in excluded_types:
                return False
            
            return True
            
        except Exception as ex:
            self.logger.error("Error checking view compatibility: {}".format(str(ex)))
            return False
    
    def get_target_views(self):
        """Get the list of views to color (auto-detected)."""
        if self.target_views:
            return self.target_views
        return [self.active_view]
    
    # ============================================================
    # Category Storage Methods
    # ============================================================

    def store_selected_categories(self):
        """Store the names of currently selected categories for persistence between views."""
        self.selected_category_names = []
        for item in self.categoryListBox.Items:
            if item.IsSelected:
                self.selected_category_names.append(item.name)
        return self.selected_category_names

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
            # Mark window as inactive to disable event handling
            self.is_window_active = False
            
            # Safely unregister view activation events
            try:
                if hasattr(self, 'uiapp'):
                    # Store references to the event handlers
                    view_activating_handler = self.on_view_activating
                    view_activated_handler = self.on_view_activated
                    
                    # Try to unregister each event handler separately
                    try:
                        if hasattr(self.uiapp, 'ViewActivating'):
                            self.uiapp.ViewActivating -= view_activating_handler
                    except:
                        pass  # Ignore errors when unregistering ViewActivating
                        
                    try:
                        if hasattr(self.uiapp, 'ViewActivated'):
                            self.uiapp.ViewActivated -= view_activated_handler
                    except:
                        pass  # Ignore errors when unregistering ViewActivated
            except Exception as ex:
                self.logger.error("Error unregistering view events: %s", str(ex))
                
            # Reset colors in all target views
            try:
                views_to_reset = self.get_target_views()
                if views_to_reset:
                    self.reset_colors_handler.set_specific_views(views_to_reset)
                    self.reset_colors_event.Raise()
            except Exception as ex:
                self.logger.error("Error resetting colors: %s", str(ex))
                
        except Exception as ex:
            # Just log the error, don't show to user since window is closing
            self.logger.error("Error in window closing: %s", str(ex))
    
    def OnClosing(self, e):
        """Override WPF window's OnClosing method to ensure colors are reset when window closes.
        This catches ALL closing scenarios including the Windows X button."""
        try:
            # Mark window as inactive to disable event handling
            self.is_window_active = False
            
            # Safely unregister view activation events
            try:
                if hasattr(self, 'uiapp'):
                    # Store references to the event handlers
                    view_activating_handler = self.on_view_activating
                    view_activated_handler = self.on_view_activated
                    
                    # Try to unregister each event handler separately
                    try:
                        if hasattr(self.uiapp, 'ViewActivating'):
                            self.uiapp.ViewActivating -= view_activating_handler
                    except:
                        pass  # Ignore errors when unregistering ViewActivating
                        
                    try:
                        if hasattr(self.uiapp, 'ViewActivated'):
                            self.uiapp.ViewActivated -= view_activated_handler
                    except:
                        pass  # Ignore errors when unregistering ViewActivated
            except Exception as ex:
                self.logger.error("Error unregistering view events: %s", str(ex))
                
            # Reset colors in all target views
            try:
                views_to_reset = self.get_target_views()
                if views_to_reset:
                    self.reset_colors_handler.set_specific_views(views_to_reset)
                    self.reset_colors_event.Raise()
            except Exception as ex:
                self.logger.error("Error resetting colors: %s", str(ex))
                    
        except Exception as ex:
            # Just log the error, don't show to user since window is closing
            self.logger.error("Error in OnClosing: %s", str(ex))
        
        # Call the base class implementation using the fully qualified name
        try:
            from pyrevit import forms  # Import at the method level to ensure it's available
            forms.WPFWindow.OnClosing(self, e)
        except Exception as ex:
            self.logger.error("Error in base class OnClosing: %s", str(ex))
    
    def on_view_activating(self, sender, args):
        """Handle Revit view activating events (fires BEFORE the view changes).
        
        When on a sheet with multiple views, we don't reset colors when navigating
        between views on the same sheet.
        """
        try:
            # Skip if window is not active or still initializing
            if not self.is_window_active or getattr(self, '_is_initializing', False):
                return
            
            # If we're working with multiple views on a sheet, don't reset on view change
            if self.current_sheet and len(self.target_views) > 1:
                return
            
            # Store the current view before it changes
            current_view = self.active_view
            
            # Store selected categories for persistence between views
            self.store_selected_categories()
            
            # Reset colors in the current view
            self.reset_colors_handler.set_specific_view(current_view)
            self.reset_colors_event.Raise()
                
        except Exception as ex:
            self.logger.error("Error handling view activation: %s", ex)

    def on_view_activated(self, sender, args):
        """Handle Revit view activated events (fires AFTER the view changes).
        
        When on a sheet with multiple views, we don't auto-reload when navigating
        between views on the same sheet.
        """
        try:
            # Skip if window is not active
            if not self.is_window_active:
                return
            
            # Always update the active view reference
            self.active_view = self.doc.ActiveView
            
            # If we're working with multiple views on a sheet, don't auto-reload
            if self.current_sheet and len(self.target_views) > 1:
                return
            
            # Re-detect target views (user may have navigated to a different sheet or view)
            self.detect_target_views()
                
            # Reload categories for the new context
            self.load_categories()
            self.apply_colors_event.Raise()
        except Exception as ex:
            self.logger.error("Error handling view activation: %s", ex)
    

    # Helper methods to get selected items
    def get_selected_category(self):
        """Get the first checked category."""
        # Check all category items and return the first one that is checked
        for item in self.categoryListBox.Items:
            if item.IsSelected:
                return item.category_info
        return None
    
    def get_parameter_values(self, param, view):
        """Get all values for a parameter across all selected categories and target views.
        
        Aggregates elements from all target views (sheet views or single view).
        The 'view' parameter is kept for compatibility but target_views is used.
        """
        values = []
        used_colors = set()
        processed_element_ids = set()  # Track processed elements to avoid duplicates across views
        
        # Get all selected categories
        selected_categories = self.get_selected_categories()
        
        # Get all target views (auto-detected based on sheet context)
        views_to_query = self.get_target_views()
        
        for category in selected_categories:
            # Try to find BuiltInCategory if possible
            bic = None
            for sample_bic in System.Enum.GetValues(DB.BuiltInCategory):
                if category.int_id == int(sample_bic):
                    bic = sample_bic
                    break
            
            if not bic:
                continue
            
            # Get elements from all selected views
            for query_view in views_to_query:
                try:
                    # Get all elements of this category in this view
                    collector = DB.FilteredElementCollector(self.doc, query_view.Id) \
                                .OfCategory(bic) \
                                .WhereElementIsNotElementType() \
                                .WhereElementIsViewIndependent() \
                                .ToElements()
                except Exception as ex:
                    self.logger.debug("Error collecting elements from view {}: {}".format(query_view.Name, str(ex)))
                    continue
                
                for element in collector:
                    # Skip if we've already processed this element (can appear in multiple views)
                    element_int_id = self.get_elementid_value(element.Id)
                    if element_int_id in processed_element_ids:
                        continue
                    processed_element_ids.add(element_int_id)
                    
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
                                if parameter.StorageType == DB.StorageType.Double:
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
                value_item.color = Media.Color.FromRgb(r, g, b)
                
        # Always use gray for None values
        for value_item in none_values:
            value_item.n1 = 192
            value_item.n2 = 192
            value_item.n3 = 192
            # Update the color property with a new WPF color object
            value_item.color = Media.Color.FromRgb(192, 192, 192)
        
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

    def get_parameter_value(self, para):
        """Get parameter value as string."""
        if not para.HasValue:
            return "None"
        
        if para.StorageType == DB.StorageType.Double:
            return para.AsValueString()
        elif para.StorageType == DB.StorageType.ElementId:
            id_val = para.AsElementId()
            if self.get_elementid_value(id_val) >= 0:
                return DB.Element.Name.GetValue(self.doc.GetElement(id_val))
            else:
                return "None"
        elif para.StorageType == DB.StorageType.Integer:
            if self.version > 2021:
                param_type = para.Definition.GetDataType()
                if DB.SpecTypeId.Boolean.YesNo == param_type:
                    return "True" if para.AsInteger() == 1 else "False"
                else:
                    return para.AsValueString()
            else:
                param_type = para.Definition.ParameterType
                if DB.ParameterType.YesNo == param_type:
                    return "True" if para.AsInteger() == 1 else "False"
                else:
                    return para.AsValueString()
        elif para.StorageType == DB.StorageType.String:
            return para.AsString() or "None"
        else:
            return "None"

    def generate_color_range(self, count):
        """Generate a range of visually distinct colors based on count.
        
        Args:
            count: Number of colors needed
            
        Returns:
            List of (r, g, b) tuples representing colors
        """

        if count <= 5:
            # Jewel Bright palette (from Tableau)
            jewel_bright = [
                (253, 111, 48),   # #fd6f30 - Orange
                (249, 167, 41),   # #f9a729 - Yellow-orange
                (249, 210, 60),   # #f9d23c - Yellow
                (95, 187, 104),   # #5fbb68 - Green
                (100, 205, 204),  # #64cdcc - Teal
            ]
            return jewel_bright[:count]
            
        # Define color palettes based on count ranges
        if count <= 9:
            # Jewel Bright palette (from Tableau)
            jewel_bright = [
                (235, 30, 44),    # #eb1e2c - Red
                (253, 111, 48),   # #fd6f30 - Orange
                (249, 167, 41),   # #f9a729 - Yellow-orange
                (249, 210, 60),   # #f9d23c - Yellow
                (95, 187, 104),   # #5fbb68 - Green
                (100, 205, 204),  # #64cdcc - Teal
                (145, 220, 234),  # #91dcea - Light blue
                (164, 164, 213),  # #a4a4d5 - Lavender
                (187, 201, 229)   # #bbc9e5 - Light purple
            ]
            return jewel_bright[:count]
            
        elif count <= 19:
            # Hue Circle palette (from Tableau)
            hue_circle = [
                (27, 163, 198),   # #1ba3c6
                (44, 181, 192),   # #2cb5c0
                (48, 188, 173),   # #30bcad
                (33, 176, 135),   # #21b087
                (51, 166, 92),    # #33a65c
                (87, 163, 55),    # #57a337
                (162, 182, 39),   # #a2b627
                (213, 187, 33),   # #d5bb21
                (248, 182, 32),   # #f8b620
                (248, 146, 23),   # #f89217
                (240, 103, 25),   # #f06719
                (224, 52, 38),    # #e03426
                (246, 73, 113),   # #f64971
                (252, 113, 158),  # #fc719e
                (235, 115, 179),  # #eb73b3
                (206, 105, 190),  # #ce69be
                (162, 109, 194),  # #a26dc2
                (120, 115, 192),  # #7873c0
                (79, 124, 186)    # #4f7cba
            ]
            return hue_circle[:count]
            
        else:
            # For large sets, evenly distribute around the color wheel
            # with Tableau-like saturation and brightness
            colors = []
            for i in range(count):
                # Distribute hues evenly around the color wheel
                h = float(i) / count
                
                # Set saturation and value to match Tableau's vibrant colors
                s = 0.85  # High saturation but not 100%
                v = 0.9   # High brightness but not 100%
                
                # Convert HSV to RGB
                h_i = int(h * 6)
                f = h * 6 - h_i
                
                # Calculate components with Tableau-like saturation and brightness
                p = int(255 * v * (1 - s))
                q = int(255 * v * (1 - s * f))
                t = int(255 * v * (1 - s * (1 - f)))
                v_scaled = int(255 * v)
                
                if h_i == 0:
                    colors.append((v_scaled, t, p))
                elif h_i == 1:
                    colors.append((q, v_scaled, p))
                elif h_i == 2:
                    colors.append((p, v_scaled, t))
                elif h_i == 3:
                    colors.append((p, q, v_scaled))
                elif h_i == 4:
                    colors.append((t, p, v_scaled))
                else:
                    colors.append((v_scaled, p, q))
            
            return colors

    def ValueCheckbox_Changed(self, sender, args):
        """Handle checkbox state changes in the values list"""
        if getattr(self, '_is_updating_selection', False):
            return
            
        try:
            self._is_updating_selection = True
# Skip if checkbox interaction is disabled
            if not self.showElementsCheckbox.IsChecked:
                return
            
            # If the clicked item is part of a multi-selection, sync other selected items
            clicked_item = sender.DataContext
            if clicked_item and hasattr(self, 'valuesListBox') and self.valuesListBox.SelectedItems.Count > 1:
                if self.valuesListBox.SelectedItems.Contains(clicked_item):
                    target_state = sender.IsChecked
                    
                    # Temporarily disable this handler to avoid recursion
                    original_handler = self.ValueCheckbox_Changed
                    self.ValueCheckbox_Changed = None
                    
                    try:
                        for item in self.valuesListBox.SelectedItems:
                            if item != clicked_item:
                                item.IsChecked = target_state
                        
                        # Refresh UI
                        self.valuesListBox.Items.Refresh()
                    finally:
                        self.ValueCheckbox_Changed = original_handler
                
            # Select elements based on all checked checkboxes
            self.select_checked_elements()
        except Exception as ex:
            self.logger.error("Checkbox change error: %s", str(ex))
        finally:
            self._is_updating_selection = False
    
    def select_checked_elements(self):
        """Select all elements that have their checkboxes checked"""
        try:
            if not self.showElementsCheckbox.IsChecked or not self.valuesListBox.Items.Count:
                return
                
            # Create element ID collection for selection
            element_ids = self.System.Collections.Generic.List[self.DB.ElementId]()
            checked_count = 0
            
            # Add elements from all checked items to the selection
            for value_item in self.valuesListBox.Items:
                if value_item.IsChecked:
                    checked_count += 1
                    for element_id in value_item.ele_id:
                        element_ids.Add(element_id)
            
            # Set selection in the active document
            self.uidoc.Selection.SetElementIds(element_ids)
            
            # Update status
            if checked_count > 0:
                self.statusText.Text = "Selected elements from {} checked values.".format(checked_count)
            else:
                self.statusText.Text = "No values checked. Select checkboxes to highlight elements."
                
        except Exception as ex:
            self.statusText.Text = "Error selecting elements: " + str(ex)
            self.logger.error("Selection error: %s", str(ex))

    def HeaderCheckBox_Changed(self, sender, args):
        """Handle header checkbox state changes to check/uncheck all values."""
        try:
            # Get the state of the header checkbox
            is_checked = sender.IsChecked
            
            # Update all value items
            if self.valuesListBox.Items:
                # Store original handler reference
                original_handler = self.ValueCheckbox_Changed
                
                # Temporarily set the handler to None to prevent it from firing
                self.ValueCheckbox_Changed = None
                
                # Update all checkboxes
                for value_item in self.valuesListBox.Items:
                    value_item.IsChecked = is_checked
                
                # Refresh the list view
                self.valuesListBox.Items.Refresh()

                # Restore the original handler
                self.ValueCheckbox_Changed = original_handler
                
                # Update selection in Revit if show elements is enabled
                if self.showElementsCheckbox.IsChecked:
                    self.select_checked_elements()
                
                # Update status
                action = "checked" if is_checked else "unchecked"
                self.statusText.Text = "All values {}.".format(action)
                
        except Exception as ex:
            self.statusText.Text = "Error updating checkboxes: " + str(ex)
            self.logger.error("Header checkbox error: %s", str(ex))

# Main script execution
if __name__ == "__main__":
    # Check if we have an active view
    active_view = get_active_view(doc)
    
    if active_view:
        # Start the UI
        try:
            colorizer_ui = RevitColorizerWindow()
            colorizer_ui.Show()
        except Exception as ex:
            logger.error("Error showing window: {}".format(str(ex)))
            raise
    else:
        UI.TaskDialog.Show(
            "PyRevit Colorizer", 
            "Please open a view where visibility settings can be modified."
        )