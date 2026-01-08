# -*- coding: utf-8 -*-
"""
Class Crawler
Sends element types from the active view to a classification service
and displays the results.
"""

__title__ = "Class Crawler"
__author__ = "pyByggstyrning"
__doc__ = """Send element types from active view to Byggstyrning's IFC classification service"""
__highlight__ = 'new'

import clr
import json
import os
import os.path as op
import re
import codecs
import time
from collections import OrderedDict

# .NET imports
clr.AddReference('System')
clr.AddReference('System.Net')
clr.AddReference('System.Web')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
import System
from System import Uri, Object, Action
from System.Net import WebClient, WebRequest, WebHeaderCollection
from System.Text import Encoding
from System.IO import StreamReader, File, Directory
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs, PropertyChangedEventHandler
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Window
import System.Windows
from System.Windows.Controls import DataGrid, ListBox, TextBox, ComboBox, Button, CheckBox
from System.Windows.Threading import DispatcherPriority
from System.Windows.Data import Binding, BindingMode

# Revit imports
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# PyRevit imports
import pyrevit
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms
from pyrevit.revit.db import query
from pyrevit.forms import WPFWindow

# Add lib path for styles
import sys
script_path = __file__
extension_dir = op.dirname(op.dirname(op.dirname(op.dirname(script_path))))
lib_path = op.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from styles import load_styles_to_window

# Get current document and logger
doc = revit.doc
logger = script.get_logger()
logger.set_quiet_mode()  # Suppress INFO messages

# Classification endpoint
CLASSIFICATION_URL = "https://n8n.byggstyrning.se/webhook/classification"

# IFC Schema URLs
IFC_SCHEMA_URLS = {
    "IFC2X3": "https://bonsaibim.org/assets/IFC2X3.json",
    "IFC4": "https://bonsaibim.org/assets/IFC4.json",
    "IFC4X3": "https://bonsaibim.org/assets/IFC4X3.json"
}

# Global cache for IFC schema data
_ifc_schema_cache = {}

def get_cache_dir():
    """Get the cache directory for storing IFC schema JSON files"""
    script_path = __file__
    extension_dir = op.dirname(op.dirname(op.dirname(op.dirname(script_path))))
    cache_dir = op.join(extension_dir, "cache", "ifc_schemas")
    if not op.exists(cache_dir):
        Directory.CreateDirectory(cache_dir)
    return cache_dir

def get_cache_path(version):
    """Get the cache file path for a specific IFC version"""
    cache_dir = get_cache_dir()
    return op.join(cache_dir, "{}.json".format(version))

def load_ifc_schema(version):
    """Load IFC schema JSON from cache or fetch from URL"""
    global _ifc_schema_cache
    
    # Check memory cache first
    if version in _ifc_schema_cache:
        return _ifc_schema_cache[version]
    
    cache_path = get_cache_path(version)
    
    # Try to load from disk cache
    if op.exists(cache_path):
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                schema_data = json.load(f)
                _ifc_schema_cache[version] = schema_data
                logger.info("Loaded IFC {} schema from cache".format(version))
                return schema_data
        except Exception as e:
            logger.warning("Error loading cached schema {}: {}".format(version, str(e)))
    
    # Fetch from URL
    if version not in IFC_SCHEMA_URLS:
        logger.error("Unknown IFC version: {}".format(version))
        return None
    
    url = IFC_SCHEMA_URLS[version]
    logger.info("Fetching IFC {} schema from {}".format(version, url))
    
    try:
        client = WebClient()
        client.Encoding = Encoding.UTF8
        response = client.DownloadString(url)
        schema_data = json.loads(response)
        
        # Cache to disk
        try:
            with open(cache_path, 'w', encoding='utf-8') as f:
                json.dump(schema_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning("Error caching schema {}: {}".format(version, str(e)))
        
        # Cache in memory
        _ifc_schema_cache[version] = schema_data
        logger.info("Loaded IFC {} schema from URL".format(version))
        return schema_data
        
    except Exception as e:
        logger.error("Error fetching IFC schema {}: {}".format(version, str(e)))
        return None
    finally:
        if 'client' in locals():
            client.Dispose()

def search_ifc_classes(query, schema_data):
    """Fuzzy search IFC classes in schema data"""
    if not schema_data:
        return []
    
    # Handle None or empty query - show all classes
    if not query:
        query = ""
    
    query_lower = query.lower().strip()
    if not query_lower:
        # Return all classes if query is empty
        results = []
        for class_name, class_data in schema_data.items():
            if isinstance(class_data, dict):
                description = class_data.get('description', '')
                predefined_types = class_data.get('predefined_types', {})
                spec_url = class_data.get('spec_url', '')
                results.append(IfcClassInfo(class_name, description, predefined_types, spec_url))
        return results[:100]  # Limit to 100 results
    
    results = []
    
    for class_name, class_data in schema_data.items():
        if not isinstance(class_data, dict):
            continue
        
        score = 0
        description = class_data.get('description', '')
        predefined_types = class_data.get('predefined_types', {})
        spec_url = class_data.get('spec_url', '')
        
        # Check class name match
        class_name_lower = class_name.lower()
        if query_lower in class_name_lower:
            score += 10
            # Boost exact matches
            if class_name_lower.startswith(query_lower):
                score += 5
        
        # Check description match
        if description:
            description_lower = description.lower()
            if query_lower in description_lower:
                score += 3
        
        # Check predefined types
        for type_name, type_desc in predefined_types.items():
            type_name_lower = type_name.lower()
            if query_lower in type_name_lower:
                score += 2
            if type_desc and query_lower in type_desc.lower():
                score += 1
        
        if score > 0:
            results.append((score, IfcClassInfo(class_name, description, predefined_types, spec_url)))
    
    # Sort by score (descending) and return top results
    results.sort(key=lambda x: x[0], reverse=True)
    return [item[1] for item in results[:50]]  # Return top 50 matches

class SelectableElementType(Object, INotifyPropertyChanged):
    """Wrapper class for element type with selection state for WPF binding"""
    
    def __init__(self, type_info):
        self._type_info = type_info
        self._is_selected = False
        self._is_processing = False
        self._is_queued = False
        self._propertyChanged = None  # Event handler storage
        
        # Load IFC values from Revit parameters
        element_type = type_info.element_type
        self._ifc_class = ""
        self._predefined_type = ""
        
        # Read IFC Class parameter
        ifc_param = element_type.LookupParameter("Export Type to IFC As")
        if ifc_param and ifc_param.HasValue:
            self._ifc_class = ifc_param.AsString() or ""
        
        # Read Predefined Type parameter
        predefined_param = element_type.LookupParameter("Type IFC Predefined Type")
        if not predefined_param:
            predefined_param = element_type.LookupParameter("Predefined Type")
        if predefined_param and predefined_param.HasValue:
            self._predefined_type = predefined_param.AsString() or ""
    
    # INotifyPropertyChanged implementation
    def add_PropertyChanged(self, handler):
        # Wrap non-delegate handlers to ensure type compatibility when combining
        if not isinstance(handler, PropertyChangedEventHandler):
            handler = PropertyChangedEventHandler(handler)
        if self._propertyChanged is None:
            self._propertyChanged = handler
        else:
            self._propertyChanged += handler
    
    def remove_PropertyChanged(self, handler):
        if not isinstance(handler, PropertyChangedEventHandler):
            handler = PropertyChangedEventHandler(handler)
        if self._propertyChanged is not None:
            self._propertyChanged -= handler
    
    def OnPropertyChanged(self, property_name):
        if self._propertyChanged is not None:
            self._propertyChanged(self, PropertyChangedEventArgs(property_name))
    
    @property
    def IsSelected(self):
        return self._is_selected
    
    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self.OnPropertyChanged("IsSelected")
    
    @property
    def Category(self):
        return self._type_info.category
    
    @property
    def Family(self):
        return self._type_info.family
    
    @property
    def TypeName(self):
        return self._type_info.type_name
    
    @property
    def Manufacturer(self):
        return self._type_info.manufacturer
    
    @property
    def ManufacturerDisplay(self):
        if self._type_info.manufacturer:
            return "Manufacturer: {}".format(self._type_info.manufacturer)
        return ""
    
    @property
    def HasManufacturer(self):
        return bool(self._type_info.manufacturer)
    
    @property
    def type_info(self):
        return self._type_info
    
    @property
    def IfcClass(self):
        return self._ifc_class
    
    @IfcClass.setter
    def IfcClass(self, value):
        if self._ifc_class != value:
            self._ifc_class = value or ""
            self.OnPropertyChanged("IfcClass")
    
    @property
    def PredefinedType(self):
        return self._predefined_type
    
    @PredefinedType.setter
    def PredefinedType(self, value):
        if self._predefined_type != value:
            self._predefined_type = value or ""
            self.OnPropertyChanged("PredefinedType")
    
    @property
    def IsProcessing(self):
        return self._is_processing
    
    @IsProcessing.setter
    def IsProcessing(self, value):
        if self._is_processing != value:
            self._is_processing = value
            self.OnPropertyChanged("IsProcessing")
    
    @property
    def IsQueued(self):
        return self._is_queued
    
    @IsQueued.setter
    def IsQueued(self, value):
        if self._is_queued != value:
            self._is_queued = value
            self.OnPropertyChanged("IsQueued")
    
    def matches_search(self, search_text):
        """Check if this item matches the search text"""
        if not search_text:
            return True
        search_lower = search_text.lower()
        return (
            search_lower in self.Category.lower() or
            search_lower in self.Family.lower() or
            search_lower in self.TypeName.lower() or
            (self.Manufacturer and search_lower in self.Manufacturer.lower()) or
            search_lower in self.IfcClass.lower() or
            search_lower in self.PredefinedType.lower()
        )


class ElementTypeInfo(object):
    """Container for element type information"""
    
    def __init__(self, element_type):
        self.element_type = element_type
        self.category = self._get_category_name(element_type)
        self.family = self._get_family_name(element_type)
        self.type_name = self._get_type_name(element_type)
        self.manufacturer = self._get_parameter_value(element_type, "Manufacturer")
        
    def _get_category_name(self, element_type):
        """Get the category name"""
        try:
            if element_type.Category:
                return element_type.Category.Name
        except:
            pass
        return "Unknown"
    
    def _get_family_name(self, element_type):
        """Get the family name"""
        try:
            if hasattr(element_type, 'FamilyName'):
                return element_type.FamilyName
            elif hasattr(element_type, 'Family') and element_type.Family:
                return element_type.Family.Name
        except:
            pass
        return "Unknown"
    
    def _get_type_name(self, element_type):
        """Get the type name safely using PyRevit's method"""
        try:
            # Use PyRevit's get_name function which handles different Revit versions
            return query.get_name(element_type)
        except:
            try:
                # Fallback to direct access
                return element_type.Name
            except:
                return "Unknown Type"
    
    def _get_parameter_value(self, element, param_name):
        """Get parameter value by name"""
        try:
            param = element.LookupParameter(param_name)
            if param and param.HasValue:
                if param.StorageType == StorageType.String:
                    return param.AsString() or ""
                elif param.StorageType == StorageType.Integer:
                    return str(param.AsInteger())
                elif param.StorageType == StorageType.Double:
                    return str(param.AsDouble())
        except:
            pass
        return ""
    
    def to_dict(self):
        """Convert to dictionary for API request"""
        return {
            "Category": self.category,
            "Family": self.family,
            "Type": self.type_name,
            "Manufacturer": self.manufacturer
        }
    
    def __str__(self):
        """String representation for the selection list"""
        return "{} - {} - {}".format(self.category, self.family, self.type_name)

class ClassificationItem(Object, INotifyPropertyChanged):
    """Data class for classification results with INotifyPropertyChanged for WPF binding"""
    
    def __init__(self, type_info, ifc_class, predefined_type, reasoning):
        self._is_approved = True
        self._category = type_info.category
        self._family = type_info.family
        self._type_name = type_info.type_name
        self._ifc_class = ifc_class or ""
        self._predefined_type = predefined_type or ""
        self._reasoning = reasoning or ""
        self.element_type = type_info.element_type
        self.type_info = type_info
        self._propertyChanged = None  # Event handler storage
    
    # INotifyPropertyChanged implementation
    def add_PropertyChanged(self, handler):
        # Wrap non-delegate handlers to ensure type compatibility when combining
        if not isinstance(handler, PropertyChangedEventHandler):
            handler = PropertyChangedEventHandler(handler)
        if self._propertyChanged is None:
            self._propertyChanged = handler
        else:
            self._propertyChanged += handler
    
    def remove_PropertyChanged(self, handler):
        if not isinstance(handler, PropertyChangedEventHandler):
            handler = PropertyChangedEventHandler(handler)
        if self._propertyChanged is not None:
            self._propertyChanged -= handler
    
    def OnPropertyChanged(self, property_name):
        if self._propertyChanged is not None:
            self._propertyChanged(self, PropertyChangedEventArgs(property_name))
    
    @property
    def IsApproved(self):
        return self._is_approved
    
    @IsApproved.setter
    def IsApproved(self, value):
        if self._is_approved != value:
            self._is_approved = value
            self.OnPropertyChanged("IsApproved")
    
    @property
    def Category(self):
        return self._category
    
    @property
    def Family(self):
        return self._family
    
    @property
    def TypeName(self):
        return self._type_name
    
    @property
    def IfcClass(self):
        return self._ifc_class
    
    @IfcClass.setter
    def IfcClass(self, value):
        if self._ifc_class != value:
            self._ifc_class = value or ""
            self.OnPropertyChanged("IfcClass")
    
    @property
    def PredefinedType(self):
        return self._predefined_type
    
    @PredefinedType.setter
    def PredefinedType(self, value):
        if self._predefined_type != value:
            self._predefined_type = value or ""
            self.OnPropertyChanged("PredefinedType")
    
    @property
    def Reasoning(self):
        return self._reasoning

class IfcClassInfo(object):
    """Data class for IFC schema browser results"""
    
    def __init__(self, name, description, predefined_types, spec_url):
        self.name = name
        self.description = description or ""
        self.predefined_types = predefined_types or {}
        self.spec_url = spec_url or ""
    
    @property
    def Name(self):
        """Property for WPF binding (capitalized)"""
        return self.name
    
    @property
    def Description(self):
        """Property for WPF binding (capitalized)"""
        return self.description
    
    @property
    def PredefinedTypesDisplay(self):
        """Get formatted predefined types string"""
        if not self.predefined_types:
            return "None"
        types_list = list(self.predefined_types.keys())[:5]  # Show first 5
        result = ", ".join(types_list)
        if len(self.predefined_types) > 5:
            result += "..."
        return result


class ClassificationWindow(WPFWindow):
    """Single-tab WPF Window for IFC classification with editable IFC Class and Predefined Type columns"""
    
    def __init__(self, element_types):
        try:
            # #region agent log
            import codecs, json, time
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1","location":"script.py:520","message":"ClassificationWindow.__init__ entry","data":{"element_types_count":len(element_types)},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            xaml_path = op.join(op.dirname(__file__), "ClassificationWindow.xaml")
            logger.info("Loading ClassificationWindow from: {}".format(xaml_path))
            
            if not op.exists(xaml_path):
                raise Exception("XAML file not found: {}".format(xaml_path))
            
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1,H5","location":"script.py:528","message":"Before WPFWindow.__init__","data":{"xaml_exists":True},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            WPFWindow.__init__(self, xaml_path)
            
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1,H5","location":"script.py:530","message":"After WPFWindow.__init__","data":{"success":True},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            logger.info("WPFWindow initialized successfully")
            
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H2","location":"script.py:532","message":"Before load_styles_to_window","data":{},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            # Load styles
            load_styles_to_window(self)
            
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H2","location":"script.py:533","message":"After load_styles_to_window","data":{},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            logger.info("Styles loaded")
            
            # Store element types
            self.element_types = element_types
            
            # State
            self._is_updating_selection = False
            self.all_items = []
            for type_info in element_types:
                self.all_items.append(SelectableElementType(type_info))
            self.all_items.sort(key=lambda x: (x.Category, x.Family, x.TypeName))
            
            # Sort state
            self._sort_column = None
            self._sort_direction = None  # 'asc' or 'desc'
            
            # Handle window closing
            self.Closing += self.Window_Closing
            
            # Handle window key events - prevent ESC from closing
            self.PreviewKeyDown += self.Window_PreviewKeyDown
            
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1","location":"script.py:549","message":"Before _initialize_ui","data":{},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            # Initialize UI
            self._initialize_ui()
            
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1","location":"script.py:550","message":"After _initialize_ui","data":{},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            logger.info("ClassificationWindow initialization complete")
        except Exception as ex:
            logger.error("Error initializing ClassificationWindow: {}".format(str(ex)))
            import traceback
            logger.error(traceback.format_exc())
            forms.alert("Error initializing window: {}\n\nCheck the log for details.".format(str(ex)), title="Initialization Error")
            raise
    
    def _initialize_ui(self):
        """Initialize UI components"""
        try:
            from System.Collections.ObjectModel import ObservableCollection
            self.displayed_items = ObservableCollection[SelectableElementType]()
            
            for item in self.all_items:
                self.displayed_items.Add(item)
            
            self.elementTypesDataGrid.ItemsSource = self.displayed_items
            
            # #region agent log
            # Check column configuration after setting ItemsSource
            for col in self.elementTypesDataGrid.Columns:
                col_header = str(col.Header) if col.Header else None
                if col_header == "IFC Class":
                    from System.Windows.Controls import DataGridTemplateColumn
                    if isinstance(col, DataGridTemplateColumn):
                        with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                            f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H5","location":"script.py:_initialize_ui","message":"IFC Class column configuration","data":{"is_read_only":col.IsReadOnly,"has_cell_template":col.CellTemplate is not None,"has_cell_editing_template":col.CellEditingTemplate is not None},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            # IFC search state
            self.current_schema_version = "IFC2X3"
            self.current_schema_data = None
            self._selected_ifc_class = None
            self._selected_predefined_type = None
            self._target_item_for_ifc = None
            
            # Set up event handlers - wrap each in try-except to identify which one fails
            try:
                self.searchTextBox.TextChanged += self.SearchTextBox_TextChanged
            except Exception as ex:
                logger.error("Error binding searchTextBox: {}".format(str(ex)))
                raise
            
            try:
                self.selectAllButton.Click += self.SelectAllButton_Click
                self.deselectAllButton.Click += self.DeselectAllButton_Click
            except Exception as ex:
                logger.error("Error binding select/deselect buttons: {}".format(str(ex)))
                raise
            
            # IFC Search Panel - bind event handlers
            try:
                self.ifcSearchTextBox.TextChanged += self.IfcSearchTextBox_TextChanged
            except Exception as ex:
                logger.error("Error binding ifcSearchTextBox: {}".format(str(ex)))
                raise
            
            try:
                self.schemaVersionComboBox.SelectionChanged += self.SchemaVersionComboBox_SelectionChanged
            except Exception as ex:
                logger.error("Error binding schemaVersionComboBox: {}".format(str(ex)))
                raise
            
            try:
                self.ifcReferenceButton.Click += self.IfcReferenceButton_Click
            except Exception as ex:
                logger.error("Error binding ifcReferenceButton: {}".format(str(ex)))
                raise
            
            # Create IFC results ListBox programmatically (to avoid XAML parsing issues)
            try:
                from System.Collections.ObjectModel import ObservableCollection
                self.ifc_results = ObservableCollection[IfcClassInfo]()
                self._create_ifc_results_listbox()
            except Exception as ex:
                logger.error("Error creating IFC results ListBox: {}".format(str(ex)))
                raise
            
            # Set ComboBox to IFC2X3 and load IFC schema by default
            try:
                # Ensure ComboBox is set to IFC2X3
                if self.schemaVersionComboBox.Items.Count > 0:
                    self.schemaVersionComboBox.SelectedIndex = 0
                # Load IFC2X3 schema by default
                self._load_ifc_schema("IFC2X3")
            except Exception as ex:
                logger.error("Error loading default IFC schema: {}".format(str(ex)))
                # Don't raise - allow window to open even if schema fails to load
            
            # Apply button removed - IFC class is applied automatically on selection
            
            try:
                self.getAiSuggestionsButton.Click += self.GetAiSuggestionsButton_Click
            except Exception as ex:
                logger.error("Error binding getAiSuggestionsButton: {}".format(str(ex)))
                raise
            
            try:
                self.updateButton.Click += self.UpdateButton_Click
            except Exception as ex:
                logger.error("Error binding updateButton: {}".format(str(ex)))
                raise
            
            # Handle row clicks
            self.elementTypesDataGrid.PreviewMouseLeftButtonUp += self.DataGrid_PreviewMouseLeftButtonUp
            self.elementTypesDataGrid.PreviewKeyDown += self.DataGrid_PreviewKeyDown
            
            # #region agent log
            # Add DataGrid editing event handlers for debugging
            self.elementTypesDataGrid.BeginningEdit += self.DataGrid_BeginningEdit
            self.elementTypesDataGrid.PreparingCellForEdit += self.DataGrid_PreparingCellForEdit
            self.elementTypesDataGrid.CellEditEnding += self.DataGrid_CellEditEnding
            # #endregion
            
            # Status update timer
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan
            self._status_timer = DispatcherTimer()
            self._status_timer.Interval = TimeSpan.FromMilliseconds(200)
            self._status_timer.Tick += lambda s, e: self._update_status()
            self._status_timer.Start()
            
            self._update_status()
        except Exception as ex:
            logger.error("Error in _initialize_ui: {}".format(str(ex)))
            import traceback
            logger.error(traceback.format_exc())
            raise
    
    def SortAscButton_Click(self, sender, e):
        """Handle ascending sort button click"""
        column_name = sender.Tag
        if column_name:
            self._sort_column = column_name
            self._sort_direction = 'asc'
            self._apply_sort()
    
    def SortDescButton_Click(self, sender, e):
        """Handle descending sort button click"""
        column_name = sender.Tag
        if column_name:
            self._sort_column = column_name
            self._sort_direction = 'desc'
            self._apply_sort()
    
    def _apply_sort(self):
        """Apply sorting to displayed_items based on current sort state"""
        if not self._sort_column or not self._sort_direction:
            return
        
        try:
            # Get the property name for sorting
            property_name = self._sort_column
            
            # Create a list from the ObservableCollection
            items_list = list(self.displayed_items)
            
            # Define sort key function
            def get_sort_key(item):
                value = getattr(item, property_name, None)
                if value is None:
                    return ""
                # Handle string comparison (case-insensitive)
                if isinstance(value, str):
                    return value.lower()
                return value
            
            # Sort the list
            reverse = (self._sort_direction == 'desc')
            items_list.sort(key=get_sort_key, reverse=reverse)
            
            # Clear and repopulate the ObservableCollection
            self.displayed_items.Clear()
            for item in items_list:
                self.displayed_items.Add(item)
            
            # Refresh the DataGrid
            self.elementTypesDataGrid.Items.Refresh()
            
        except Exception as ex:
            logger.error("Error applying sort: {}".format(str(ex)))
            import traceback
            logger.error(traceback.format_exc())
    
    def Window_PreviewKeyDown(self, sender, e):
        """Handle window key events - prevent ESC from closing the window"""
        from System.Windows.Input import Key
        if e.Key == Key.Escape:
            # If editing a cell, cancel the edit instead of closing the window
            if self.elementTypesDataGrid.CurrentCell is not None and self.elementTypesDataGrid.CurrentCell.IsEditing:
                self.elementTypesDataGrid.CancelEdit()
                e.Handled = True
            else:
                # Don't close the window on ESC - just mark as handled
                e.Handled = True
    
    def Window_Closing(self, sender, e):
        """Handle window closing - ensure proper cleanup"""
        # Stop any running timers
        if hasattr(self, '_status_timer') and self._status_timer is not None:
            try:
                self._status_timer.Stop()
            except:
                pass
    
    def _update_status(self):
        """Update the status text"""
        selected_count = sum(1 for item in self.all_items if item.IsSelected)
        total_count = len(self.all_items)
        displayed_count = self.displayed_items.Count
        
        if displayed_count == total_count:
            self.statusTextBlock.Text = "{} of {} selected".format(selected_count, total_count)
        else:
            self.statusTextBlock.Text = "{} of {} selected (showing {} filtered)".format(
                selected_count, total_count, displayed_count)
        
        # Enable/disable buttons
        self.getAiSuggestionsButton.IsEnabled = selected_count > 0
        self.updateButton.IsEnabled = selected_count > 0
    
    def _filter_items(self):
        """Filter items based on search text"""
        search_text = self.searchTextBox.Text or ""
        self.displayed_items.Clear()
        
        for item in self.all_items:
            if item.matches_search(search_text):
                self.displayed_items.Add(item)
        
        # Re-apply sorting if active
        if self._sort_column and self._sort_direction:
            self._apply_sort()
        
        self._update_status()
    
    def SearchTextBox_TextChanged(self, sender, e):
        """Handle search text changed"""
        self._filter_items()
    
    def SelectAllButton_Click(self, sender, e):
        """Select all visible items"""
        for item in self.displayed_items:
            item.IsSelected = True
        self._update_status()
    
    def DeselectAllButton_Click(self, sender, e):
        """Deselect all visible items"""
        for item in self.displayed_items:
            item.IsSelected = False
        self._update_status()
    
    def TextBox_Loaded(self, sender, e):
        """Focus and select text when TextBox is loaded in edit mode"""
        try:
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H2","location":"script.py:TextBox_Loaded","message":"TextBox_Loaded called","data":{"sender_type":str(type(sender).__name__) if sender else None},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            text_box = sender
            if text_box is not None:
                # #region agent log
                with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                    f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H2","location":"script.py:TextBox_Loaded","message":"TextBox properties before focus","data":{"is_visible":text_box.IsVisible,"opacity":text_box.Opacity,"background":str(text_box.Background) if text_box.Background else None,"foreground":str(text_box.Foreground) if text_box.Foreground else None,"text":text_box.Text[:50] if text_box.Text else None,"width":text_box.ActualWidth,"height":text_box.ActualHeight},"timestamp":int(time.time()*1000)})+'\n')
                # #endregion
                
                text_box.Focus()
                text_box.SelectAll()
                
                # #region agent log
                with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                    f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H2","location":"script.py:TextBox_Loaded","message":"TextBox focused and selected","data":{"is_focused":text_box.IsFocused,"selection_length":text_box.SelectionLength},"timestamp":int(time.time()*1000)})+'\n')
                # #endregion
        except Exception as ex:
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H2","location":"script.py:TextBox_Loaded","message":"TextBox_Loaded error","data":{"error":str(ex)},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            logger.debug("TextBox_Loaded error: {}".format(str(ex)))
    
    def DataGrid_PreviewMouseLeftButtonUp(self, sender, args):
        """Handle mouse up to toggle checkboxes for multi-selected rows"""
        if self._is_updating_selection:
            return
        
        try:
            from System.Windows.Media import VisualTreeHelper
            
            point = args.GetPosition(self.elementTypesDataGrid)
            hit = VisualTreeHelper.HitTest(self.elementTypesDataGrid, point)
            
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1","location":"script.py:PreviewMouseLeftButtonUp","message":"PreviewMouseLeftButtonUp called","data":{"point_x":point.X,"point_y":point.Y,"has_hit":hit is not None},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            if not hit or not hit.VisualHit:
                return
            
            parent = hit.VisualHit
            clicked_on_checkbox = False
            clicked_element_type = None
            while parent and parent != self.elementTypesDataGrid:
                type_name = parent.GetType().Name
                # #region agent log
                with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                    f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1","location":"script.py:PreviewMouseLeftButtonUp","message":"Checking parent element","data":{"type_name":type_name},"timestamp":int(time.time()*1000)})+'\n')
                # #endregion
                if "CheckBox" in type_name:
                    clicked_on_checkbox = True
                    break
                if "DataGridCell" in type_name or "TextBlock" in type_name:
                    clicked_element_type = type_name
                parent = VisualTreeHelper.GetParent(parent)
            
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1","location":"script.py:PreviewMouseLeftButtonUp","message":"Click analysis","data":{"clicked_on_checkbox":clicked_on_checkbox,"clicked_element_type":clicked_element_type},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            # Only toggle IsSelected if checkbox was clicked
            if clicked_on_checkbox:
                selected_items = list(self.elementTypesDataGrid.SelectedItems)
                if len(selected_items) > 1:
                    parent = hit.VisualHit
                    clicked_item = None
                    while parent and parent != self.elementTypesDataGrid:
                        if hasattr(parent, 'DataContext') and parent.DataContext in selected_items:
                            clicked_item = parent.DataContext
                            break
                        parent = VisualTreeHelper.GetParent(parent)
                    
                    if clicked_item:
                        target_state = not clicked_item.IsSelected
                        
                        def apply_to_others():
                            self._is_updating_selection = True
                            try:
                                for item in selected_items:
                                    if item != clicked_item:
                                        item.IsSelected = target_state
                                self._update_status()
                            finally:
                                self._is_updating_selection = False
                        
                        self.Dispatcher.BeginInvoke(
                            DispatcherPriority.Input,
                            Action(apply_to_others)
                        )
                # If checkbox was clicked, don't do anything else - let the checkbox handle it
                # #region agent log
                with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                    f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1","location":"script.py:PreviewMouseLeftButtonUp","message":"Returning early - checkbox clicked","data":{},"timestamp":int(time.time()*1000)})+'\n')
                # #endregion
                return
            
            # If checkbox was NOT clicked, do not toggle IsSelected
            # The row selection is handled by the DataGrid itself
            # This prevents clicking the row from affecting the checkbox state
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1","location":"script.py:PreviewMouseLeftButtonUp","message":"Not a checkbox click - allowing event to propagate","data":{},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
                    
        except Exception as ex:
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H1","location":"script.py:PreviewMouseLeftButtonUp","message":"Error in PreviewMouseLeftButtonUp","data":{"error":str(ex)},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            logger.debug("DataGrid mouse up error: {}".format(str(ex)))
    
    def DataGrid_BeginningEdit(self, sender, e):
        """Handle when DataGrid begins editing"""
        try:
            column_header = str(e.Column.Header) if e.Column else None
            # #region agent log
            from System.Windows.Controls import DataGridTemplateColumn
            col_info = {}
            if e.Column and isinstance(e.Column, DataGridTemplateColumn):
                col_info = {
                    "is_read_only": e.Column.IsReadOnly,
                    "has_cell_template": e.Column.CellTemplate is not None,
                    "has_cell_editing_template": e.Column.CellEditingTemplate is not None
                }
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:BeginningEdit","message":"DataGrid beginning edit","data":{"column_header":column_header,"row_index":e.Row.GetIndex(),"column_info":col_info},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            # Workaround: Since PreparingCellForEdit doesn't fire reliably for editable columns,
            # manually find and focus the TextBox after a short delay
            if column_header == "IFC Class" or column_header == "Predefined Type":
                # #region agent log
                with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                    f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:BeginningEdit","message":"Starting TextBox search workaround","data":{},"timestamp":int(time.time()*1000)})+'\n')
                # #endregion
                
                def find_and_focus_textbox():
                    try:
                        # #region agent log
                        with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                            f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:BeginningEdit:find_and_focus_textbox","message":"Callback executing","data":{},"timestamp":int(time.time()*1000)})+'\n')
                        # #endregion
                        
                        from System.Windows.Media import VisualTreeHelper
                        from System.Windows.Controls import DataGridCell
                        
                        def find_textbox(element):
                            if isinstance(element, TextBox):
                                return element
                            for i in range(VisualTreeHelper.GetChildrenCount(element)):
                                child = VisualTreeHelper.GetChild(element, i)
                                result = find_textbox(child)
                                if result:
                                    return result
                            return None
                        
                        # Find the cell from the row
                        row = e.Row
                        if row:
                            # #region agent log
                            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:BeginningEdit:find_and_focus_textbox","message":"Searching row for cell","data":{"children_count":VisualTreeHelper.GetChildrenCount(row)},"timestamp":int(time.time()*1000)})+'\n')
                            # #endregion
                            
                            # Search for the DataGridCell that matches this column
                            for i in range(VisualTreeHelper.GetChildrenCount(row)):
                                child = VisualTreeHelper.GetChild(row, i)
                                if isinstance(child, DataGridCell):
                                    # #region agent log
                                    with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                                        f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:BeginningEdit:find_and_focus_textbox","message":"Found DataGridCell","data":{"column_match":child.Column == e.Column,"is_editing":child.IsEditing},"timestamp":int(time.time()*1000)})+'\n')
                                    # #endregion
                                    
                                    if child.Column == e.Column:
                                        # Found the cell, now search for TextBox
                                        text_box = find_textbox(child)
                                        if text_box:
                                            # #region agent log
                                            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                                                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:BeginningEdit:find_and_focus_textbox","message":"Found TextBox manually","data":{"is_visible":text_box.IsVisible,"text":text_box.Text[:50] if text_box.Text else None},"timestamp":int(time.time()*1000)})+'\n')
                                            # #endregion
                                            text_box.Focus()
                                            text_box.SelectAll()
                                            return
                                        else:
                                            # #region agent log
                                            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                                                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:BeginningEdit:find_and_focus_textbox","message":"TextBox not found in cell","data":{},"timestamp":int(time.time()*1000)})+'\n')
                                            # #endregion
                    except Exception as ex:
                        # #region agent log
                        with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                            f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:BeginningEdit:find_and_focus_textbox","message":"Error finding TextBox","data":{"error":str(ex)},"timestamp":int(time.time()*1000)})+'\n')
                        # #endregion
                        import traceback
                        logger.error("Error in find_and_focus_textbox: {}".format(traceback.format_exc()))
                
                # Use Dispatcher to delay the search slightly to allow template to instantiate
                # Try multiple priority levels to ensure it executes
                self.Dispatcher.BeginInvoke(DispatcherPriority.Loaded, Action(find_and_focus_textbox))
                # Also try with Input priority as backup
                self.Dispatcher.BeginInvoke(DispatcherPriority.Input, Action(find_and_focus_textbox))
        except Exception as ex:
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:BeginningEdit","message":"BeginningEdit error","data":{"error":str(ex)},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
    
    def DataGrid_PreparingCellForEdit(self, sender, e):
        """Handle when DataGrid prepares cell for editing"""
        try:
            # #region agent log
            editing_element_type = None
            if e.EditingElement:
                editing_element_type = str(type(e.EditingElement).__name__)
            column_header = str(e.Column.Header) if e.Column else None
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:PreparingCellForEdit","message":"DataGrid preparing cell for edit","data":{"column_header":column_header,"editing_element_type":editing_element_type,"has_editing_element":e.EditingElement is not None},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
            
            # For IFC Class and Predefined Type columns, try to find TextBox in the visual tree
            if column_header == "IFC Class" or column_header == "Predefined Type":
                # #region agent log
                from System.Windows.Media import VisualTreeHelper
                def find_textbox_in_tree(element, depth=0, path=""):
                    if depth > 10:  # Prevent infinite recursion
                        return None
                    if isinstance(element, TextBox):
                        # #region agent log
                        with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                            f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H4","location":"script.py:PreparingCellForEdit:find_textbox","message":"Found TextBox!","data":{"depth":depth,"path":path},"timestamp":int(time.time()*1000)})+'\n')
                        # #endregion
                        return element
                    try:
                        child_count = VisualTreeHelper.GetChildrenCount(element)
                        for i in range(child_count):
                            child = VisualTreeHelper.GetChild(element, i)
                            child_type = str(type(child).__name__)
                            result = find_textbox_in_tree(child, depth + 1, path + "/" + child_type)
                            if result:
                                return result
                    except Exception as ex:
                        # #region agent log
                        with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                            f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H4","location":"script.py:PreparingCellForEdit:find_textbox","message":"Error in search","data":{"depth":depth,"error":str(ex)},"timestamp":int(time.time()*1000)})+'\n')
                        # #endregion
                    return None
                
                text_box = None
                if e.EditingElement:
                    # #region agent log
                    with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                        f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H4","location":"script.py:PreparingCellForEdit","message":"Starting TextBox search","data":{"editing_element_type":editing_element_type,"has_editing_element":e.EditingElement is not None},"timestamp":int(time.time()*1000)})+'\n')
                    # #endregion
                    text_box = find_textbox_in_tree(e.EditingElement, 0, editing_element_type)
                
                with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                    f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H4","location":"script.py:PreparingCellForEdit","message":"{} - searching for TextBox".format(column_header),"data":{"found_textbox":text_box is not None,"editing_element_type":editing_element_type},"timestamp":int(time.time()*1000)})+'\n')
                # #endregion
                
                if text_box:
                    # #region agent log
                    with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                        f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H4","location":"script.py:PreparingCellForEdit","message":"{} TextBox properties".format(column_header),"data":{"is_visible":text_box.IsVisible,"opacity":text_box.Opacity,"background":str(text_box.Background) if text_box.Background else None,"foreground":str(text_box.Foreground) if text_box.Foreground else None,"text":text_box.Text[:50] if text_box.Text else None,"width":text_box.ActualWidth,"height":text_box.ActualHeight,"is_focusable":text_box.Focusable,"is_enabled":text_box.IsEnabled},"timestamp":int(time.time()*1000)})+'\n')
                    # #endregion
                    # Focus and select text immediately
                    try:
                        text_box.Focus()
                        text_box.SelectAll()
                    except Exception as ex:
                        # #region agent log
                        with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                            f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H4","location":"script.py:PreparingCellForEdit","message":"Error focusing TextBox","data":{"error":str(ex)},"timestamp":int(time.time()*1000)})+'\n')
                        # #endregion
                        # Try again with dispatcher as fallback
                        def focus_textbox():
                            try:
                                text_box.Focus()
                                text_box.SelectAll()
                            except:
                                pass
                        self.Dispatcher.BeginInvoke(DispatcherPriority.Input, Action(focus_textbox))
            # Check if it's a TextBox and log its properties
            elif e.EditingElement and isinstance(e.EditingElement, TextBox):
                text_box = e.EditingElement
                # #region agent log
                with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                    f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H4","location":"script.py:PreparingCellForEdit","message":"TextBox found in PreparingCellForEdit","data":{"is_visible":text_box.IsVisible,"opacity":text_box.Opacity,"background":str(text_box.Background) if text_box.Background else None,"foreground":str(text_box.Foreground) if text_box.Foreground else None,"text":text_box.Text[:50] if text_box.Text else None,"width":text_box.ActualWidth,"height":text_box.ActualHeight,"is_focusable":text_box.Focusable},"timestamp":int(time.time()*1000)})+'\n')
                # #endregion
        except Exception as ex:
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:PreparingCellForEdit","message":"PreparingCellForEdit error","data":{"error":str(ex)},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
    
    def DataGrid_CellEditEnding(self, sender, e):
        """Handle when DataGrid cell edit ends"""
        try:
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:CellEditEnding","message":"DataGrid cell edit ending","data":{"column_header":str(e.Column.Header) if e.Column else None,"edit_action":str(e.EditAction)},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
        except Exception as ex:
            # #region agent log
            with codecs.open(r'c:\code\pyRevit Extensions\pyByggstyrning.extension\.cursor\debug.log', 'a', 'utf-8') as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"initial","hypothesisId":"H3","location":"script.py:CellEditEnding","message":"CellEditEnding error","data":{"error":str(ex)},"timestamp":int(time.time()*1000)})+'\n')
            # #endregion
    
    def DataGrid_PreviewKeyDown(self, sender, args):
        """Handle keyboard shortcuts for toggling selection"""
        from System.Windows.Input import Key
        
        try:
            if args.Key == Key.Space:
                self._toggle_selected_rows()
                args.Handled = True
                return
            
            if args.Key == Key.A and (args.KeyboardDevice.Modifiers & 4):
                for item in self.displayed_items:
                    item.IsSelected = True
                self._update_status()
                args.Handled = True
                return
                
        except Exception as ex:
            logger.debug("DataGrid key down error: {}".format(str(ex)))
    
    def _toggle_selected_rows(self, target_state=None):
        """Toggle IsSelected for all grid-selected rows"""
        if self._is_updating_selection:
            return
        
        selected_items = list(self.elementTypesDataGrid.SelectedItems)
        if not selected_items:
            return
        
        self._is_updating_selection = True
        try:
            if target_state is None:
                target_state = not selected_items[0].IsSelected
            
            for item in selected_items:
                item.IsSelected = target_state
            
            self._update_status()
        finally:
            self._is_updating_selection = False
    
    def GetAiSuggestionsButton_Click(self, sender, e):
        """Get AI suggestions for selected element types and update IFC columns"""
        try:
            # Collect selected items
            selected_items = [item for item in self.all_items if item.IsSelected]
            
            if not selected_items:
                forms.alert("Please select at least one element type to get AI suggestions.", title="No Selection")
                return
            
            # Initialize classification state
            self._cancel_requested = False
            self._current_index = 0
            self._total_count = len(selected_items)
            self._selected_items_for_classification = selected_items
            
            # Mark all selected items as queued
            for item in selected_items:
                item.IsQueued = True
                item.IsProcessing = False
            
            # Disable button during processing
            self.getAiSuggestionsButton.IsEnabled = False
            
            # Start processing
            self._process_next_classification()
            
        except Exception as ex:
            logger.error("Error in GetAiSuggestionsButton_Click: {}".format(str(ex)))
            forms.alert("An error occurred when getting AI suggestions: {}".format(str(ex)), title="Error")
            self.getAiSuggestionsButton.IsEnabled = True
    
    def _process_next_classification(self):
        """Process the next item in the classification queue"""
        if self._cancel_requested:
            self._finish_classification()
            return
        
        if self._current_index >= self._total_count:
            logger.info("All items processed.")
            self._finish_classification()
            return
        
        item = self._selected_items_for_classification[self._current_index]
        type_info = item.type_info
        
        # Set processing state to show spinner (remove from queue, mark as processing)
        item.IsQueued = False
        item.IsProcessing = True
        
        def do_classification():
            try:
                type_data = type_info.to_dict()
                response = send_classification_request(type_data)
                
                def update_ui():
                    try:
                        # Clear processing state
                        item.IsProcessing = False
                        
                        if response:
                            # Extract classification data
                            ifc_class, predefined_type, reasoning = extract_classification_data(response)
                            
                            # Update the item's IFC properties
                            if ifc_class and ifc_class != 'Not classified' and ifc_class != 'Error':
                                # Add "Type" suffix if needed
                                if ifc_class and not ifc_class.endswith("Type") and not ifc_class.endswith("StandardCase"):
                                    ifc_class += "Type"
                                item.IfcClass = ifc_class
                            
                            if predefined_type and predefined_type != 'Not specified' and predefined_type != 'Error':
                                item.PredefinedType = predefined_type
                            
                            logger.debug("Updated {}: IFC={}, Predefined={}".format(
                                type_info.type_name, item.IfcClass, item.PredefinedType))
                        else:
                            logger.warning("Failed to classify: {}".format(str(type_info)))
                        
                        self._current_index += 1
                        self._process_next_classification()
                    except Exception as ex:
                        logger.error("Error updating UI after classification: {}".format(str(ex)))
                        item.IsProcessing = False
                        item.IsQueued = False
                        self._current_index += 1
                        self._process_next_classification()
                
                self.Dispatcher.BeginInvoke(
                    DispatcherPriority.Normal,
                    Action(update_ui)
                )
                
            except Exception as ex:
                logger.error("Classification error: {}".format(str(ex)))
                def handle_error():
                    try:
                        item.IsProcessing = False
                        item.IsQueued = False
                        self._current_index += 1
                        self._process_next_classification()
                    except:
                        pass
                
                self.Dispatcher.BeginInvoke(
                    DispatcherPriority.Normal,
                    Action(handle_error)
                )
        
        from System.Threading import Thread, ThreadStart, ApartmentState
        thread = Thread(ThreadStart(do_classification))
        thread.IsBackground = True
        thread.SetApartmentState(ApartmentState.STA)
        thread.Start()
    
    def _finish_classification(self):
        """Finish the classification process"""
        try:
            logger.info("Finishing classification.")
            
            # Re-enable button
            self.getAiSuggestionsButton.IsEnabled = True
            
            # Clear any remaining processing states
            for item in self._selected_items_for_classification:
                if item.IsProcessing:
                    item.IsProcessing = False
                if item.IsQueued:
                    item.IsQueued = False
                    
        except Exception as ex:
            logger.error("Error in _finish_classification: {}".format(str(ex)))
    
    def _create_ifc_results_listbox(self):
        """Create the IFC results ListBox programmatically to avoid XAML parsing issues"""
        try:
            from System.Windows import FrameworkElementFactory, DataTemplate
            from System.Windows.Controls import ListBox, Border as WpfBorder, StackPanel, TextBlock, ToolTip
            from System.Windows.Data import Binding
            from System.Windows import Thickness
            from System.Windows.Media import Brushes
            
            # Create ListBox
            self.ifcResultsListBox = ListBox()
            self.ifcResultsListBox.SelectionMode = System.Windows.Controls.SelectionMode.Single
            self.ifcResultsListBox.Background = Brushes.Transparent
            self.ifcResultsListBox.BorderThickness = Thickness(0)
            
            # Create tooltip template for rich tooltip
            tooltip_stackpanel = FrameworkElementFactory(StackPanel)
            tooltip_stackpanel.SetValue(StackPanel.MarginProperty, Thickness(5))
            
            tooltip_name = FrameworkElementFactory(TextBlock)
            tooltip_name.SetBinding(TextBlock.TextProperty, Binding("Name"))
            tooltip_name.SetValue(TextBlock.FontWeightProperty, System.Windows.FontWeights.Bold)
            tooltip_name.SetValue(TextBlock.MarginProperty, Thickness(0, 0, 0, 5))
            tooltip_stackpanel.AppendChild(tooltip_name)
            
            tooltip_desc = FrameworkElementFactory(TextBlock)
            tooltip_desc.SetBinding(TextBlock.TextProperty, Binding("Description"))
            tooltip_desc.SetValue(TextBlock.TextWrappingProperty, System.Windows.TextWrapping.Wrap)
            tooltip_desc.SetValue(TextBlock.MarginProperty, Thickness(0, 0, 0, 5))
            tooltip_stackpanel.AppendChild(tooltip_desc)
            
            # Predefined types in tooltip
            tooltip_types_label = FrameworkElementFactory(TextBlock)
            tooltip_types_label.SetValue(TextBlock.TextProperty, "Predefined Types:")
            tooltip_types_label.SetValue(TextBlock.FontWeightProperty, System.Windows.FontWeights.SemiBold)
            tooltip_types_label.SetValue(TextBlock.MarginProperty, Thickness(0, 0, 0, 3))
            tooltip_stackpanel.AppendChild(tooltip_types_label)
            
            tooltip_types = FrameworkElementFactory(TextBlock)
            tooltip_types.SetBinding(TextBlock.TextProperty, Binding("PredefinedTypesDisplay"))
            tooltip_types.SetValue(TextBlock.TextWrappingProperty, System.Windows.TextWrapping.Wrap)
            tooltip_stackpanel.AppendChild(tooltip_types)
            
            tooltip_template = DataTemplate()
            tooltip_template.VisualTree = tooltip_stackpanel
            
            # Create DataTemplate programmatically
            # Root: Border
            border_factory = FrameworkElementFactory(WpfBorder)
            border_factory.SetValue(WpfBorder.BorderThicknessProperty, Thickness(0, 0, 0, 1))
            border_factory.SetValue(WpfBorder.PaddingProperty, Thickness(8, 6, 8, 6))
            border_factory.SetValue(WpfBorder.MarginProperty, Thickness(0))
            
            # Set tooltip on border
            tooltip_binding = Binding()
            tooltip_binding.Path = System.Windows.PropertyPath(".")
            tooltip_binding.RelativeSource = System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent)
            
            # StackPanel inside Border
            stackpanel_factory = FrameworkElementFactory(StackPanel)
            border_factory.AppendChild(stackpanel_factory)
            
            # TextBlock 1: Name
            name_text = FrameworkElementFactory(TextBlock)
            name_text.SetBinding(TextBlock.TextProperty, Binding("Name"))
            name_text.SetValue(TextBlock.FontWeightProperty, System.Windows.FontWeights.SemiBold)
            name_text.SetValue(TextBlock.MarginProperty, Thickness(0, 0, 0, 3))
            stackpanel_factory.AppendChild(name_text)
            
            # TextBlock 2: Description
            desc_text = FrameworkElementFactory(TextBlock)
            desc_text.SetBinding(TextBlock.TextProperty, Binding("Description"))
            desc_text.SetValue(TextBlock.TextWrappingProperty, System.Windows.TextWrapping.Wrap)
            desc_text.SetValue(TextBlock.FontSizeProperty, 11.0)
            desc_text.SetValue(TextBlock.OpacityProperty, 0.8)
            stackpanel_factory.AppendChild(desc_text)
            
            # Create and set DataTemplate
            data_template = DataTemplate()
            data_template.VisualTree = border_factory
            
            # Add tooltip to data template using a style setter
            # We'll handle tooltips via MouseEnter event instead for simplicity
            self.ifcResultsListBox.ItemTemplate = data_template
            
            # Handle MouseEnter to set tooltips dynamically
            self.ifcResultsListBox.Loaded += self._setup_tooltips_on_items
            
            # Set ItemsSource
            self.ifcResultsListBox.ItemsSource = self.ifc_results
            
            # Bind SelectionChanged event
            self.ifcResultsListBox.SelectionChanged += self.IfcResultsListBox_SelectionChanged
            
            # Create "No results" text block
            self.noResultsTextBlock = TextBlock()
            self.noResultsTextBlock.Text = "No results..."
            self.noResultsTextBlock.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
            self.noResultsTextBlock.VerticalAlignment = System.Windows.VerticalAlignment.Center
            self.noResultsTextBlock.Foreground = Brushes.Gray
            self.noResultsTextBlock.FontStyle = System.Windows.FontStyles.Italic
            self.noResultsTextBlock.Visibility = System.Windows.Visibility.Collapsed
            
            # Create container grid for ListBox and "No results" message
            container_grid = System.Windows.Controls.Grid()
            # Make ListBox fill the grid
            self.ifcResultsListBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
            self.ifcResultsListBox.VerticalAlignment = System.Windows.VerticalAlignment.Stretch
            container_grid.Children.Add(self.ifcResultsListBox)
            container_grid.Children.Add(self.noResultsTextBlock)
            
            # Add to container - grid fills the border
            container_grid.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
            container_grid.VerticalAlignment = System.Windows.VerticalAlignment.Stretch
            self.ifcResultsContainer.Child = container_grid
            
            logger.info("IFC results ListBox created programmatically")
        except Exception as ex:
            logger.error("Error creating IFC results ListBox: {}".format(str(ex)))
            raise
    
    def _setup_tooltips_on_items(self, sender, e):
        """Set up tooltips on listbox items when loaded"""
        try:
            from System.Windows.Controls import ItemContainerGenerator, ToolTip, Border as WpfBorder
            from System.Windows.Media import VisualTreeHelper
            
            # This will be called when items are generated
            def on_items_changed(s, args):
                try:
                    if not hasattr(self, 'ifcResultsListBox') or not self.ifcResultsListBox:
                        return
                    
                    generator = self.ifcResultsListBox.ItemContainerGenerator
                    for i in range(self.ifc_results.Count):
                        try:
                            container = generator.ContainerFromIndex(i)
                            if container:
                                item = self.ifc_results[i]
                                tooltip = self._create_tooltip_for_item(item)
                                if tooltip:
                                    # Find the border in the visual tree
                                    border = self._find_child(container, WpfBorder)
                                    if border:
                                        border.ToolTip = tooltip
                        except:
                            pass
                except:
                    pass
            
            # Subscribe to collection changes
            if hasattr(self, 'ifc_results') and self.ifc_results:
                self.ifc_results.CollectionChanged += lambda s, args: self.Dispatcher.BeginInvoke(
                    System.Windows.Threading.DispatcherPriority.Loaded,
                    System.Action(lambda: on_items_changed(None, None))
                )
        except Exception as ex:
            logger.debug("Error setting up tooltips: {}".format(str(ex)))
    
    def _find_child(self, parent, target_type):
        """Find a child element of a specific type in the visual tree"""
        try:
            from System.Windows.Media import VisualTreeHelper
            if parent is None:
                return None
            for i in range(VisualTreeHelper.GetChildrenCount(parent)):
                child = VisualTreeHelper.GetChild(parent, i)
                if child and isinstance(child, target_type):
                    return child
                result = self._find_child(child, target_type)
                if result:
                    return result
        except:
            pass
        return None
    
    def _create_tooltip_for_item(self, item):
        """Create tooltip text for an IFC class item"""
        if not item:
            return ""
        
        from System.Windows.Controls import StackPanel, TextBlock
        from System.Windows import Thickness
        
        tooltip_panel = StackPanel()
        tooltip_panel.Margin = Thickness(5)
        
        # Name
        name_block = TextBlock()
        name_block.Text = item.name
        name_block.FontWeight = System.Windows.FontWeights.Bold
        name_block.Margin = Thickness(0, 0, 0, 5)
        tooltip_panel.Children.Add(name_block)
        
        # Description
        if item.description:
            desc_block = TextBlock()
            desc_block.Text = item.description
            desc_block.TextWrapping = System.Windows.TextWrapping.Wrap
            desc_block.MaxWidth = 400
            desc_block.Margin = Thickness(0, 0, 0, 5)
            tooltip_panel.Children.Add(desc_block)
        
        # Predefined types
        if item.predefined_types:
            types_label = TextBlock()
            types_label.Text = "Predefined Types:"
            types_label.FontWeight = System.Windows.FontWeights.SemiBold
            types_label.Margin = Thickness(0, 0, 0, 3)
            tooltip_panel.Children.Add(types_label)
            
            types_list = list(item.predefined_types.keys())[:10]
            types_str = ", ".join(types_list)
            if len(item.predefined_types) > 10:
                types_str += "..."
            
            types_block = TextBlock()
            types_block.Text = types_str
            types_block.TextWrapping = System.Windows.TextWrapping.Wrap
            types_block.MaxWidth = 400
            tooltip_panel.Children.Add(types_block)
        
        return tooltip_panel
    
    def IfcReferenceButton_Click(self, sender, e):
        """Open Bonsai BIM IFC Class Reference website"""
        try:
            import webbrowser
            webbrowser.open("https://bonsaibim.org/search-ifc-class.html")
            logger.info("Opened Bonsai BIM IFC Class Reference")
        except Exception as ex:
            logger.error("Error opening IFC reference: {}".format(str(ex)))
            forms.alert("Could not open browser: {}".format(str(ex)), title="Error")
    
    def OpenIfcSearch_Click(self, sender, e):
        """Load IFC schema when user clicks search button in DataGrid cell"""
        try:
            button = sender
            item = button.Tag
            if item:
                # Store target item for IFC class application
                self._target_item_for_ifc = item
                # Load schema if not already loaded
                if not self.current_schema_data:
                    self._load_ifc_schema("IFC2X3")
                else:
                    self._perform_ifc_search()
        except Exception as ex:
            logger.error("Error in OpenIfcSearch_Click: {}".format(str(ex)))
            forms.alert("Error: {}".format(str(ex)), title="Error")
    
    def _load_ifc_schema(self, version):
        """Load IFC schema and update UI"""
        # Busy overlay removed - using synchronous loading
        logger.info("Loading IFC {} schema...".format(version))
        
        try:
            schema_data = load_ifc_schema(version)
            if schema_data:
                self.current_schema_version = version
                self.current_schema_data = schema_data
                # Trigger search with current query
                self._perform_ifc_search()
            else:
                forms.alert("Failed to load IFC {} schema".format(version), title="Error")
        except Exception as e:
            logger.error("Error loading IFC schema: {}".format(str(e)))
            forms.alert("Error loading IFC schema: {}".format(str(e)), title="Error")
    
    def _perform_ifc_search(self):
        """Perform fuzzy search on IFC classes"""
        if not self.current_schema_data:
            return
        
        query_text = self.ifcSearchTextBox.Text or ""
        results = search_ifc_classes(query_text, self.current_schema_data)
        
        # Update ListBox - clear selection first to prevent IsSelected errors
        try:
            # Temporarily unbind SelectionChanged to prevent errors during update
            try:
                self.ifcResultsListBox.SelectionChanged -= self.IfcResultsListBox_SelectionChanged
            except:
                pass
            
            # Clear existing items
            if not hasattr(self, 'ifc_results'):
                from System.Collections.ObjectModel import ObservableCollection
                self.ifc_results = ObservableCollection[IfcClassInfo]()
            else:
                self.ifc_results.Clear()
            
            # Add new results
            for result in results:
                self.ifc_results.Add(result)
            
            # Show/hide "No results" message
            if hasattr(self, 'noResultsTextBlock'):
                if len(results) == 0:
                    self.noResultsTextBlock.Visibility = System.Windows.Visibility.Visible
                    self.ifcResultsListBox.Visibility = System.Windows.Visibility.Collapsed
                else:
                    self.noResultsTextBlock.Visibility = System.Windows.Visibility.Collapsed
                    self.ifcResultsListBox.Visibility = System.Windows.Visibility.Visible
                
                # Set up tooltips for new items
                self._update_tooltips_for_items()
            
            # Rebind SelectionChanged
            self.ifcResultsListBox.SelectionChanged += self.IfcResultsListBox_SelectionChanged
        except Exception as ex:
            logger.error("Error updating IFC results list: {}".format(str(ex)))
            # Try to rebind anyway
            try:
                self.ifcResultsListBox.SelectionChanged += self.IfcResultsListBox_SelectionChanged
            except:
                pass
    
    def _update_tooltips_for_items(self):
        """Update tooltips for all items in the listbox"""
        try:
            from System.Windows.Controls import Border as WpfBorder
            from System.Windows.Media import VisualTreeHelper
            from System.Windows.Threading import DispatcherPriority
            from System import Action
            
            if not hasattr(self, 'ifcResultsListBox') or not self.ifcResultsListBox:
                return
            
            def update_tooltips():
                try:
                    generator = self.ifcResultsListBox.ItemContainerGenerator
                    for i in range(self.ifc_results.Count):
                        try:
                            container = generator.ContainerFromIndex(i)
                            if container:
                                item = self.ifc_results[i]
                                tooltip = self._create_tooltip_for_item(item)
                                if tooltip:
                                    # Find the border in the visual tree
                                    border = self._find_child(container, WpfBorder)
                                    if border:
                                        border.ToolTip = tooltip
                        except:
                            pass
                except:
                    pass
            
            # Defer to after items are rendered
            self.Dispatcher.BeginInvoke(
                DispatcherPriority.Loaded,
                Action(update_tooltips)
            )
        except Exception as ex:
            logger.debug("Error updating tooltips: {}".format(str(ex)))
    
    def SchemaVersionComboBox_SelectionChanged(self, sender, e):
        """Handle schema version selection changed"""
        combo_box = sender
        selected_item = combo_box.SelectedItem
        if selected_item:
            version = selected_item.Content
            self._load_ifc_schema(version)
    
    def IfcSearchTextBox_TextChanged(self, sender, e):
        """Handle IFC search text changed"""
        self._perform_ifc_search()
    
    def IfcResultsListBox_SelectionChanged(self, sender, e):
        """Handle IFC results list selection changed - apply automatically to selected rows"""
        try:
            if not self.ifcResultsListBox:
                return
            selected_item = self.ifcResultsListBox.SelectedItem
            if selected_item and isinstance(selected_item, IfcClassInfo):
                self._selected_ifc_class = selected_item
                self._selected_predefined_type = None
                
                # Apply automatically to selected rows (details now shown in tooltip)
                self._apply_ifc_class_to_selected()
            else:
                self._selected_ifc_class = None
                self._selected_predefined_type = None
        except Exception as ex:
            logger.error("Error in IfcResultsListBox_SelectionChanged: {}".format(str(ex)))
            # Don't re-raise - just log the error to prevent window crash
    
    def _apply_ifc_class_to_selected(self):
        """Apply selected IFC class to selected rows (called automatically)"""
        if not self._selected_ifc_class:
            return
        
        # Determine target items
        if hasattr(self, '_target_item_for_ifc') and self._target_item_for_ifc:
            # Apply to specific item (from edit mode search button)
            target_items = [self._target_item_for_ifc]
            self._target_item_for_ifc = None
        else:
            # Apply to selected items
            target_items = [item for item in self.all_items if item.IsSelected]
            if not target_items:
                # If nothing selected, don't apply
                logger.info("No rows selected - IFC class not applied")
                return
        
        for item in target_items:
            item.IfcClass = self._selected_ifc_class.name
            if self._selected_predefined_type:
                item.PredefinedType = self._selected_predefined_type
        
        logger.info("Applied {} to {} items".format(self._selected_ifc_class.name, len(target_items)))
    
    def UpdateButton_Click(self, sender, e):
        """Update selected element types with their IFC Class and Predefined Type values"""
        try:
            # Collect selected items
            selected_items = [item for item in self.all_items if item.IsSelected]
            
            if not selected_items:
                forms.alert("Please select at least one element type to update.", title="No Selection")
                return
            
            # Prepare results in the format expected by apply_classifications_with_progress
            results = []
            for item in selected_items:
                results.append({
                    'type_info': item.type_info,
                    'ifc_class': item.IfcClass or None,
                    'predefined_type': item.PredefinedType or None
                })
            
            # Update Revit parameters
            with forms.ProgressBar(title="Updating IFC Classifications...") as pb:
                pb.update_progress(0, len(results))
                updated_count = apply_classifications_with_progress(results, pb)
            
            logger.info("Updated {} element types.".format(updated_count))
            
            # Show success notification
            if updated_count > 0:
                forms.show_balloon(
                    header="IFC Classification",
                    text="{} element type{} were updated.".format(
                        updated_count, 
                        "s" if updated_count != 1 else ""
                    ),
                    tooltip="IFC Class and Predefined Type parameters have been updated successfully.",
                    is_new=True
                )
            
            # Close window
            self.Close()
            
        except Exception as ex:
            logger.error("Error in UpdateButton_Click: {}".format(str(ex)))
            forms.alert("An error occurred when updating classifications: {}".format(str(ex)), title="Error")


def get_element_types_in_view():
    """Get all element types visible in the active view"""
    active_view = doc.ActiveView
    if not active_view:
        forms.alert("No active view found.", title="Error")
        return []
    
    logger.info("Getting element types from view: {}".format(active_view.Name))
    
    # Get all elements in view
    collector = FilteredElementCollector(doc, active_view.Id)
    elements = collector.WhereElementIsNotElementType().ToElements()
    
    # Extract unique element types
    type_ids = set()
    for element in elements:
        try:
            type_id = element.GetTypeId()
            if type_id and type_id != ElementId.InvalidElementId:
                type_ids.add(type_id)
        except:
            continue
    
    # Get the actual type elements and validate them
    element_types = []
    for type_id in type_ids:
        try:
            element_type = doc.GetElement(type_id)
            if element_type and isinstance(element_type, ElementType):
                type_info = ElementTypeInfo(element_type)
                element_types.append(type_info)
                logger.debug("Found ElementType: ID={}".format(type_id))
            else:
                if element_type:
                    logger.debug("Element {} is not an ElementType: {}".format(
                        type_id, type(element_type).__name__))
        except Exception as e:
            logger.debug("Error processing type {}: {}".format(type_id, str(e)))
            continue
    
    logger.info("Found {} unique element types".format(len(element_types)))
    return element_types

def extract_classification_data(classification):
    """Extract IFC class, predefined type, and reasoning from classification response"""
    ifc_class = None
    predefined_type = None
    reasoning = None
    
    try:
        if isinstance(classification, dict):
            # Handle the new API response format with 'output' field
            output_data = classification.get('output', {})
            if output_data:
                ifc_class = output_data.get('Class', 'Not classified')
                predefined_type = output_data.get('PredefinedType', 'Not specified')
                reasoning = output_data.get('Reasoning', 'No reasoning provided')
            else:
                # Fallback to direct dictionary response
                ifc_class = classification.get('ifc_class', 'Not classified')
                predefined_type = classification.get('predefined_type', 'Not specified')
                reasoning = classification.get('reasoning', 'No reasoning provided')
        elif isinstance(classification, list) and len(classification) > 0:
            # Handle legacy list format
            class_output = classification[0].get('output', {})
            ifc_class = class_output.get('Class', 'Not classified')
            predefined_type = class_output.get('PredefinedType', class_output.get('Type', 'Not specified'))
            reasoning = class_output.get('Reasoning', 'No reasoning provided')
    except Exception as e:
        logger.error("Error extracting classification data: {}".format(str(e)))
        ifc_class = 'Error'
        predefined_type = 'Error'
        reasoning = 'Error parsing data'
    
    return ifc_class, predefined_type, reasoning

def send_classification_request(type_data):
    """Send classification request to the API"""
    try:
        # Create web client with proper encoding
        client = WebClient()
        client.Headers.Add("Content-Type", "application/json; charset=utf-8")
        client.Encoding = Encoding.UTF8
        
        # Convert data to JSON with proper encoding
        json_data = json.dumps(type_data, ensure_ascii=False)
        logger.debug("Sending data: {}".format(json_data))
        
        # Send request with UTF-8 encoding
        response_bytes = client.UploadData(CLASSIFICATION_URL, "POST", Encoding.UTF8.GetBytes(json_data))
        response = Encoding.UTF8.GetString(response_bytes)
        
        # Parse response
        response_data = json.loads(response)
        logger.debug("Received response: {}".format(response))
        
        return response_data
        
    except Exception as e:
        logger.error("Error sending classification request: {}".format(str(e)))
        return None
    finally:
        if 'client' in locals():
            client.Dispose()

def main():
    """Main function"""
    try:
        # Get element types from active view
        element_types = get_element_types_in_view()
        
        if not element_types:
            forms.alert("No element types found in the active view.", 
                       title="No Types Found")
            return
        
        # Show classification window (single tab, all element types with editable IFC columns)
        classification_window = ClassificationWindow(element_types)
        dialog_result = classification_window.ShowDialog()
        
        logger.info("Classification window closed. Dialog result: {}".format(dialog_result))
            
    except Exception as e:
        logger.error("Error in main function: {}".format(str(e)))
        forms.alert("An error occurred: {}".format(str(e)), title="Error")

def apply_classifications_with_progress(approved_results, progress_bar):
    """Apply the classifications to element types with progress tracking"""
    if not approved_results:
        return 0
    
    updated_count = 0
    total_count = len(approved_results)
    
    with revit.Transaction("Apply IFC Classifications"):
        for i, result in enumerate(approved_results):
            if not result:
                progress_bar.update_progress(i + 1, total_count)
                continue
                
            try:
                type_info = result['type_info']
                ifc_class = result.get('ifc_class', None)
                predefined_type = result.get('predefined_type', None)
                
                element_type = type_info.element_type
                
                # Set IFC Class parameter
                if ifc_class:
                    ifc_param = element_type.LookupParameter("Export Type to IFC As")
                    if ifc_param:
                        ifc_param.Set(ifc_class)
                        updated_count += 1
                
                # Set Predefined Type parameter
                if predefined_type:
                    predefined_param = element_type.LookupParameter("Type IFC Predefined Type")
                    if not predefined_param:
                        predefined_param = element_type.LookupParameter("Predefined Type")
                    if predefined_param:
                        predefined_param.Set(predefined_type)
                        if not ifc_class:
                            updated_count += 1
                
                progress_bar.update_progress(i + 1, total_count)
                
            except Exception as e:
                logger.error("Error applying classification to {}: {}".format(
                    type_info.type_name, str(e)))
                progress_bar.update_progress(i + 1, total_count)
                continue
    
    return updated_count

# Call main function
main()
