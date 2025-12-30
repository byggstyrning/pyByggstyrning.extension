# -*- coding: utf-8 -*-
"""Configure spatial parameter mappings for 3D Zones, Rooms and Areas.

Allows users to create, edit, and manage configurations for mapping
parameters from spatial elements to contained elements.
"""

__title__ = "Edit Spatial\nMappings"
__author__ = "Byggstyrning AB"
__doc__ = "Configure spatial parameter mappings for 3D Zones, Rooms and Areas"

# Import standard libraries
import sys
import os

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import *

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Add the extension directory to the path
import os.path as op
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

# Initialize logger
logger = script.get_logger()

# Import zone3d libraries
try:
    from zone3d import config
except ImportError as e:
    logger.error("Failed to import zone3d libraries: {}".format(e))
    forms.alert("Failed to import required libraries. Check logs for details.")
    script.exit()

# Import WPF components
from System import EventHandler, Action
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Specialized import NotifyCollectionChangedEventArgs
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows import MessageBox, MessageBoxButton, MessageBoxResult, Visibility, Thickness
from System.Windows.Controls import SelectionMode
from System.Windows.Documents import Run, LineBreak
from System.Windows.Media import Brushes, SolidColorBrush
from System.Windows import FontWeights
from System.Windows.Media import Colors
import System.Windows.Forms
import System.Windows.Threading

# Special marker for 3D Zone filter (Generic Models with family name containing 3DZone)
THREE_D_ZONE_MARKER = "3DZONE_FILTER"

# --- Helper Classes ---

class ConfigItem(INotifyPropertyChanged):
    """Configuration item for display in the list."""
    def __init__(self, config_dict):
        self.config_dict = config_dict
        self.id = config_dict.get("id", "")
        self.name = config_dict.get("name", "Unknown")
        self.order = config_dict.get("order", 0)
        self.enabled = config_dict.get("enabled", False)
        self.source_categories = config_dict.get("source_categories", [])
        self.source_params = config_dict.get("source_params", [])
        self.target_params = config_dict.get("target_params", [])
        self.target_filter_categories = config_dict.get("target_filter_categories", [])
        self.write_before_ifc_export = config_dict.get("write_before_ifc_export", False)
        self.ifc_export_only_empty = config_dict.get("ifc_export_only_empty", False)
        self.use_linked_document = config_dict.get("use_linked_document", False)
        self.linked_document_name = config_dict.get("linked_document_name", None)
        self.source_sort_property = config_dict.get("source_sort_property", "ElementId")
        # Initialize the event handler list
        self._property_changed_handlers = []
    
    def add_PropertyChanged(self, handler):
        """Add a PropertyChanged event handler."""
        if handler is not None:
            self._property_changed_handlers.append(handler)
    
    def remove_PropertyChanged(self, handler):
        """Remove a PropertyChanged event handler."""
        if handler is not None and handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    
    def _notify_property_changed(self, property_name):
        """Notify that a property has changed."""
        # Invoke all registered handlers
        if self._property_changed_handlers:
            args = PropertyChangedEventArgs(property_name)
            for handler in self._property_changed_handlers:
                try:
                    handler(self, args)
                except Exception as e:
                    pass
    
    @property
    def Name(self):
        return self.name
    
    @property
    def Order(self):
        return self.order
    
    @property
    def Status(self):
        return "Enabled" if self.enabled else "Disabled"
    
    @property
    def SourceCategoriesDisplay(self):
        """Display source categories as comma-separated string."""
        cat_names = []
        for cat in self.source_categories:
            if cat == THREE_D_ZONE_MARKER:
                cat_names.append("3D Zone")
            elif isinstance(cat, BuiltInCategory):
                cat_names.append(str(cat))
            else:
                cat_names.append(str(cat))
        return ", ".join(cat_names[:3]) + ("..." if len(cat_names) > 3 else "")
    
    @property
    def MappingsDisplay(self):
        """Display parameter mappings."""
        if not self.source_params or not self.target_params:
            return "No mappings"
        mappings = []
        for src, tgt in zip(self.source_params, self.target_params):
            mappings.append("{} -> {}".format(src, tgt))
        return "; ".join(mappings[:2]) + ("..." if len(mappings) > 2 else "")
    
    @property
    def MappingsFormattedText(self):
        """Display parameter mappings with line breaks and styled arrows."""
        if not self.source_params or not self.target_params:
            return "No mappings"
        mappings = []
        for src, tgt in zip(self.source_params, self.target_params):
            # Format: source -> target (arrow will be styled in XAML converter)
            mappings.append("{} -> {}".format(src, tgt))
        # Join with line breaks
        return "\n".join(mappings)
    
    @property
    def IfcExportDisplay(self):
        return "Yes" if self.write_before_ifc_export else "No"
    
    @property
    def Enabled(self):
        """Property for checkbox binding."""
        return self.enabled
    
    @Enabled.setter
    def Enabled(self, value):
        """Set enabled state and update config dict."""
        old_value = self.enabled
        new_value = bool(value)
        if old_value != new_value:
            self.enabled = new_value
            self.config_dict["enabled"] = self.enabled
            self._notify_property_changed("Enabled")
    
    @property
    def WriteBeforeIfcExport(self):
        """Property for checkbox binding."""
        return self.write_before_ifc_export
    
    @WriteBeforeIfcExport.setter
    def WriteBeforeIfcExport(self, value):
        """Set write_before_ifc_export state and update config dict."""
        old_value = self.write_before_ifc_export
        new_value = bool(value)
        if old_value != new_value:
            self.write_before_ifc_export = new_value
            self.config_dict["write_before_ifc_export"] = self.write_before_ifc_export
            # Notify WPF that the property changed
            self._notify_property_changed("WriteBeforeIfcExport")
    
    @property
    def OnlyEmpty(self):
        """Property for checkbox binding."""
        return self.ifc_export_only_empty
    
    @OnlyEmpty.setter
    def OnlyEmpty(self, value):
        """Set ifc_export_only_empty state and update config dict."""
        old_value = self.ifc_export_only_empty
        new_value = bool(value)
        if old_value != new_value:
            self.ifc_export_only_empty = new_value
            self.config_dict["ifc_export_only_empty"] = self.ifc_export_only_empty
            self._notify_property_changed("OnlyEmpty")
    
    @property
    def UseLinkedDocument(self):
        """Property for checkbox binding."""
        return self.use_linked_document
    
    @UseLinkedDocument.setter
    def UseLinkedDocument(self, value):
        """Set use_linked_document state and update config dict."""
        old_value = self.use_linked_document
        new_value = bool(value)
        if old_value != new_value:
            self.use_linked_document = new_value
            self.config_dict["use_linked_document"] = self.use_linked_document
            self._notify_property_changed("UseLinkedDocument")
    
    @property
    def LinkedDocumentName(self):
        """Property for dropdown binding."""
        return self.linked_document_name
    
    @LinkedDocumentName.setter
    def LinkedDocumentName(self, value):
        """Set linked_document_name and update config dict."""
        old_value = self.linked_document_name
        new_value = value if value else None
        if old_value != new_value:
            self.linked_document_name = new_value
            self.config_dict["linked_document_name"] = self.linked_document_name
            self._notify_property_changed("LinkedDocumentName")

class ParameterMappingEntry(object):
    """Parameter mapping entry for DataGrid."""
    def __init__(self, SourceParameter="", TargetParameter=""):
        self._source_parameter = SourceParameter
        self._target_parameter = TargetParameter
        # Set arrow image source
        try:
            arrow_path = op.join(tab_dir, 'lib', 'styles', 'icons', 'arrow.png')
            if op.exists(arrow_path):
                from System.Windows.Media.Imaging import BitmapImage
                from System import Uri
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(arrow_path)
                bitmap.EndInit()
                self._arrow_image_source = bitmap
            else:
                logger.debug("Arrow image not found at: {}".format(arrow_path))
                self._arrow_image_source = None
        except Exception as e:
            logger.warning("Could not load arrow image: {}".format(str(e)))
            import traceback
            logger.debug(traceback.format_exc())
            self._arrow_image_source = None

    @property
    def SourceParameter(self):
        return self._source_parameter

    @SourceParameter.setter
    def SourceParameter(self, value):
        self._source_parameter = value

    @property
    def TargetParameter(self):
        return self._target_parameter

    @TargetParameter.setter
    def TargetParameter(self, value):
        self._target_parameter = value
    
    @property
    def ArrowImageSource(self):
        return self._arrow_image_source

class ParameterSelectorDialog(forms.WPFWindow):
    """Dialog for selecting source and target parameters side by side."""
    def __init__(self, available_params, initial_source=None, initial_target=None):
        """Initialize the parameter selector dialog.
        
        Args:
            available_params: List of available parameter names
            initial_source: Optional initial source parameter name to preselect
            initial_target: Optional initial target parameter name to preselect
        """
        xaml_path = op.join(pushbutton_dir, 'ParameterSelector.xaml')
        forms.WPFWindow.__init__(self, xaml_path)
        
        # Load common styles programmatically (same as Zone3DConfigEditorUI)
        self.load_styles()
        
        self.selected_source = None
        self.selected_target = None
        
        # Store initial selections for sorting
        self.initial_source = initial_source
        self.initial_target = initial_target
        
        # Sort parameters: put initial selections at the top
        source_params_sorted = self.sort_params_with_priority(available_params, initial_source)
        target_params_sorted = self.sort_params_with_priority(available_params, initial_target)
        
        # Store full lists for filtering (maintain sort order)
        self.all_source_params = source_params_sorted
        self.all_target_params = target_params_sorted
        
        # Populate list boxes with sorted parameters
        source_selected_index = -1
        target_selected_index = -1
        for i, param in enumerate(source_params_sorted):
            self.sourceParameterListBox.Items.Add(param)
            if initial_source and param == initial_source:
                source_selected_index = i
        
        for i, param in enumerate(target_params_sorted):
            self.targetParameterListBox.Items.Add(param)
            if initial_target and param == initial_target:
                target_selected_index = i
        
        # Preselect initial values if provided
        if source_selected_index >= 0:
            self.sourceParameterListBox.SelectedIndex = source_selected_index
        if target_selected_index >= 0:
            self.targetParameterListBox.SelectedIndex = target_selected_index
        
        # Set up event handlers
        self.okButton.Click += self.ok_button_click
        self.cancelButton.Click += self.cancel_button_click
        
        # Wire up filter text changed events
        self.sourceFilterTextBox.TextChanged += self.source_filter_text_changed
        self.targetFilterTextBox.TextChanged += self.target_filter_text_changed
        
        # Enable OK button only when both are selected (or if initial values were provided)
        has_source = self.sourceParameterListBox.SelectedItem is not None
        has_target = self.targetParameterListBox.SelectedItem is not None
        self.okButton.IsEnabled = has_source and has_target
        self.sourceParameterListBox.SelectionChanged += self.selection_changed
        self.targetParameterListBox.SelectionChanged += self.selection_changed
    
    def load_styles(self):
        """Load the common styles ResourceDictionary."""
        try:
            import styles
            styles.load_common_styles(self)
        except ImportError:
            try:
                from lib import styles
                styles.load_common_styles(self)
            except Exception as e:
                logger.warning("ParameterSelectorDialog: Could not load styles: {}".format(e))
        except Exception as e:
            logger.warning("ParameterSelectorDialog: Could not load styles: {}".format(e))
    
    def sort_params_with_priority(self, params, priority_param):
        """Sort parameters list with priority parameter at the top.
        
        Args:
            params: List of parameter names
            priority_param: Parameter name to put at the top (if exists)
            
        Returns:
            Sorted list with priority parameter first
        """
        if not priority_param or priority_param not in params:
            return params
        
        # Create a copy and move priority to front
        sorted_params = list(params)
        if priority_param in sorted_params:
            sorted_params.remove(priority_param)
            sorted_params.insert(0, priority_param)
        
        return sorted_params
    
    def source_filter_text_changed(self, sender, args):
        """Handle source filter text changed event."""
        self.filter_listbox(self.sourceFilterTextBox, self.sourceParameterListBox, self.all_source_params, self.initial_source)
    
    def target_filter_text_changed(self, sender, args):
        """Handle target filter text changed event."""
        self.filter_listbox(self.targetFilterTextBox, self.targetParameterListBox, self.all_target_params, self.initial_target)
    
    def filter_listbox(self, filter_textbox, listbox, all_params, priority_param=None):
        """Filter a ListBox based on TextBox input, maintaining priority order.
        
        Args:
            filter_textbox: The TextBox containing the filter text
            listbox: The ListBox to filter
            all_params: List of all available parameters (already sorted with priority)
            priority_param: Optional parameter to keep at top when filtering
        """
        filter_text = filter_textbox.Text.lower() if filter_textbox.Text else ""
        
        # Remember current selection
        current_selection = listbox.SelectedItem
        
        # Clear and repopulate with filtered items
        listbox.Items.Clear()
        
        # Filter while maintaining priority order
        filtered_params = []
        priority_added = False
        
        for param in all_params:
            if not filter_text or filter_text in param.lower():
                # If this is the priority param, add it first
                if priority_param and param == priority_param:
                    filtered_params.insert(0, param)
                    priority_added = True
                else:
                    filtered_params.append(param)
        
        # Add filtered items to listbox
        for param in filtered_params:
            listbox.Items.Add(param)
        
        # Restore selection if it still exists in filtered list
        if current_selection and current_selection in listbox.Items:
            listbox.SelectedItem = current_selection
    
    def selection_changed(self, sender, args):
        """Handle selection changed in either list box."""
        has_source = self.sourceParameterListBox.SelectedItem is not None
        has_target = self.targetParameterListBox.SelectedItem is not None
        self.okButton.IsEnabled = has_source and has_target
    
    def ok_button_click(self, sender, args):
        """Handle OK button click."""
        if self.sourceParameterListBox.SelectedItem and self.targetParameterListBox.SelectedItem:
            self.selected_source = str(self.sourceParameterListBox.SelectedItem)
            self.selected_target = str(self.targetParameterListBox.SelectedItem)
            self.DialogResult = True
            self.Close()
    
    def cancel_button_click(self, sender, args):
        """Handle Cancel button click."""
        self.DialogResult = False
        self.Close()

# --- Helper Functions ---

def get_category_options():
    """Get list of common BuiltInCategory options for selection."""
    return [
        ("Rooms", BuiltInCategory.OST_Rooms),
        ("Spaces", BuiltInCategory.OST_MEPSpaces),
        ("Areas", BuiltInCategory.OST_Areas),
        ("Mass", BuiltInCategory.OST_Mass),
        ("3D Zone (custom family)", THREE_D_ZONE_MARKER)
    ]

def get_linked_documents(doc):
    """Get all linked Revit documents in the current document.
    
    Args:
        doc: Revit document
        
    Returns:
        list: List of tuples (link_name, link_instance) for all linked documents
    """
    linked_docs = []
    try:
        link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
        for link in link_instances:
            try:
                link_doc = link.GetLinkDocument()
                if link_doc:
                    # Use the link instance name (this is what appears in Revit UI)
                    link_name = link.Name
                    linked_docs.append((link_name, link))
            except Exception as e:
                logger.debug("Error accessing linked document {}: {}".format(link.Name if hasattr(link, 'Name') else 'Unknown', str(e)))
                continue
        # Sort by name for consistent display
        linked_docs.sort(key=lambda x: x[0])
        return linked_docs
    except Exception as e:
        logger.error("Error getting linked documents: {}".format(str(e)))
        return []

def get_linked_document_by_name(doc, link_name):
    """Find a linked document by its name.
    
    Args:
        doc: Revit document
        link_name: Name of the linked document to find
        
    Returns:
        RevitLinkInstance: The link instance if found, None otherwise
    """
    if not link_name:
        return None
    try:
        linked_docs = get_linked_documents(doc)
        for name, link_instance in linked_docs:
            if name == link_name:
                return link_instance
        return None
    except Exception as e:
        logger.error("Error finding linked document by name '{}': {}".format(link_name, str(e)))
        return None

def get_all_model_categories(doc):
    """Get all Model categories from the document.
    
    Args:
        doc: Revit document
        
    Returns:
        list: List of tuples (category_name, BuiltInCategory) for all Model categories
    """
    from Autodesk.Revit.DB import CategoryType
    
    categories = []
    try:
        # Get all categories from document settings
        all_categories = doc.Settings.Categories
        
        for category in all_categories:
            # Only include Model categories that allow bound parameters
            if category.CategoryType == CategoryType.Model and category.AllowsBoundParameters:
                try:
                    # Get BuiltInCategory from category ID
                    cat_id = category.Id.IntegerValue
                    # Built-in categories have negative IDs
                    if cat_id < 0:
                        builtin_cat = BuiltInCategory(cat_id)
                        categories.append((category.Name, builtin_cat))
                except:
                    # Skip categories that can't be converted to BuiltInCategory
                    continue
        
        # Sort by category name
        categories.sort(key=lambda x: x[0])
        return categories
    except Exception as e:
        logger.error("Error getting Model categories: {}".format(str(e)))
        # Fallback to common categories if error occurs
        return [
        ("Rooms", BuiltInCategory.OST_Rooms),
        ("Spaces", BuiltInCategory.OST_MEPSpaces),
        ("Areas", BuiltInCategory.OST_Areas),
        ("Mass", BuiltInCategory.OST_Mass)
        ]

def get_parameters_from_element(doc, element):
    """Get list of parameter names from an element."""
    params = []
    if element:
        for param in element.Parameters:
            if param and not param.IsReadOnly:
                params.append(param.Definition.Name)
    return sorted(set(params))

def get_all_writable_instance_parameters(doc):
    """Get all writable instance parameters from the project without iterating all elements.
    
    Uses ParameterBindings for project/shared parameters and samples one element
    per category for built-in and category-specific parameters. Much faster than
    iterating all elements.
    
    Returns:
        list: Sorted list of unique writable instance parameter names
    """
    params = set()
    processed_categories = set()
    
    # Method 1: Get project parameters and shared parameters from ParameterBindings
    try:
        param_bindings = doc.ParameterBindings
        iterator = param_bindings.ForwardIterator()
        
        while iterator.MoveNext():
            definition = iterator.Key
            binding = iterator.Current
            
            if isinstance(binding, InstanceBinding):
                param_name = definition.Name
                try:
                    if definition.ParameterType in [
                        ParameterType.Text,
                        ParameterType.Number,
                        ParameterType.Integer,
                        ParameterType.YesNo
                    ]:
                        params.add(param_name)
                except:
                    try:
                        if definition.StorageType in [
                            StorageType.String,
                            StorageType.Double,
                            StorageType.Integer
                        ]:
                            params.add(param_name)
                    except:
                        pass
    except Exception as e:
        logger.warning("Error reading ParameterBindings: {}".format(str(e)))
    
    # Method 2: Sample one element from each category
    try:
        category_collector = FilteredElementCollector(doc)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        for element in category_collector:
            try:
                if not element or not element.Category:
                    continue
                
                cat_id = element.Category.Id.IntegerValue
                
                if cat_id in processed_categories:
                    continue
                
                processed_categories.add(cat_id)
                
                for param in element.Parameters:
                    try:
                        if (param and 
                            not param.IsReadOnly and 
                            param.StorageType in [StorageType.String, StorageType.Double, StorageType.Integer, StorageType.ElementId]):
                            params.add(param.Definition.Name)
                    except:
                        continue
            except:
                continue
    except Exception as e:
        logger.warning("Error sampling category parameters: {}".format(str(e)))
    
    return sorted(list(params))

def category_to_display_name(category):
    """Convert category to display name."""
    if category == THREE_D_ZONE_MARKER:
        return "3D Zone (Generic Models with family name containing 3DZone)"
    elif isinstance(category, BuiltInCategory):
        return str(category)
    else:
        return str(category)

def display_name_to_category(display_name, category_options):
    """Convert display name back to category."""
    cat_map = {opt[0]: opt[1] for opt in category_options}
    return cat_map.get(display_name)

def get_source_element_instance_properties(doc, source_category, use_linked_document=False, linked_document_name=None):
    """Get all instance properties from source elements of a specific category.
    
    Args:
        doc: Main Revit document
        source_category: BuiltInCategory or THREE_D_ZONE_MARKER string
        use_linked_document: Whether to search in linked document
        linked_document_name: Name of linked document if use_linked_document is True
        
    Returns:
        list: Sorted list of instance parameter names, with "ElementId" as first item
    """
    params = set()
    
    try:
        # Get source document
        source_doc = doc
        if use_linked_document and linked_document_name:
            link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
            for link in link_instances:
                if link.Name == linked_document_name:
                    try:
                        link_doc = link.GetLinkDocument()
                        if link_doc:
                            source_doc = link_doc
                            break
                    except:
                        pass
        
        # Collect elements from source category
        source_elements = []
        if source_category == THREE_D_ZONE_MARKER:
            # Special case: Generic Models with family name containing "3DZone"
            collector = FilteredElementCollector(source_doc)\
                .WhereElementIsNotElementType()\
                .OfCategory(BuiltInCategory.OST_GenericModel)\
                .ToElements()
            for el in collector:
                try:
                    if hasattr(el, "Symbol") and el.Symbol:
                        family_name = el.Symbol.FamilyName
                        if family_name and "3DZone" in family_name:
                            source_elements.append(el)
                except:
                    continue
        elif isinstance(source_category, BuiltInCategory):
            collector = FilteredElementCollector(source_doc)\
                .WhereElementIsNotElementType()\
                .OfCategory(source_category)\
                .ToElements()
            source_elements = list(collector)
        
        # Sample up to 10 elements to get properties (to avoid performance issues)
        sample_elements = source_elements[:10] if len(source_elements) > 10 else source_elements
        
        # Extract instance parameters from sample elements
        for element in sample_elements:
            try:
                for param in element.Parameters:
                    try:
                        if (param and 
                            param.StorageType in [StorageType.String, StorageType.Double, StorageType.Integer, StorageType.ElementId]):
                            params.add(param.Definition.Name)
                    except:
                        continue
            except:
                continue
        
        # Always include "ElementId" as first option
        param_list = sorted(list(params))
        if "ElementId" not in param_list:
            param_list.insert(0, "ElementId")
        else:
            # Move ElementId to first position
            param_list.remove("ElementId")
            param_list.insert(0, "ElementId")
        
        return param_list
    except Exception as e:
        logger.warning("Error getting source element instance properties: {}".format(str(e)))
        # Return default list with ElementId
        return ["ElementId"]

# --- Main UI Class ---

# Module-level cache for styles ResourceDictionary
_styles_dict_cache = None

def ensure_styles_loaded():
    """Ensure CommonStyles are loaded into Application.Resources before XAML parsing."""
    global _styles_dict_cache
    
    try:
        from System.Windows import Application
        from System.Windows.Markup import XamlReader
        from System.IO import File
        
        # Check if styles are already loaded in Application.Resources
        if Application.Current is not None and Application.Current.Resources is not None:
            try:
                test_resource = Application.Current.Resources['EnhancedDataGridStyle']
                if test_resource is not None:
                    return  # Already loaded
            except:
                pass
        
        # Load styles if not cached
        if _styles_dict_cache is None:
            # Calculate extension directory correctly
            # tab_dir is already the extension root (pyByggstyrning.extension)
            styles_path = op.join(tab_dir, 'lib', 'styles', 'CommonStyles.xaml')
            
            if op.exists(styles_path):
                # Read and parse XAML
                xaml_content = File.ReadAllText(styles_path)
                _styles_dict_cache = XamlReader.Parse(xaml_content)
                logger.debug("Loaded styles from: {}".format(styles_path))
            else:
                logger.warning("CommonStyles.xaml not found at: {}".format(styles_path))
                return
        
        # Ensure Application.Current exists
        if Application.Current is None:
            from System.Windows import Application as App
            app = App()
        
        # Merge into Application.Resources
        if Application.Current.Resources is None:
            from System.Windows import ResourceDictionary
            Application.Current.Resources = ResourceDictionary()
        
        # Merge styles using MergedDictionaries (proper WPF way)
        try:
            Application.Current.Resources.MergedDictionaries.Add(_styles_dict_cache)
        except:
            # Fallback: try to copy resources manually if MergedDictionaries fails
            try:
                for key in _styles_dict_cache.Keys:
                    try:
                        if Application.Current.Resources[key] is None:
                            pass
                    except:
                        Application.Current.Resources[key] = _styles_dict_cache[key]
            except Exception as e:
                logger.warning("Could not merge styles dictionary: {}".format(str(e)))
    except Exception as e:
        logger.warning("Could not load styles into Application.Resources: {}".format(str(e)))


class Zone3DConfigEditorUI(forms.WPFWindow):
    """3D Zone Configuration Editor UI implementation."""
    
    def __init__(self):
        """Initialize the Configuration Editor UI."""
        # Load styles BEFORE window initialization
        ensure_styles_loaded()
        
        # Initialize WPF window
        xaml_path = op.join(pushbutton_dir, 'Zone3DConfigEditor.xaml')
        forms.WPFWindow.__init__(self, xaml_path)
        
        # Also load styles into window resources as fallback
        self.load_styles()
        
        # Initialize data collections
        self.configs = ObservableCollection[object]()
        self.parameter_mappings = ObservableCollection[object]()
        self.current_config = None
        self.is_new_config = False
        
        # Category mapping dictionaries for preselection
        self.source_category_map = {}  # display_name -> category_value
        self.target_category_map = {}  # display_name -> category_value
        
        # Set up event handlers
        self.configsListView.SelectionChanged += self.config_selection_changed
        self.configsListView.MouseDoubleClick += self.config_double_click
        self.addButton.Click += self.add_button_click
        self.editButton.Click += self.edit_button_click
        self.deleteButton.Click += self.delete_button_click
        self.saveButton.Click += self.save_button_click
        self.backButton.Click += self.back_button_click
        self.moveUpButton.Click += self.move_up_button_click
        self.moveDownButton.Click += self.move_down_button_click
        self.addParameterMappingButton.Click += self.add_parameter_mapping_button_click
        self.editParameterMappingButton.Click += self.edit_parameter_mapping_button_click
        self.deleteParameterMappingButton.Click += self.delete_parameter_mapping_button_click
        self.parameterMappingsDataGrid.SelectionChanged += self.parameter_mapping_selection_changed
        self.tabControl.SelectionChanged += self.tab_selection_changed
        
        # Set up validation event handlers
        self.nameTextBox.TextChanged += self.name_text_changed
        self.sourceCategoriesListBox.SelectionChanged += self.source_category_selection_changed
        self.parameter_mappings.CollectionChanged += self.parameter_mappings_collection_changed
        
        # Set up linked document event handlers
        self.useLinkedDocumentCheckBox.Checked += self.use_linked_document_checked
        self.useLinkedDocumentCheckBox.Unchecked += self.use_linked_document_unchecked
        self.linkedDocumentComboBox.SelectionChanged += self.linked_document_selection_changed
        
        # Set up sort property ComboBox handler
        self.sourceSortPropertyComboBox.SelectionChanged += self.source_sort_property_selection_changed
        
        # Initialize UI
        self.configsListView.ItemsSource = self.configs
        self.parameterMappingsDataGrid.ItemsSource = self.parameter_mappings
        
        # Initialize sort property ComboBox with default "ElementId"
        self.sourceSortPropertyComboBox.Items.Clear()
        self.sourceSortPropertyComboBox.Items.Add("ElementId")
        self.sourceSortPropertyComboBox.SelectedItem = "ElementId"
        
        # Initial UI state
        self.tabControl.SelectedItem = self.configsTab
        self.editConfigTab.IsEnabled = False
        
        # Initialize validation state
        self._validation_errors = {
            "name": False,
            "source_category": False,
            "mappings": False,
            "linked_document": False
        }
        
        # Load configurations
        self.load_configurations()
    
    def load_styles(self):
        """Load the common styles ResourceDictionary."""
        try:
            import styles
            styles.load_common_styles(self)
        except ImportError:
            try:
                from lib import styles
                styles.load_common_styles(self)
            except Exception as e:
                logger.warning("Could not load styles: {}".format(e))
        except Exception as e:
            logger.warning("Could not load styles: {}".format(e))
    
    def EnabledCheckBox_Checked(self, sender, args):
        """Handle Enabled checkbox checked event."""
        config_item = sender.DataContext
        if config_item:
            config_item.Enabled = True
            self.save_config_from_item(config_item)
    
    def EnabledCheckBox_Unchecked(self, sender, args):
        """Handle Enabled checkbox unchecked event."""
        config_item = sender.DataContext
        if config_item:
            config_item.Enabled = False
            self.save_config_from_item(config_item)
    
    def WriteBeforeIfcCheckBox_Checked(self, sender, args):
        """Handle Write before IFC Export checkbox checked event."""
        config_item = sender.DataContext
        if config_item:
            config_item.WriteBeforeIfcExport = True
            self.save_config_from_item(config_item)
            # Enable Only Empty checkbox
            self.update_only_empty_checkbox_state(config_item)
    
    def WriteBeforeIfcCheckBox_Unchecked(self, sender, args):
        """Handle Write before IFC Export checkbox unchecked event."""
        config_item = sender.DataContext
        if config_item:
            config_item.WriteBeforeIfcExport = False
            self.save_config_from_item(config_item)
            # Note: Only Empty checkbox will be disabled via binding, but its state is preserved
            self.update_only_empty_checkbox_state(config_item)
    
    def OnlyEmptyCheckBox_Checked(self, sender, args):
        """Handle Only Empty checkbox checked event."""
        config_item = sender.DataContext
        if config_item:
            config_item.OnlyEmpty = True
            self.save_config_from_item(config_item)
    
    def OnlyEmptyCheckBox_Unchecked(self, sender, args):
        """Handle Only Empty checkbox unchecked event."""
        config_item = sender.DataContext
        if config_item:
            config_item.OnlyEmpty = False
            self.save_config_from_item(config_item)
    
    def OnlyEmptyCheckBox_Loaded(self, sender, args):
        """Handle Only Empty checkbox loaded event - check IsEnabled state."""
        pass
    
    def MappingsTextBlock_Loaded(self, sender, args):
        """Format mappings TextBlock with styled arrows."""
        config_item = sender.DataContext
        if not config_item or not config_item.source_params or not config_item.target_params:
            sender.Text = "No mappings"
            return
        
        # Clear existing inlines
        sender.Inlines.Clear()
        
        # Create formatted text with blue bold arrows
        for i, (src, tgt) in enumerate(zip(config_item.source_params, config_item.target_params)):
            if i > 0:
                sender.Inlines.Add(LineBreak())
            
            # Add source parameter
            sender.Inlines.Add(Run(src))
            
            # Add blue bold arrow
            arrow_run = Run(" -> ")
            arrow_run.Foreground = Brushes.Blue
            arrow_run.FontWeight = FontWeights.Bold
            sender.Inlines.Add(arrow_run)
            
            # Add target parameter
            sender.Inlines.Add(Run(tgt))
    
    def update_only_empty_checkbox_state(self, config_item):
        """Update Only Empty checkbox enabled state based on WriteBeforeIfcExport."""
        # Find the checkbox in the ListView item
        # This is handled by the binding in XAML, but we can force update if needed
        pass
    
    def save_config_from_item(self, config_item):
        """Save configuration from a ConfigItem."""
        try:
            # Get all configurations
            all_configs = config.load_configs(revit.doc)
            
            # Find and update the configuration
            for i, cfg in enumerate(all_configs):
                if cfg.get("id") == config_item.id:
                    # Update config with values from ConfigItem
                    all_configs[i] = {
                        "id": config_item.id,
                        "name": config_item.name,
                        "order": config_item.order,
                        "enabled": config_item.enabled,
                        "source_categories": config_item.source_categories,
                        "source_params": config_item.source_params,
                        "target_params": config_item.target_params,
                        "target_filter_categories": config_item.target_filter_categories,
                        "write_before_ifc_export": config_item.write_before_ifc_export,
                        "ifc_export_only_empty": config_item.ifc_export_only_empty,
                        "use_linked_document": config_item.use_linked_document if hasattr(config_item, 'use_linked_document') else False,
                        "linked_document_name": config_item.linked_document_name if hasattr(config_item, 'linked_document_name') else None,
                        "source_sort_property": config_item.source_sort_property if hasattr(config_item, 'source_sort_property') else "ElementId"
                    }
                    break
            
            # Save configurations
            if config.save_configs(revit.doc, all_configs):
                self.update_status("Mapping updated")
            else:
                self.update_status("Error updating mapping")
        except Exception as e:
            logger.error("Error saving configuration from item: {}".format(str(e)))
            self.update_status("Error: {}".format(str(e)))
    
    def load_configurations(self):
        """Load all configurations from storage."""
        try:
            self.configs.Clear()
            
            loaded_configs = config.load_configs(revit.doc)
            logger.debug("Loaded {} configurations from storage".format(len(loaded_configs)))
            
            if loaded_configs:
                # Sort by order
                sorted_configs = sorted(loaded_configs, key=lambda x: x.get("order", 999))
                
                for cfg in sorted_configs:
                    config_item = ConfigItem(cfg)
                    self.configs.Add(config_item)
                
                self.update_status("Loaded {} mappings".format(len(self.configs)))
            else:
                self.update_status("No mappings found")
            
            self.refresh_ui()
            
        except Exception as e:
            logger.error("Error loading configurations: {}".format(str(e)))
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))
            self.update_status("Error loading mappings: {}".format(str(e)))
    
    def refresh_ui(self):
        """Force UI refresh."""
        logger.debug("Forcing UI refresh")
        self.configsListView.ItemsSource = None
        self.configsListView.ItemsSource = self.configs
        self.configsListView.UpdateLayout()
    
    def update_status(self, message):
        """Update the status display."""
        self.statusTextBlock.Text = message
        self.process_ui_events()
    
    def validate_form(self):
        """Validate all form fields and update Save button state."""
        # Validate name
        name_valid = bool(self.nameTextBox.Text and self.nameTextBox.Text.strip())
        self._validation_errors["name"] = not name_valid
        
        # Validate source category (exactly one)
        source_category_valid = self.sourceCategoriesListBox.SelectedItem is not None
        self._validation_errors["source_category"] = not source_category_valid
        
        # Validate mappings (at least one)
        mappings_valid = False
        if self.parameter_mappings.Count > 0:
            # Check if at least one mapping has both source and target
            for mapping in self.parameter_mappings:
                if (mapping.SourceParameter and mapping.SourceParameter.strip() and
                    mapping.TargetParameter and mapping.TargetParameter.strip()):
                    mappings_valid = True
                    break
        self._validation_errors["mappings"] = not mappings_valid
        
        # Validate linked document (only if checkbox is checked)
        linked_document_valid = True
        if self.useLinkedDocumentCheckBox.IsChecked:
            linked_doc_selected = self.linkedDocumentComboBox.SelectedItem is not None
            if not linked_doc_selected:
                linked_document_valid = False
            else:
                # Also check if the selected link still exists
                selected_link_name = str(self.linkedDocumentComboBox.SelectedItem)
                link_instance = get_linked_document_by_name(revit.doc, selected_link_name)
                if not link_instance:
                    linked_document_valid = False
        self._validation_errors["linked_document"] = not linked_document_valid
        
        # Update Save button state
        all_valid = name_valid and source_category_valid and mappings_valid and linked_document_valid
        self.saveButton.IsEnabled = all_valid
        
        # Update visual feedback
        self.update_validation_visual_feedback()
        
        return all_valid
    
    def update_validation_visual_feedback(self):
        """Update visual feedback for validation errors."""
        # Update name TextBox border and error message
        if self._validation_errors["name"]:
            self.nameTextBox.BorderBrush = SolidColorBrush(Colors.Red)
            self.nameTextBox.BorderThickness = Thickness(2)
            if hasattr(self, 'nameErrorTextBlock'):
                self.nameErrorTextBlock.Visibility = Visibility.Visible
        else:
            self.nameTextBox.BorderBrush = None
            self.nameTextBox.BorderThickness = Thickness(1)
            if hasattr(self, 'nameErrorTextBlock'):
                self.nameErrorTextBlock.Visibility = Visibility.Collapsed
        
        # Update source category ListBox border and error message
        if self._validation_errors["source_category"]:
            self.sourceCategoriesListBox.BorderBrush = SolidColorBrush(Colors.Red)
            self.sourceCategoriesListBox.BorderThickness = Thickness(2)
            if hasattr(self, 'sourceCategoryErrorTextBlock'):
                self.sourceCategoryErrorTextBlock.Visibility = Visibility.Visible
        else:
            self.sourceCategoriesListBox.BorderBrush = None
            self.sourceCategoriesListBox.BorderThickness = Thickness(1)
            if hasattr(self, 'sourceCategoryErrorTextBlock'):
                self.sourceCategoryErrorTextBlock.Visibility = Visibility.Collapsed
        
        # Update mappings DataGrid border and error message
        if self._validation_errors["mappings"]:
            self.parameterMappingsDataGrid.BorderBrush = SolidColorBrush(Colors.Red)
            self.parameterMappingsDataGrid.BorderThickness = Thickness(2)
            if hasattr(self, 'mappingsErrorTextBlock'):
                self.mappingsErrorTextBlock.Visibility = Visibility.Visible
        else:
            self.parameterMappingsDataGrid.BorderBrush = None
            self.parameterMappingsDataGrid.BorderThickness = Thickness(1)
            if hasattr(self, 'mappingsErrorTextBlock'):
                self.mappingsErrorTextBlock.Visibility = Visibility.Collapsed
        
        # Update linked document ComboBox border and error message
        if self._validation_errors["linked_document"]:
            self.linkedDocumentComboBox.BorderBrush = SolidColorBrush(Colors.Red)
            self.linkedDocumentComboBox.BorderThickness = Thickness(2)
            if hasattr(self, 'linkedDocumentErrorTextBlock'):
                self.linkedDocumentErrorTextBlock.Visibility = Visibility.Visible
        else:
            self.linkedDocumentComboBox.BorderBrush = None
            self.linkedDocumentComboBox.BorderThickness = Thickness(1)
            if hasattr(self, 'linkedDocumentErrorTextBlock'):
                self.linkedDocumentErrorTextBlock.Visibility = Visibility.Collapsed
    
    def name_text_changed(self, sender, args):
        """Handle name text changed event."""
        self.validate_form()
    
    def source_category_selection_changed(self, sender, args):
        """Handle source category selection changed event."""
        self.validate_form()
        # Load instance properties for the selected source category
        self.load_source_element_properties()
    
    def load_source_element_properties(self):
        """Load instance properties from source elements into the sort property dropdown."""
        try:
            # Clear existing items
            self.sourceSortPropertyComboBox.Items.Clear()
            
            # Get selected source category
            selected_item = self.sourceCategoriesListBox.SelectedItem
            if not selected_item:
                # No category selected, just add ElementId as default
                self.sourceSortPropertyComboBox.Items.Add("ElementId")
                self.sourceSortPropertyComboBox.SelectedItem = "ElementId"
                return
            
            display_name = str(selected_item)
            source_category = self.source_category_map.get(display_name)
            
            if not source_category:
                # Category not found, use default
                self.sourceSortPropertyComboBox.Items.Add("ElementId")
                self.sourceSortPropertyComboBox.SelectedItem = "ElementId"
                return
            
            # Get linked document settings
            use_linked = self.useLinkedDocumentCheckBox.IsChecked
            linked_doc_name = None
            if use_linked and self.linkedDocumentComboBox.SelectedItem:
                linked_doc_name = str(self.linkedDocumentComboBox.SelectedItem)
            
            # Get instance properties from source elements
            properties = get_source_element_instance_properties(
                revit.doc, 
                source_category,
                use_linked_document=use_linked,
                linked_document_name=linked_doc_name
            )
            
            # Populate ComboBox
            current_selection = self.sourceSortPropertyComboBox.SelectedItem
            for prop in properties:
                self.sourceSortPropertyComboBox.Items.Add(prop)
            
            # Restore previous selection if it still exists, otherwise select ElementId
            if current_selection and current_selection in properties:
                self.sourceSortPropertyComboBox.SelectedItem = current_selection
            else:
                self.sourceSortPropertyComboBox.SelectedItem = "ElementId"
        except Exception as e:
            logger.error("Error loading source element properties: {}".format(str(e)))
            # Fallback to ElementId
            self.sourceSortPropertyComboBox.Items.Clear()
            self.sourceSortPropertyComboBox.Items.Add("ElementId")
            self.sourceSortPropertyComboBox.SelectedItem = "ElementId"
    
    def source_sort_property_selection_changed(self, sender, args):
        """Handle source sort property selection changed event."""
        # No validation needed, this is optional
        pass
    
    def parameter_mappings_collection_changed(self, sender, args):
        """Handle parameter mappings collection changed event."""
        self.validate_form()
    
    def use_linked_document_checked(self, sender, args):
        """Handle use linked document checkbox checked event."""
        self.linkedDocumentComboBox.IsEnabled = True
        self.validate_form()
        # Reload source element properties when linked document option changes
        if self.sourceCategoriesListBox.SelectedItem:
            self.load_source_element_properties()
    
    def use_linked_document_unchecked(self, sender, args):
        """Handle use linked document checkbox unchecked event."""
        self.linkedDocumentComboBox.IsEnabled = False
        self.linkedDocumentComboBox.SelectedItem = None
        self.validate_form()
        # Reload source element properties when linked document option changes
        if self.sourceCategoriesListBox.SelectedItem:
            self.load_source_element_properties()
    
    def linked_document_selection_changed(self, sender, args):
        """Handle linked document selection changed event."""
        self.validate_form()
        # Reload source element properties when linked document changes
        if self.sourceCategoriesListBox.SelectedItem:
            self.load_source_element_properties()
    
    def process_ui_events(self):
        """Process UI events to update the display."""
        try:
            System.Windows.Forms.Application.DoEvents()
            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(
                Action(lambda: None),
                System.Windows.Threading.DispatcherPriority.Background
            )
        except Exception as e:
            logger.debug("Error processing UI events: {}".format(str(e)))
    
    def config_selection_changed(self, sender, args):
        """Handle configuration selection changed."""
        has_selection = self.configsListView.SelectedItem is not None
        self.editButton.IsEnabled = has_selection
        self.deleteButton.IsEnabled = has_selection
        
        # Enable/disable move buttons based on selection and position
        if has_selection:
            selected_index = self.configsListView.SelectedIndex
            self.moveUpButton.IsEnabled = selected_index > 0
            self.moveDownButton.IsEnabled = selected_index < len(self.configs) - 1
        else:
            self.moveUpButton.IsEnabled = False
            self.moveDownButton.IsEnabled = False
    
    def add_button_click(self, sender, args):
        """Handle add button click."""
        self.is_new_config = True
        self.current_config = None
        
        # Clear form
        self.nameTextBox.Text = ""
        
        # Load category options
        self.load_category_options()
        
        # Clear selections
        self.sourceCategoriesListBox.SelectedItem = None
        self.targetFilterCategoriesListBox.SelectedItems.Clear()
        self.parameter_mappings.Clear()
        
        # Reset linked document settings
        self.useLinkedDocumentCheckBox.IsChecked = False
        self.linkedDocumentComboBox.SelectedItem = None
        
        # Reset validation state
        self._validation_errors = {
            "name": False,
            "source_category": False,
            "mappings": False,
            "linked_document": False
        }
        
        # Switch to edit tab
        self.editConfigTab.IsEnabled = True
        self.tabControl.SelectedItem = self.editConfigTab
        
        # Ensure parameter mapping buttons are enabled
        self.addParameterMappingButton.IsEnabled = True
        self.editParameterMappingButton.IsEnabled = False
        self.deleteParameterMappingButton.IsEnabled = False
        
        # Validate form to update Save button state
        self.validate_form()
        
        self.update_status("Adding new mapping")
    
    def config_double_click(self, sender, args):
        """Handle double-click on configuration in list."""
        # Double-clicking a configuration should open it for editing
        self.edit_button_click(sender, args)
    
    def edit_button_click(self, sender, args):
        """Handle edit button click."""
        selected_config = self.configsListView.SelectedItem
        
        if not selected_config:
            self.update_status("No configuration selected")
            return
        
        self.is_new_config = False
        self.current_config = selected_config
        
        # Populate form
        self.nameTextBox.Text = selected_config.name
        
        # Load category options
        self.load_category_options()
        
        # Select source category using exact dictionary matching (single selection)
        self.sourceCategoriesListBox.SelectedItem = None
        # Get the first source category from config (since we only support one now)
        if selected_config.source_categories:
            config_cat = selected_config.source_categories[0]
            for display_name, category_value in self.source_category_map.items():
                # Handle special marker
                if config_cat == THREE_D_ZONE_MARKER and category_value == THREE_D_ZONE_MARKER:
                    # Find the ListBox item and select it
                    for item in self.sourceCategoriesListBox.Items:
                        if str(item) == display_name:
                            self.sourceCategoriesListBox.SelectedItem = item
                            break
                    break
                # Handle BuiltInCategory - compare by integer value for exact match
                elif isinstance(config_cat, BuiltInCategory) and isinstance(category_value, BuiltInCategory):
                    if int(config_cat) == int(category_value):
                        # Find the ListBox item and select it
                        for item in self.sourceCategoriesListBox.Items:
                            if str(item) == display_name:
                                self.sourceCategoriesListBox.SelectedItem = item
                                break
                        break
        
        # Select target filter categories using exact dictionary matching
        self.targetFilterCategoriesListBox.SelectedItems.Clear()
        for display_name, category_value in self.target_category_map.items():
            # Check if this category is in the selected config's target filter categories
            for config_cat in selected_config.target_filter_categories:
                # Handle BuiltInCategory - compare by integer value for exact match
                if isinstance(config_cat, BuiltInCategory) and isinstance(category_value, BuiltInCategory):
                    if int(config_cat) == int(category_value):
                        # Find the ListBox item and select it
                        for item in self.targetFilterCategoriesListBox.Items:
                            if str(item) == display_name:
                                self.targetFilterCategoriesListBox.SelectedItems.Add(item)
                                break
                        break
        
        # Load parameter mappings
        self.parameter_mappings.Clear()
        for src_param, tgt_param in zip(selected_config.source_params, selected_config.target_params):
            self.parameter_mappings.Add(ParameterMappingEntry(
                SourceParameter=src_param,
                TargetParameter=tgt_param
            ))
        
        # Restore linked document settings
        use_linked = selected_config.use_linked_document if hasattr(selected_config, 'use_linked_document') else False
        linked_doc_name = selected_config.linked_document_name if hasattr(selected_config, 'linked_document_name') else None
        
        self.useLinkedDocumentCheckBox.IsChecked = use_linked
        self.linkedDocumentComboBox.IsEnabled = use_linked
        
        # Try to select the configured linked document
        if linked_doc_name:
            # Check if the linked document still exists
            link_instance = get_linked_document_by_name(revit.doc, linked_doc_name)
            if link_instance:
                # Find and select the item in the dropdown
                for item in self.linkedDocumentComboBox.Items:
                    if str(item) == linked_doc_name:
                        self.linkedDocumentComboBox.SelectedItem = item
                        break
            else:
                # Linked document not found - will show validation error
                self.linkedDocumentComboBox.SelectedItem = None
        else:
            self.linkedDocumentComboBox.SelectedItem = None
        
        # Load sort property (trigger property loading first if category is selected)
        if self.sourceCategoriesListBox.SelectedItem:
            self.load_source_element_properties()
            # Set the sort property value
            sort_property = selected_config.source_sort_property if hasattr(selected_config, 'source_sort_property') else "ElementId"
            if sort_property in self.sourceSortPropertyComboBox.Items:
                self.sourceSortPropertyComboBox.SelectedItem = sort_property
            else:
                self.sourceSortPropertyComboBox.SelectedItem = "ElementId"
        else:
            # No category selected, just set default
            self.sourceSortPropertyComboBox.Items.Clear()
            self.sourceSortPropertyComboBox.Items.Add("ElementId")
            self.sourceSortPropertyComboBox.SelectedItem = "ElementId"
        
        # Switch to edit tab
        self.editConfigTab.IsEnabled = True
        self.tabControl.SelectedItem = self.editConfigTab
        
        # Ensure parameter mapping buttons are enabled
        self.addParameterMappingButton.IsEnabled = True
        # Edit/Delete buttons will be enabled/disabled based on DataGrid selection
        
        # Validate form to update Save button state
        self.validate_form()
        
        self.update_status("Editing mapping: {}".format(selected_config.name))
    
    def load_category_options(self):
        """Load category options into list boxes."""
        # Load source categories
        self.sourceCategoriesListBox.Items.Clear()
        self.source_category_map = {}
        category_options = get_category_options()
        for opt in category_options:
            display_name = opt[0]
            category_value = opt[1]
            self.sourceCategoriesListBox.Items.Add(display_name)
            self.source_category_map[display_name] = category_value
        
        # Load target filter categories
        self.targetFilterCategoriesListBox.Items.Clear()
        self.target_category_map = {}
        all_model_cats = get_all_model_categories(revit.doc)
        for opt in all_model_cats:
            display_name = opt[0]
            category_value = opt[1]
            self.targetFilterCategoriesListBox.Items.Add(display_name)
            self.target_category_map[display_name] = category_value
        
        # Load linked documents
        self.linkedDocumentComboBox.Items.Clear()
        linked_docs = get_linked_documents(revit.doc)
        for link_name, link_instance in linked_docs:
            self.linkedDocumentComboBox.Items.Add(link_name)
    
    def save_button_click(self, sender, args):
        """Handle save button click."""
        try:
            # Validate form before saving
            if not self.validate_form():
                # Build error message
                errors = []
                if self._validation_errors["name"]:
                    errors.append("Mapping name is required")
                if self._validation_errors["source_category"]:
                    errors.append("A source category must be selected")
                if self._validation_errors["mappings"]:
                    errors.append("At least one parameter mapping is required")
                if self._validation_errors["linked_document"]:
                    errors.append("Linked document must be selected when 'Use Linked Document' is checked")
                
                error_message = "Please fix the following errors:\n\n" + "\n".join("- " + e for e in errors)
                MessageBox.Show(error_message, "Validation Error", MessageBoxButton.OK)
                return
            
            # Get validated inputs
            name = self.nameTextBox.Text.strip()

            # Get order - use existing order if editing, otherwise get next order
            if self.is_new_config:
                order = config.get_next_order(revit.doc)
            else:
                if self.current_config is None:
                    # This shouldn't happen if edit_button_click was called properly
                    # Try to find the config by name as a fallback
                    all_configs_temp = config.load_configs(revit.doc)
                    found_config = None
                    for cfg in all_configs_temp:
                        if cfg.get("name") == name:
                            found_config = cfg
                            break
                    
                    if found_config:
                        # Create a temporary ConfigItem to get the order
                        temp_config_item = ConfigItem(found_config)
                        order = temp_config_item.order
                    else:
                        # Last resort: create new config
                        self.is_new_config = True
                        order = config.get_next_order(revit.doc)
                else:
                    order = self.current_config.order

            # Get selected source category (single selection)
            selected_source_cats = []
            category_options = get_category_options()
            cat_map = {opt[0]: opt[1] for opt in category_options}
            
            if self.sourceCategoriesListBox.SelectedItem:
                display_name = str(self.sourceCategoriesListBox.SelectedItem)
                cat_value = cat_map.get(display_name)
                if cat_value:
                    selected_source_cats.append(cat_value)
            
            # Get parameter mappings
            source_params = []
            target_params = []
            for mapping in self.parameter_mappings:
                src = mapping.SourceParameter.strip()
                tgt = mapping.TargetParameter.strip()
                if src and tgt:
                    source_params.append(src)
                    target_params.append(tgt)
            
            # Additional validation check (should not happen if validate_form passed)
            if len(source_params) != len(target_params):
                MessageBox.Show("Source and target parameters count must match.", "Validation Error", MessageBoxButton.OK)
                return
            
            # Get selected target filter categories
            selected_target_filter_cats = []
            
            for item in self.targetFilterCategoriesListBox.SelectedItems:
                display_name = str(item)
                cat_value = self.target_category_map.get(display_name)
                
                if cat_value:
                    selected_target_filter_cats.append(cat_value)
            
            # Get linked document settings
            use_linked_document = self.useLinkedDocumentCheckBox.IsChecked
            linked_document_name = None
            if use_linked_document and self.linkedDocumentComboBox.SelectedItem:
                linked_document_name = str(self.linkedDocumentComboBox.SelectedItem)
            
            # Get sort property setting (default to ElementId if not selected)
            source_sort_property = "ElementId"
            if self.sourceSortPropertyComboBox.SelectedItem:
                source_sort_property = str(self.sourceSortPropertyComboBox.SelectedItem)
            
            # Get all configurations
            all_configs = config.load_configs(revit.doc)
            
            if self.is_new_config:
                # Create new configuration
                new_config = {
                    "id": config.generate_config_id(),
                    "name": name,
                    "order": order,
                    "enabled": False,  # Default to False, user can enable in ListView
                    "source_categories": selected_source_cats,
                    "source_params": source_params,
                    "target_params": target_params,
                    "target_filter_categories": selected_target_filter_cats,
                    "write_before_ifc_export": False,  # Default to False, user can enable in ListView
                    "ifc_export_only_empty": False,  # Default to False, user can enable in ListView
                    "use_linked_document": use_linked_document,
                    "linked_document_name": linked_document_name,
                    "source_sort_property": source_sort_property
                }
                all_configs.append(new_config)
            else:
                # Update existing configuration
                if self.current_config is None:
                    # Try to find the config we're editing by matching name
                    config_found = False
                    for i, cfg in enumerate(all_configs):
                        # Match by name (assuming user didn't change the name)
                        if cfg.get("name") == name:
                            # Found it - update this one
                            # Use the found config's ID and order
                            config_id = cfg.get("id")
                            found_order = cfg.get("order", order)
                            all_configs[i] = {
                                "id": config_id,
                                "name": name,
                                "order": found_order,
                                "enabled": cfg.get("enabled", False),
                                "source_categories": selected_source_cats,
                                "source_params": source_params,
                                "target_params": target_params,
                                "target_filter_categories": selected_target_filter_cats,
                                "write_before_ifc_export": cfg.get("write_before_ifc_export", False),
                                "ifc_export_only_empty": cfg.get("ifc_export_only_empty", False),
                                "use_linked_document": use_linked_document,
                                "linked_document_name": linked_document_name,
                                "source_sort_property": source_sort_property
                            }
                            config_found = True
                            break
                    
                    if not config_found:
                        # Last resort: create new config
                        new_config = {
                            "id": config.generate_config_id(),
                            "name": name,
                            "order": order,
                            "enabled": False,
                            "source_categories": selected_source_cats,
                            "source_params": source_params,
                            "target_params": target_params,
                            "target_filter_categories": selected_target_filter_cats,
                            "write_before_ifc_export": False,
                            "ifc_export_only_empty": False,
                            "use_linked_document": use_linked_document,
                            "linked_document_name": linked_document_name
                        }
                        all_configs.append(new_config)
                else:
                    config_id = self.current_config.id
                    for i, cfg in enumerate(all_configs):
                        if cfg.get("id") == config_id:
                            # Preserve checkbox values from current_config (ConfigItem)
                            all_configs[i] = {
                                "id": config_id,
                                "name": name,
                                "order": order,
                                "enabled": self.current_config.enabled if self.current_config else False,
                                "source_categories": selected_source_cats,
                                "source_params": source_params,
                                "target_params": target_params,
                                "target_filter_categories": selected_target_filter_cats,
                                "write_before_ifc_export": self.current_config.write_before_ifc_export if self.current_config else False,
                                "ifc_export_only_empty": self.current_config.ifc_export_only_empty if self.current_config else False,
                                "use_linked_document": use_linked_document,
                                "linked_document_name": linked_document_name,
                                "source_sort_property": source_sort_property
                            }
                            break
            
            # Save configurations
            save_result = config.save_configs(revit.doc, all_configs)
            
            if save_result:
                # Reload configurations
                self.load_configurations()
                
                # Go back to configs tab
                self.back_button_click(None, None)
                
                self.update_status("Mapping saved successfully")
            else:
                self.update_status("Error saving mapping")
                MessageBox.Show("Failed to save mapping.", "Error", MessageBoxButton.OK)
            
        except Exception as e:
            logger.error("Error saving configuration: {}".format(str(e)))
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))
            self.update_status("Error saving mapping: {}".format(str(e)))
            MessageBox.Show("Error saving mapping: {}".format(str(e)), "Error", MessageBoxButton.OK)
    
    def delete_button_click(self, sender, args):
        """Handle delete button click."""
        selected_config = self.configsListView.SelectedItem
        
        if not selected_config:
            self.update_status("No configuration selected")
            return
        
        # Confirm deletion
        result = MessageBox.Show(
            "Are you sure you want to delete this mapping?\n\n{}".format(selected_config.name),
            "Confirm Deletion",
            MessageBoxButton.YesNo
        )
        
        if result != MessageBoxResult.Yes:
            return
        
        try:
            config_id = selected_config.id
            if config.delete_config(revit.doc, config_id):
                # Reload configurations
                self.load_configurations()
                self.update_status("Mapping deleted successfully")
            else:
                self.update_status("Error deleting mapping")
                MessageBox.Show("Failed to delete mapping.", "Error", MessageBoxButton.OK)
        
        except Exception as e:
            logger.error("Error deleting configuration: {}".format(str(e)))
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))
            self.update_status("Error deleting mapping: {}".format(str(e)))
            MessageBox.Show("Error deleting mapping: {}".format(str(e)), "Error", MessageBoxButton.OK)
    
    def tab_selection_changed(self, sender, args):
        """Handle tab selection change."""
        # If switching from edit tab back to configs tab, reset the form
        if self.tabControl.SelectedItem == self.configsTab:
            # Check if we were in edit mode
            if self.current_config is not None or self.is_new_config:
                # Reset the form like Back button would
                self.back_button_click(None, None)
    
    def back_button_click(self, sender, args):
        """Handle back button click."""
        # Reset current configuration
        self.current_config = None
        self.is_new_config = False
        
        # Clear mappings
        self.parameter_mappings.Clear()
        
        # Reset validation state
        self._validation_errors = {
            "name": False,
            "source_category": False,
            "mappings": False
        }
        
        # Update UI
        self.editConfigTab.IsEnabled = False
        self.tabControl.SelectedItem = self.configsTab
        self.update_status("Returned to mappings list")
    
    def add_parameter_mapping_button_click(self, sender, args):
        """Handle add parameter mapping button click."""
        # Get available parameters
        available_params = get_all_writable_instance_parameters(revit.doc)
        
        if not available_params:
            MessageBox.Show("No writable instance parameters found in the project.", "No Parameters", MessageBoxButton.OK)
            return
        
        # Create and show custom parameter selector dialog
        selector_dialog = ParameterSelectorDialog(available_params)
        result = selector_dialog.ShowDialog()
        
        if result and selector_dialog.selected_source and selector_dialog.selected_target:
            # Add new mapping entry
            new_mapping = ParameterMappingEntry(
                SourceParameter=selector_dialog.selected_source,
                TargetParameter=selector_dialog.selected_target
            )
            self.parameter_mappings.Add(new_mapping)
            
            # Validate form to update Save button state
            self.validate_form()
            
            # Update status
            self.update_status("Added parameter mapping: {} -> {}".format(
                selector_dialog.selected_source, 
                selector_dialog.selected_target
            ))
    
    def parameter_mapping_selection_changed(self, sender, args):
        """Handle selection change in parameter mapping DataGrid."""
        has_selection = self.parameterMappingsDataGrid.SelectedItem is not None
        self.editParameterMappingButton.IsEnabled = has_selection
        self.deleteParameterMappingButton.IsEnabled = has_selection
    
    def edit_parameter_mapping_button_click(self, sender, args):
        """Handle edit parameter mapping button click."""
        selected_mapping = self.parameterMappingsDataGrid.SelectedItem
        if not selected_mapping:
            MessageBox.Show("Please select a parameter mapping to edit.", "No Selection", MessageBoxButton.OK)
            return
        
        self.edit_parameter_mapping(selected_mapping)
    
    def delete_parameter_mapping_button_click(self, sender, args):
        """Handle delete parameter mapping button click."""
        selected_mapping = self.parameterMappingsDataGrid.SelectedItem
        if not selected_mapping:
            MessageBox.Show("Please select a parameter mapping to delete.", "No Selection", MessageBoxButton.OK)
            return
        
        self.delete_parameter_mapping(selected_mapping)
    
    def edit_parameter_mapping(self, mapping_entry):
        """Edit a parameter mapping entry."""
        # Get current values
        current_source = mapping_entry.SourceParameter.strip() if mapping_entry.SourceParameter else None
        current_target = mapping_entry.TargetParameter.strip() if mapping_entry.TargetParameter else None
        
        # Get available parameters
        available_params = get_all_writable_instance_parameters(revit.doc)
        
        if not available_params:
            MessageBox.Show("No writable instance parameters found in the project.", "No Parameters", MessageBoxButton.OK)
            return
        
        # Create and show custom parameter selector dialog with current values preselected
        selector_dialog = ParameterSelectorDialog(
            available_params,
            initial_source=current_source,
            initial_target=current_target
        )
        result = selector_dialog.ShowDialog()
        
        if result and selector_dialog.selected_source and selector_dialog.selected_target:
            # Update the mapping entry
            mapping_entry.SourceParameter = selector_dialog.selected_source
            mapping_entry.TargetParameter = selector_dialog.selected_target
            
            # Validate form to update Save button state
            self.validate_form()
            
            # Update status
            self.update_status("Updated parameter mapping: {} -> {}".format(
                selector_dialog.selected_source, 
                selector_dialog.selected_target
            ))
    
    def delete_parameter_mapping(self, mapping_entry):
        """Delete a parameter mapping entry."""
        if mapping_entry in self.parameter_mappings:
            self.parameter_mappings.Remove(mapping_entry)
            # Validate form to update Save button state
            self.validate_form()
            self.update_status("Deleted parameter mapping")
    
    def move_up_button_click(self, sender, args):
        """Move selected configuration up in order."""
        selected_index = self.configsListView.SelectedIndex
        if selected_index <= 0:
            return
        
        try:
            all_configs = config.load_configs(revit.doc)
            sorted_configs = sorted(all_configs, key=lambda x: x.get("order", 999))
            
            # Swap orders
            config1 = sorted_configs[selected_index - 1]
            config2 = sorted_configs[selected_index]
            
            temp_order = config1.get("order")
            config1["order"] = config2.get("order")
            config2["order"] = temp_order
            
            # Save
            if config.save_configs(revit.doc, all_configs):
                self.load_configurations()
                # Reselect the moved item
                self.configsListView.SelectedIndex = selected_index - 1
                self.update_status("Mapping moved up")
            else:
                MessageBox.Show("Failed to reorder mappings.", "Error", MessageBoxButton.OK)
        
        except Exception as e:
            logger.error("Error moving mapping up: {}".format(str(e)))
            MessageBox.Show("Error moving mapping: {}".format(str(e)), "Error", MessageBoxButton.OK)
    
    def move_down_button_click(self, sender, args):
        """Move selected configuration down in order."""
        selected_index = self.configsListView.SelectedIndex
        if selected_index >= len(self.configs) - 1:
            return
        
        try:
            all_configs = config.load_configs(revit.doc)
            sorted_configs = sorted(all_configs, key=lambda x: x.get("order", 999))
            
            # Swap orders
            config1 = sorted_configs[selected_index]
            config2 = sorted_configs[selected_index + 1]
            
            temp_order = config1.get("order")
            config1["order"] = config2.get("order")
            config2["order"] = temp_order
            
            # Save
            if config.save_configs(revit.doc, all_configs):
                self.load_configurations()
                # Reselect the moved item
                self.configsListView.SelectedIndex = selected_index + 1
                self.update_status("Mapping moved down")
            else:
                MessageBox.Show("Failed to reorder mappings.", "Error", MessageBoxButton.OK)
        
        except Exception as e:
            logger.error("Error moving mapping down: {}".format(str(e)))
            MessageBox.Show("Error moving mapping: {}".format(str(e)), "Error", MessageBoxButton.OK)

# --- Main Execution ---

if __name__ == '__main__':
    # Show the Configuration Editor UI
    Zone3DConfigEditorUI().ShowDialog()

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.
