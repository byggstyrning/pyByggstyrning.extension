# -*- coding: utf-8 -*-
__title__ = "Edit Configs"
__author__ = "Byggstyrning AB"
__doc__ = """Edit StreamBIM checklist configurations stored in extensible storage.

This tool allows you to view, edit, and delete mapping configurations
that were created by the Checklist Importer tool."""

import os
import sys
import clr
import json
import imp
from collections import namedtuple
import pickle
import base64

# Add the extension directory to the path - FIXED PATH RESOLUTION
import os.path as op
script_path = __file__
script_dir = op.dirname(script_path)
stack_dir = op.dirname(script_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Import the streambim_api module
from streambim import streambim_api

# Import StreamBIMSettingsSchema and related functions directly from the module
from streambim.streambim_api import StreamBIMSettingsSchema
from streambim.streambim_api import get_or_create_settings_storage
from streambim.streambim_api import load_configs_with_pickle
from streambim.streambim_api import save_configs_with_pickle
from streambim.streambim_api import get_saved_project_id

# Try direct import from current directory's parent path
sys.path.append(op.dirname(op.dirname(panel_dir)))

# Add reference to WPF
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import *

from System import EventHandler, Action
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import MessageBox, MessageBoxButton, MessageBoxResult, Visibility
from System.Windows.Controls import SelectionMode
import System.Windows.Forms
import System.Windows.Threading

from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Import extensible storage
from extensible_storage import BaseSchema, simple_field

# Import revit_utils functions
from revit.revit_utils import get_element_by_ifc_guid

# Initialize logger
logger = script.get_logger()

# Define a Configuration class for displaying in the list
class ConfigItem(object):
    def __init__(self, id=None, checklist_id=None, checklist_name=None, streambim_property=None, revit_parameter=None, mapping_enabled=False, mapping_config=None):
        self.id = id  # Use the data storage element ID
        self.checklist_id = checklist_id
        self.checklist_name = checklist_name
        self.streambim_property = streambim_property
        self.revit_parameter = revit_parameter
        
        # Make sure mapping_enabled is treated as a boolean
        if isinstance(mapping_enabled, bool):
            self.mapping_enabled = mapping_enabled
        elif isinstance(mapping_enabled, str):
            self.mapping_enabled = mapping_enabled.lower() == "true"
        else:
            self.mapping_enabled = bool(mapping_enabled)
            
        self.mapping_config = mapping_config
        self.elements_total = 0
        self.elements_processed = 0
        self.elements_updated = 0
        self.mapping_count = 0
        if mapping_config:
            try:
                mapping_data = json.loads(mapping_config)
                self.mapping_count = len(mapping_data)
            except:
                self.mapping_count = 0
    
    @property
    def DisplayName(self):
        return "{} -> {}".format(self.streambim_property, self.revit_parameter)
    
    @property
    def MappingDetails(self):
        if self.mapping_enabled and self.mapping_count > 0:
            return "{} value mappings".format(self.mapping_count)
        return "No mappings"
    
    @property
    def Status(self):
        if self.elements_total == 0:
            return "Not processed"
        else:
            return "{}/{} elements processed, {} updated".format(
                self.elements_processed, 
                self.elements_total,
                self.elements_updated
            )
        
    @property 
    def ChecklistName(self):
        return self.checklist_name or "Unknown Checklist"

# Define a MappingEntry class for the mapping DataGrid
class MappingEntry(object):
    def __init__(self, ChecklistValue="", RevitValue=""):
        self._checklist_value = ChecklistValue
        self._revit_value = RevitValue

    @property
    def ChecklistValue(self):
        return self._checklist_value

    @ChecklistValue.setter
    def ChecklistValue(self, value):
        self._checklist_value = value

    @property
    def RevitValue(self):
        return self._revit_value

    @RevitValue.setter
    def RevitValue(self, value):
        self._revit_value = value

# ensure_styles_loaded() is now imported from lib.styles

class ConfigEditorUI(forms.WPFWindow):
    """Configuration Editor UI implementation."""
    
    def __init__(self):
        """Initialize the Configuration Editor UI."""
        # Load styles into Application.Resources BEFORE creating window
        from styles import ensure_styles_loaded
        ensure_styles_loaded()
        
        # Initialize WPF window
        forms.WPFWindow.__init__(self, 'ConfigEditor.xaml')
        
        # Initialize StreamBIM API client
        self.streambim_client = streambim_api.StreamBIMClient()

        # Initialize data collections
        self.configs = ObservableCollection[object]()
        self.mappings = ObservableCollection[object]()
        self.current_config = None
        
        # Set up event handlers
        self.configsListView.SelectionChanged += self.config_selection_changed
        self.configsListView.SelectionMode = SelectionMode.Extended  # Allow multiple selection for batch import
        self.editButton.Click += self.edit_button_click
        self.saveButton.Click += self.save_button_click
        self.deleteButton.Click += self.delete_button_click
        self.backButton.Click += self.back_button_click
        self.addNewRowButton.Click += self.add_new_row_button_click
        self.removeRowButton.Click += self.remove_row_button_click
        self.enableMappingCheckBox.Checked += self.enable_mapping_checked
        self.enableMappingCheckBox.Unchecked += self.enable_mapping_unchecked
        self.runAllButton.Click += self.run_all_button_click
        self.runSelectedButton.Click += self.run_selected_button_click
        
        # Initialize UI
        self.configsListView.ItemsSource = self.configs
        self.mappingDataGrid.ItemsSource = self.mappings
        
        # Initial UI state
        self.tabControl.SelectedItem = self.configsTab
        self.editConfigTab.IsEnabled = False
        
        # Load configurations
        self.load_configurations()

        # Try automatic login (but don't show dialog if it fails - check before window creation)
        self.try_automatic_login()
    
    def load_styles(self):
        """Load the common styles ResourceDictionary."""
        try:
            import os.path as op
            script_dir = op.dirname(__file__)
            stack_dir = op.dirname(script_dir)
            panel_dir = op.dirname(stack_dir)
            tab_dir = op.dirname(panel_dir)
            extension_dir = op.dirname(tab_dir)
            styles_path = op.join(extension_dir, 'lib', 'styles', 'CommonStyles.xaml')
            
            if op.exists(styles_path):
                from System.Windows import Application
                from System.Windows.Markup import XamlReader
                from System.IO import File
                
                # Read XAML content
                xaml_content = File.ReadAllText(styles_path)
                
                # Parse as ResourceDictionary
                styles_dict = XamlReader.Parse(xaml_content)
                
                # Merge into window resources
                if self.Resources is None:
                    from System.Windows import ResourceDictionary
                    self.Resources = ResourceDictionary()
                
                # If it's a ResourceDictionary, merge its contents
                if hasattr(styles_dict, 'Keys'):
                    for key in styles_dict.Keys:
                        self.Resources[key] = styles_dict[key]
                else:
                    # Try to merge the entire dictionary
                    self.Resources.MergedDictionaries.Add(styles_dict)
                    
        except Exception as e:
            logger.warning("Could not load styles: {}. Using default styles.".format(str(e)))
            import traceback

    def set_busy(self, is_busy, message="Loading..."):
        """Show or hide the busy overlay indicator."""
        try:
            if is_busy:
                self.busyOverlay.Visibility = Visibility.Visible
                self.busyTextBlock.Text = message
            else:
                self.busyOverlay.Visibility = Visibility.Collapsed
        except Exception as e:
            pass
    
    def try_automatic_login(self):
        """Attempt to automatically log in using saved tokens."""
        # Load tokens from file first
        self.streambim_client.load_tokens()
        
        # Check if token exists
        if self.streambim_client.idToken:
            # Try to load saved project ID
            saved_project_id = get_saved_project_id(revit.doc)
            if saved_project_id:
                self.streambim_client.set_current_project(saved_project_id)
            
            return True
        else:
            return False
        
    def load_configurations(self):
        """Load all mapping configurations from storage."""
        # Show busy indicator during loading
        self.set_busy(True, "Loading configurations...")
        
        try:
            # Clear existing configurations
            self.configs.Clear()
            
            # Load configurations from consolidated storage
            loaded_configs = load_configs_with_pickle(revit.doc)
            
            if loaded_configs:
                # Add them to the observable collection
                for config in loaded_configs:
                    config_item = ConfigItem(
                        id=None,  # No element ID for new configurations
                        checklist_id=config.get('checklist_id'),
                        checklist_name=config.get('checklist_name'),
                        streambim_property=config.get('streambim_property'),
                        revit_parameter=config.get('revit_parameter'),
                        mapping_enabled=config.get('mapping_enabled', False),
                        mapping_config=config.get('mapping_config')
                    )
                    self.configs.Add(config_item)
                
            # Force the UI to refresh the configurations list
            self.refresh_ui()
            
        except Exception as e:
            logger.error("Error loading configurations: {}".format(str(e)))
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))
        finally:
            # Hide busy indicator
            self.set_busy(False)

    def run_all_button_click(self, sender, args):
        """Process all configs when the Run All button is clicked."""
        # Confirm before proceeding
        result = MessageBox.Show(
            "This will process all configurations and update all elements with matching IfcGUID parameters. Continue?",
            "Confirm Batch Import",
            MessageBoxButton.YesNo
        )
        
        # Check if user clicked Yes
        if result != MessageBoxResult.Yes:
            return
        
        # Initialize StreamBIM API client
        self.api_client = streambim_api.StreamBIMClient()
        
        # Get last used project ID or ask user to log in
        project_id = get_saved_project_id(revit.doc)
        if not project_id or not self.api_client.idToken:
            # Ask user to log in
            login_result = self.show_login_dialog()
            if not login_result:
                return
                
        # Set the project ID in the API client
        if project_id:
            self.api_client.set_current_project(project_id)
        
        # Process all configurations
        self.run_import_configurations(self.configs)
    
    def run_selected_button_click(self, sender, args):
        """Process selected configs when the Run Selected button is clicked."""
        if not self.configsListView.SelectedItems or self.configsListView.SelectedItems.Count == 0:
            return
            
        # Confirm before proceeding
        result = MessageBox.Show(
            "This will process the selected configurations and update elements with matching IfcGUID parameters. Continue?",
            "Confirm Batch Import",
            MessageBoxButton.YesNo
        )
        
        # Check if user clicked Yes
        if result != MessageBoxResult.Yes:
            return
            
        # Initialize StreamBIM API client
        self.api_client = streambim_api.StreamBIMClient()
        
        # Get last used project ID or ask user to log in
        project_id = get_saved_project_id(revit.doc)
        if not project_id or not self.api_client.idToken:
            # Ask user to log in
            login_result = self.show_login_dialog()
            if not login_result:
                return
                
        # Set the project ID in the API client
        if project_id:
            self.api_client.set_current_project(project_id)
        
        # Convert selected items to a list
        selected_configs = []
        for item in self.configsListView.SelectedItems:
            selected_configs.append(item)
            
        # Process selected configurations
        self.run_import_configurations(selected_configs)

    def show_login_dialog(self):
        """Show a dialog to log in to StreamBIM."""
        login_ui = forms.LoginWindow('StreamBIM Login')
        login_ui.ShowDialog()
        
        if login_ui.DialogResult:
            username = login_ui.username_tb.Text
            password = login_ui.password_tb.Password
            
            if not username or not password:
                return False
                
            # Login to StreamBIM
            login_result = self.api_client.login(username, password)
            
            # Handle new MFA-enabled login response (dict) or legacy boolean
            login_success = False
            if isinstance(login_result, dict):
                if login_result.get('success'):
                    login_success = True
                elif login_result.get('requires_mfa'):
                    # MFA required - show message
                    MessageBox.Show(
                        "This account requires MFA (Multi-Factor Authentication).\n\nPlease use the Checklist Importer tool to log in with MFA support.",
                        "MFA Required",
                        MessageBoxButton.OK
                    )
                    return False
                else:
                    return False
            elif login_result:
                # Legacy boolean success
                login_success = True
            else:
                return False
            
            if login_success:
                
                # Get projects
                projects = self.api_client.get_projects()
                
                if not projects or len(projects) == 0:
                    return False
                
                # Use the first project if none is specified
                if not self.api_client.current_project:
                    first_project_id = projects[0].get('id')
                    self.api_client.set_current_project(first_project_id)
                    
                    # Save the project ID
                    self.save_project_id(first_project_id)
                
                return True
            else:
                return False
        else:
            return False

    def save_project_id(self, project_id):
        """Save the StreamBIM project ID to extensible storage."""
        try:
            # Get data storage
            data_storage = get_or_create_settings_storage(revit.doc)
            
            if not data_storage:
                logger.error("Failed to get or create StreamBIM storage")
                return
                
            # Get current stored value (if any)
            schema = StreamBIMSettingsSchema(data_storage)
            current_id = schema.get("project_id")
            
            # Only update if different
            if current_id != project_id:
                with revit.Transaction("Save StreamBIM Project ID", revit.doc):
                    with StreamBIMSettingsSchema(data_storage) as entity:
                        entity.set("project_id", project_id)
        except Exception as e:
            logger.error("Error saving project ID: {}".format(str(e)))


    def run_import_configurations(self, configs):
        """Run import for a list of configurations."""
        # Disable buttons during processing
        self.runAllButton.IsEnabled = False
        self.runSelectedButton.IsEnabled = False
        
        # Show busy indicator during batch import
        self.set_busy(True, "Processing batch import...")
        
        try:
            # Update main progress bar max
            self.mainProgressBar.Maximum = len(configs)
            self.mainProgressBar.Value = 0
            
            # Track total elements processed and updated
            total_processed = 0
            total_updated = 0
            
            # Process each configuration separately
            for i, config in enumerate(configs):
                # Update main progress bar
                self.mainProgressBar.Value = i
                
                # Process UI events
                self.process_ui_events()
                
                # Skip configurations without checklist ID
                if not config.checklist_id:
                    config.elements_processed = 0
                    config.elements_updated = 0
                    continue
                
                # Process this configuration in its own transaction
                processed_count, updated_count = self.process_single_configuration(config, i, len(configs))
                
                # Update totals
                total_processed += processed_count
                total_updated += updated_count
                
            
            # Complete the main progress bar
            self.mainProgressBar.Value = len(configs)
            
            MessageBox.Show(
                "Batch import completed.\n\nProcessed {} configurations.\nUpdated {} out of {} elements.".format(
                    len(configs), total_updated, total_processed),
                "Batch Import Results",
                MessageBoxButton.OK
            )
            
        except Exception as e:
            logger.error("Error running batch import: {}".format(str(e)))
            # Log detailed stack trace
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))
            
            # Show error message box
            MessageBox.Show(
                "An error occurred during batch import:\n\n{}".format(str(e)),
                "Batch Import Error",
                MessageBoxButton.OK
            )
            
        finally:
            # Hide busy indicator
            self.set_busy(False)
            # Re-enable buttons
            self.runAllButton.IsEnabled = True
            self.runSelectedButton.IsEnabled = True
    
    def update_config_progress(self, config):
        """Update the progress display for a configuration."""
        # Find the list view item for this config
        for i in range(self.configsListView.Items.Count):
            item = self.configsListView.Items[i]
            if item == config:
                # Force the ListView to refresh this item
                # This updates the display of the Status property
                container = self.configsListView.ItemContainerGenerator.ContainerFromIndex(i)
                if container:
                    container.UpdateLayout()
                break
        
        # Process UI events to update the display
        self.process_ui_events()
    
    def process_ui_events(self):
        """Process UI events to update the display."""
        try:
            # Use the Application.DoEvents to process pending UI events
            System.Windows.Forms.Application.DoEvents()
            
            # Use the Dispatcher
            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(
                Action(lambda: None),
                System.Windows.Threading.DispatcherPriority.Background
            )
        except Exception as e:
            # Log any errors but continue execution
            pass
    
    def get_all_elements_with_ifc_guid(self):
        """Get all elements with an IfcGUID parameter."""
        result = []
        
        # Get all elements in the document
        all_elements = FilteredElementCollector(revit.doc).WhereElementIsNotElementType().ToElements()
        
        # Filter for elements with IfcGUID parameter
        for element in all_elements:
            try:
                # Check for different variations of the parameter name
                ifc_guid_param = element.LookupParameter("IFCGuid")
                if not ifc_guid_param:
                    ifc_guid_param = element.LookupParameter("IfcGUID")
                if not ifc_guid_param:
                    ifc_guid_param = element.LookupParameter("IFC GUID")
                
                if ifc_guid_param:
                    result.append(element)
            except:
                continue
        
        return result
    
    
    def update_status(self, message):
        """Update the status display."""
        # Status bar removed from UI - method kept for compatibility but does nothing
        pass

    def get_property_value(self, checklist_item, property_name):
        """Get property value from checklist item.
        Check both attributes.properties and items paths in the JSON structure."""
        try:
            # First try the attributes.properties path
            props = checklist_item.get('attributes', {}).get('properties', {})
            if props and property_name in props:
                return props.get(property_name)
            
            # Then try the items path
            if 'items' in checklist_item and property_name in checklist_item['items']:
                return checklist_item['items'][property_name]
                
            return None
        except Exception as e:
            logger.error("Error getting property value: {}".format(str(e)))
            return None
            
    def set_parameter_value(self, param, value, storage_type):
        """Set parameter value based on storage type."""
        try:
            if param.IsReadOnly:
                return False
                
            if storage_type == StorageType.String:
                # Convert to string
                str_value = str(value)
                if param.AsString() != str_value:
                    param.Set(str_value)
                    return True
                    
            elif storage_type == StorageType.Integer:
                # Try to convert to integer
                try:
                    int_value = int(value)
                    if param.AsInteger() != int_value:
                        param.Set(int_value)
                        return True
                except (ValueError, TypeError):
                    return False
                    
            elif storage_type == StorageType.Double:
                # Try to convert to double
                try:
                    double_value = float(value)
                    if param.AsDouble() != double_value:
                        param.Set(double_value)
                        return True
                except (ValueError, TypeError):
                    return False
                    
            elif storage_type == StorageType.ElementId:
                # Try to convert to ElementId
                try:
                    int_value = int(value)
                    element_id = ElementId(int_value)
                    if param.AsElementId() != element_id:
                        param.Set(element_id)
                        return True
                except (ValueError, TypeError):
                    return False
                    
            return False
        except Exception as e:
            logger.error("Error setting parameter value: {}".format(str(e)))
            return False

    def process_single_configuration(self, config, config_index, total_configs):
        """Process a single configuration with its own transaction.
        Returns a tuple of (processed_count, updated_count)."""
        processed_count = 0
        updated_count = 0
        
        try:
            # Get checklist items from StreamBIM
            try:
                # Get checklist items with proper error handling
                checklist_items = self.api_client.get_checklist_items(config.checklist_id, config.streambim_property, limit=0)
                if not checklist_items:
                    config.elements_processed = 0
                    config.elements_updated = 0
                    return (0, 0)
                
            except Exception as e:
                logger.error("Error retrieving checklist items: {}".format(str(e)))
                config.elements_processed = 0
                config.elements_updated = 0
                return (0, 0)
            
            # Create value mapping dictionary if enabled
            value_mapping = {}
            if config.mapping_enabled and config.mapping_config:
                try:
                    mapping_data = json.loads(config.mapping_config)
                    for mapping in mapping_data:
                        checklist_value = mapping.get('ChecklistValue')
                        revit_value = mapping.get('RevitValue')
                        if checklist_value and revit_value:
                            value_mapping[checklist_value] = revit_value
                except Exception as e:
                    logger.error("Error parsing mapping config: {}".format(str(e)))
            
            # Build IFC GUID lookup dictionary once - OPTIMIZATION
            ifc_guid_dict = {}
            
            # Get all elements in the document
            all_elements = FilteredElementCollector(revit.doc).WhereElementIsNotElementType().ToElements()
            
            # Track how many elements have IFC GUIDs
            ifc_guid_count = 0
            
            # Build dictionary mapping IFC GUIDs to elements
            for element in all_elements:
                try:
                    # Check for different variations of the parameter name
                    ifc_guid_param = element.LookupParameter("IFCGuid")
                    if not ifc_guid_param:
                        ifc_guid_param = element.LookupParameter("IfcGUID")
                    if not ifc_guid_param:
                        ifc_guid_param = element.LookupParameter("IFC GUID")
                    
                    if ifc_guid_param and ifc_guid_param.HasValue and ifc_guid_param.StorageType == StorageType.String:
                        guid_value = ifc_guid_param.AsString()
                        if guid_value:
                            ifc_guid_dict[guid_value] = element
                            ifc_guid_count += 1
                except:
                    continue
            
            # Start a transaction for this configuration
            t = Transaction(revit.doc, "Batch Import: " + config.DisplayName)
            t.Start()
            
            try:
                # Set an estimated number of elements for the progress tracking
                config.elements_total = len(checklist_items)
                self.update_config_progress(config)
                
                # Process each checklist item directly
                for idx, item in enumerate(checklist_items):
                    # Process UI events periodically
                    if idx % 10 == 0:
                        self.process_ui_events()
                                        
                    processed_count += 1
                    
                    try:
                        # Get the element ID from the checklist item
                        element_id = item.get('object')
                        if not element_id:
                            element_id = item.get('attributes', {}).get('elementId')
                        
                        if not element_id:
                            continue
                        
                        # Find the element by IFC GUID using our dictionary - MUCH FASTER
                        element = ifc_guid_dict.get(element_id)
                        if not element:
                            continue
                        
                        # Get property value
                        checklist_value = self.get_property_value(item, config.streambim_property)
                        
                        if checklist_value is None:
                            continue
                                                
                        # Apply value mapping if enabled
                        if config.mapping_enabled:
                            if checklist_value in value_mapping:
                                original_value = checklist_value
                                checklist_value = value_mapping[checklist_value]
                            else:
                                # Skip if mapping enabled but value not in mapping
                                continue
                        
                        # Get the parameter
                        param = element.LookupParameter(config.revit_parameter)
                        if not param:
                            continue
                        
                        # Skip read-only parameters
                        if param.IsReadOnly:
                            continue
                        
                        # Set parameter value based on type
                        param_storage_type = param.StorageType
                        
                        # Set the parameter value
                        set_value_result = self.set_parameter_value(param, checklist_value, param_storage_type)
                        if set_value_result:
                            updated_count += 1
                            
                    except Exception as e:
                        logger.error("Error processing element: {}".format(str(e)))
                    
                    # Update progress periodically
                    if idx % 5 == 0:
                        config.elements_processed = processed_count
                        config.elements_updated = updated_count
                        self.update_config_progress(config)
                
                # Commit the transaction
                t.Commit()
                
                # Update the final progress
                config.elements_processed = processed_count
                config.elements_updated = updated_count
                self.update_config_progress(config)
                
            except Exception as e:
                # Roll back the transaction if there was an error
                if t.HasStarted():
                    t.RollBack()
                logger.error("Error processing configuration: {}".format(str(e)))
                
        except Exception as e:
            logger.error("Error in process_single_configuration: {}".format(str(e)))
        
        return (processed_count, updated_count)

    
    def config_selection_changed(self, sender, args):
        """Handle configuration selection changed."""
        self.editButton.IsEnabled = self.configsListView.SelectedItem is not None
    
    def edit_button_click(self, sender, args):
        """Handle edit button click."""
        selected_config = self.configsListView.SelectedItem
        
        if not selected_config:
            return
        
        # Store the current configuration
        self.current_config = selected_config
        
        # Update the edit tab fields
        self.checklistNameTextBlock.Text = selected_config.ChecklistName
        self.streambimPropertyTextBlock.Text = selected_config.streambim_property
        self.revitParameterTextBlock.Text = selected_config.revit_parameter
        self.enableMappingCheckBox.IsChecked = selected_config.mapping_enabled
        
        # Load the mapping entries
        self.mappings.Clear()
        
        if selected_config.mapping_config:
            try:
                mapping_data = json.loads(selected_config.mapping_config)
                for mapping in mapping_data:
                    self.mappings.Add(MappingEntry(
                        ChecklistValue=mapping.get('ChecklistValue', ''),
                        RevitValue=mapping.get('RevitValue', '')
                    ))
            except Exception as e:
                logger.error("Error parsing mapping JSON: {}".format(str(e)))
        
        # Update UI
        self.editConfigTab.IsEnabled = True
        self.tabControl.SelectedItem = self.editConfigTab
    
    def save_button_click(self, sender, args):
        """Handle save button click."""
        if not self.current_config:
            return
        
        # Collect mapping data
        mapping_data = []
        for mapping in self.mappings:
            if mapping.ChecklistValue and mapping.RevitValue:
                mapping_data.append({
                    'ChecklistValue': mapping.ChecklistValue,
                    'RevitValue': mapping.RevitValue
                })
        
        mapping_json = json.dumps(mapping_data, ensure_ascii=False)
        
        try:
            # Get all existing configurations
            all_configs = load_configs_with_pickle(revit.doc)
            
            # Create updated config dictionary
            updated_config = {
                'checklist_id': self.current_config.checklist_id,
                'checklist_name': self.current_config.checklist_name,
                'streambim_property': self.current_config.streambim_property,
                'revit_parameter': self.current_config.revit_parameter,
                'mapping_enabled': self.enableMappingCheckBox.IsChecked,
                'mapping_config': mapping_json
            }
            
            # Find and update or add the config
            found = False
            for i, config in enumerate(all_configs):
                if (config.get('checklist_id') == self.current_config.checklist_id and
                    config.get('streambim_property') == self.current_config.streambim_property and
                    config.get('revit_parameter') == self.current_config.revit_parameter):
                    all_configs[i] = updated_config
                    found = True
                    break
                    
            if not found:
                all_configs.append(updated_config)
                
            # Save all configurations
            success = save_configs_with_pickle(revit.doc, all_configs)
            
            if success:
                # Reload configurations to refresh the list
                self.load_configurations()
                
                # Go back to configs tab
                self.back_button_click(None, None)
            
        except Exception as e:
            logger.error("Error saving configuration: {}".format(str(e)))
    
    def refresh_ui(self):
        """Force UI refresh."""
        # First clear and re-set the ItemsSource
        self.configsListView.ItemsSource = None
        self.configsListView.ItemsSource = self.configs
        # Force layout update
        self.configsListView.UpdateLayout()
    
    def delete_button_click(self, sender, args):
        """Handle delete button click."""
        if not self.current_config:
            return
        
        # Confirm deletion
        result = MessageBox.Show(
            "Are you sure you want to delete this configuration?\n\n{}".format(self.current_config.DisplayName),
            "Confirm Deletion",
            MessageBoxButton.YesNo
        )
        
        # Properly check the result - MessageBox.Show returns MessageBoxResult
        if result != MessageBoxResult.Yes:
            return
            
        try:
            # Get all configurations
            all_configs = load_configs_with_pickle(revit.doc)
            
            # Filter out the configuration to delete with exact matching
            updated_configs = []
            found_match = False
            for config in all_configs:
                # Log each config we're checking
                checklist_id_match = config.get('checklist_id') == self.current_config.checklist_id
                property_match = config.get('streambim_property') == self.current_config.streambim_property
                parameter_match = config.get('revit_parameter') == self.current_config.revit_parameter
                
                # Only keep configurations that don't match ALL criteria
                if not (checklist_id_match and property_match and parameter_match):
                    updated_configs.append(config)
                else:
                    found_match = True
            
            # Save the updated configurations
            success = save_configs_with_pickle(revit.doc, updated_configs)
            
            if success:
                # Go back to configs tab first to release the current config
                self.current_config = None
                self.editConfigTab.IsEnabled = False
                self.tabControl.SelectedItem = self.configsTab
                
                # Reload configurations to refresh the list
                self.load_configurations()
                
                # Force UI refresh
                self.refresh_ui()
            
        except Exception as e:
            logger.error("Error deleting configuration: {}".format(str(e)))
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))
    
    def back_button_click(self, sender, args):
        """Handle back button click."""
        # Reset current configuration
        self.current_config = None
        
        # Clear mappings
        self.mappings.Clear()
        
        # Update UI
        self.editConfigTab.IsEnabled = False
        self.tabControl.SelectedItem = self.configsTab
    
    def add_new_row_button_click(self, sender, args):
        """Add a new empty row to the mapping DataGrid."""
        self.mappings.Add(MappingEntry(ChecklistValue="", RevitValue=""))
    
    def remove_row_button_click(self, sender, args):
        """Remove the selected row from the mapping DataGrid."""
        if self.mappingDataGrid.SelectedItem:
            self.mappings.Remove(self.mappingDataGrid.SelectedItem)
    
    def enable_mapping_checked(self, sender, args):
        """Handle enabling the mapping feature."""
        self.mappingGrid.Visibility = Visibility.Visible
        
    def enable_mapping_unchecked(self, sender, args):
        """Handle disabling the mapping feature."""
        self.mappingGrid.Visibility = Visibility.Collapsed
    
    def update_status(self, message):
        """Update status text."""
        # Status bar removed from UI - method kept for compatibility but does nothing
        pass

def find_elements_with_param(param_name, doc):
    """Find all elements with a specific parameter."""
    all_elements = []
    
    try:
        # Use FilteredElementCollector to get all elements
        collector = FilteredElementCollector(doc)\
            .WhereElementIsNotElementType()\
            .WhereElementIsViewIndependent()
            
        # Filter out elements that don't have the parameter
        for element in collector:
            try:
                # Skip elements that are not modifiable
                if element.IsValidObject and not element.IsReadOnly:
                    # Check for the parameter
                    param = element.LookupParameter(param_name)
                    if param and param.HasValue:
                        all_elements.append(element)
            except:
                continue
                
        return all_elements
        
    except Exception as e:
        logger.error("Error finding elements with parameter '{}': {}".format(param_name, str(e)))
        return []
    
# Main execution
if __name__ == '__main__':
    # Check for saved login BEFORE creating the window
    temp_client = streambim_api.StreamBIMClient()
    temp_client.load_tokens()
    
    if not temp_client.idToken:
        # No saved login found - show error and exit without opening window
        MessageBox.Show(
            "No saved StreamBIM login found. Please log in using the ChecklistImporter tool first.",
            "StreamBIM Login Required",
            MessageBoxButton.OK
        )
    else:
        # Show the Configuration Editor UI
        ConfigEditorUI().ShowDialog()