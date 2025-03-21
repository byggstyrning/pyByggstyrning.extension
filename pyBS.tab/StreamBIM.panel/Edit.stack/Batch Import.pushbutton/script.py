# -*- coding: utf-8 -*-
__title__ = "Batch\nImport"
__author__ = "Byggstyrning AB"
__doc__ = """Run all saved StreamBIM checklist configurations at once.

This tool applies all saved mapping configurations to all elements
in the model that have an IfcGUID parameter."""

import os
import sys
import clr
import json
import pickle
import base64
from collections import namedtuple

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
    sys.path.append(lib_path)

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

# Import StreamBIM API
from streambim import streambim_api

# Import revit_utils functions
from revit.revit_utils import get_element_by_ifc_guid

# Import StreamBIMSettingsSchema and related functions directly from the module
from streambim.streambim_api import StreamBIMSettingsSchema
from streambim.streambim_api import get_or_create_settings_storage
from streambim.streambim_api import load_configs_with_pickle
from streambim.streambim_api import save_configs_with_pickle
from streambim.streambim_api import get_saved_project_id

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

class BatchImportUI(forms.WPFWindow):
    """Batch Import UI implementation."""
    
    def __init__(self):
        """Initialize the Batch Import UI."""
        # Initialize WPF window
        forms.WPFWindow.__init__(self, 'BatchImporter.xaml')
        
        # Initialize StreamBIM API client
        self.streambim_client = streambim_api.StreamBIMClient()
        
        # Initialize data collections
        self.configs = ObservableCollection[object]()
        
        # Set up event handlers
        self.runAllButton.Click += self.run_all_button_click
        self.runSelectedButton.Click += self.run_selected_button_click
        
        # Initialize UI
        self.configsListView.ItemsSource = self.configs
        self.configsListView.SelectionMode = SelectionMode.Extended  # Allow multiple selection
        
        # Load configurations
        self.load_configurations()
        
        # Try automatic login
        self.try_automatic_login()
        
        # Update summary text
        self.summaryTextBlock.Text = "Found {} configurations".format(len(self.configs))
    
    def try_automatic_login(self):
        """Attempt to automatically log in using saved tokens."""
        # Load tokens from file first
        self.streambim_client.load_tokens()
        
        # Check if token exists
        if self.streambim_client.idToken:
            self.update_status("Found saved StreamBIM login...")
            
            # Try to load saved project ID
            saved_project_id = get_saved_project_id(revit.doc)
            if saved_project_id:
                self.streambim_client.set_current_project(saved_project_id)
                self.update_status("Using saved project ID: {}".format(saved_project_id))
            
            return True
        else:
            self.update_status("No saved StreamBIM login found. Please log in using the ChecklistImporter first.")
            # Show a message to the user
            MessageBox.Show(
                "No saved StreamBIM login found. Please log in using the ChecklistImporter tool first.",
                "StreamBIM Login Required",
                MessageBoxButton.OK
            )
            return False
    
    def load_configurations(self):
        """Load all mapping configurations from storage."""
        self.configs.Clear()
        
        # Load configurations from consolidated storage
        logger.debug("Loading configurations from storage...")
        loaded_configs = load_configs_with_pickle(revit.doc)
        
        if loaded_configs:
            logger.debug("Found {} configurations in storage".format(len(loaded_configs)))
            # Add them to the observable collection
            for config_dict in loaded_configs:
                logger.debug("Processing config: checklist_id={}, property={}, parameter={}".format(
                    config_dict.get('checklist_id'),
                    config_dict.get('streambim_property'),
                    config_dict.get('revit_parameter')
                ))
                
                config = ConfigItem(
                    id=None,  # We don't use element IDs anymore
                    checklist_id=config_dict.get('checklist_id'),
                    checklist_name=config_dict.get('checklist_name', 'Unknown Checklist'),
                    streambim_property=config_dict.get('streambim_property'),
                    revit_parameter=config_dict.get('revit_parameter'),
                    mapping_enabled=config_dict.get('mapping_enabled'),
                    mapping_config=config_dict.get('mapping_config')
                )
                self.configs.Add(config)
            
            self.update_status("Loaded {} configurations".format(len(self.configs)))
        else:
            logger.debug("No configurations found in storage")
            self.update_status("No configurations found")
    
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
                self.update_status("Login cancelled or failed")
                return
                
        # Set the project ID in the API client
        if project_id:
            self.api_client.set_current_project(project_id)
        
        # Process all configurations
        self.run_import_configurations(self.configs)
    
    def run_selected_button_click(self, sender, args):
        """Process selected configs when the Run Selected button is clicked."""
        if not self.configsListView.SelectedItems or self.configsListView.SelectedItems.Count == 0:
            self.update_status("No configurations selected")
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
                self.update_status("Login cancelled or failed")
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
                self.update_status("Username and password are required")
                return False
                
            self.update_status("Logging in to StreamBIM...")
            
            # Login to StreamBIM
            login_success = self.api_client.login(username, password)
            
            if login_success:
                self.update_status("Logged in successfully")
                
                # Get projects
                projects = self.api_client.get_projects()
                
                if not projects or len(projects) == 0:
                    self.update_status("No projects found")
                    return False
                
                # Use the first project if none is specified
                if not self.api_client.current_project:
                    first_project_id = projects[0].get('id')
                    self.api_client.set_current_project(first_project_id)
                    
                    # Save the project ID
                    self.save_project_id(first_project_id)
                    
                    self.update_status("Using project: {}".format(
                        projects[0].get('attributes', {}).get('name', 'Unknown')
                    ))
                
                return True
            else:
                self.update_status("Login failed: {}".format(self.api_client.last_error))
                return False
        else:
            self.update_status("Login cancelled")
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
                self.update_status("Saved project ID: {}".format(project_id))
        except Exception as e:
            logger.error("Error saving project ID: {}".format(str(e)))
    
    def run_import_configurations(self, configs):
        """Run import for a list of configurations."""
        # Disable buttons during processing
        self.runAllButton.IsEnabled = False
        self.runSelectedButton.IsEnabled = False
        
        logger.debug("Starting batch import process for {} configurations".format(len(configs)))
        self.update_status("Starting batch import process...")
        
        try:
            # Update main progress bar max
            self.mainProgressBar.Maximum = len(configs)
            self.mainProgressBar.Value = 0
            
            # Track total elements processed and updated
            total_processed = 0
            total_updated = 0
            
            # Process each configuration separately
            for i, config in enumerate(configs):
                logger.debug("==== Processing configuration {}/{}: {} ====".format(
                    i + 1, len(configs), config.DisplayName
                ))
                logger.debug("Checklist: {} (ID: {})".format(config.ChecklistName, config.checklist_id))
                logger.debug("Property: {} -> Parameter: {}".format(config.streambim_property, config.revit_parameter))
                logger.debug("Mapping enabled: {}".format(config.mapping_enabled))
                
                # Update status
                self.update_status("Processing configuration {}/{}: {}".format(
                    i + 1, len(configs), config.DisplayName
                ))
                
                # Update main progress bar
                self.mainProgressBar.Value = i
                
                # Process UI events
                self.process_ui_events()
                
                # Skip configurations without checklist ID
                if not config.checklist_id:
                    logger.debug("Skipping configuration - no checklist ID")
                    self.update_status("Skipping configuration with no checklist ID: {}".format(config.DisplayName))
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
            
            logger.debug("Batch import completed. Processed {} configurations. Updated {}/{} elements.".format(
                len(configs), total_updated, total_processed))
            self.update_status("Batch import completed. Updated {}/{} elements.".format(
                total_updated, total_processed))
            
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
            logger.debug("Batch import failed with exception")
            self.update_status("Error: {}".format(str(e)))
            
            # Show error message box
            MessageBox.Show(
                "An error occurred during batch import:\n\n{}".format(str(e)),
                "Batch Import Error",
                MessageBoxButton.OK
            )
            
        finally:
            logger.debug("Batch import process completed, resetting UI state")
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
            
            # Use a simplified approach for the Dispatcher
            current_dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher
            current_dispatcher.Invoke(Action(lambda: None))
        except Exception as e:
            # Log any errors but continue execution
            logger.debug("Error processing UI events: {}".format(str(e)))
    
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
        self.statusTextBlock.Text = message
        self.process_ui_events()

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
                    logger.debug("Could not convert '{}' to integer".format(value))
                    return False
                    
            elif storage_type == StorageType.Double:
                # Try to convert to double
                try:
                    double_value = float(value)
                    if param.AsDouble() != double_value:
                        param.Set(double_value)
                        return True
                except (ValueError, TypeError):
                    logger.debug("Could not convert '{}' to double".format(value))
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
                    logger.debug("Could not convert '{}' to ElementId".format(value))
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
                logger.debug("Retrieving checklist items for checklist ID: {}".format(config.checklist_id))
                checklist_items = self.api_client.get_checklist_items(config.checklist_id, limit=0)
                if not checklist_items:
                    logger.debug("No checklist items found for checklist ID: {}".format(config.checklist_id))
                    self.update_status("No checklist items found for checklist ID: {}".format(config.checklist_id))
                    config.elements_processed = 0
                    config.elements_updated = 0
                    return (0, 0)
                    
                logger.debug("Retrieved {} checklist items for {}".format(len(checklist_items), config.ChecklistName))
                self.update_status("Retrieved {} checklist items for {}".format(
                    len(checklist_items), config.ChecklistName))
                
            except Exception as e:
                logger.error("Error retrieving checklist items: {}".format(str(e)))
                config.elements_processed = 0
                config.elements_updated = 0
                return (0, 0)
            
            # Create value mapping dictionary if enabled
            value_mapping = {}
            if config.mapping_enabled and config.mapping_config:
                try:
                    logger.debug("Loading value mappings...")
                    mapping_data = json.loads(config.mapping_config)
                    for mapping in mapping_data:
                        checklist_value = mapping.get('ChecklistValue')
                        revit_value = mapping.get('RevitValue')
                        if checklist_value and revit_value:
                            value_mapping[checklist_value] = revit_value
                    logger.debug("Loaded {} value mappings".format(len(value_mapping)))
                except Exception as e:
                    logger.error("Error parsing mapping config: {}".format(str(e)))
            
            # Start a transaction for this configuration
            t = Transaction(revit.doc, "Batch Import: " + config.DisplayName)
            t.Start()
            
            try:
                logger.debug("Starting element processing")
                
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
                        
                        # Find the element by IFC GUID
                        element = get_element_by_ifc_guid(element_id)
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
                logger.debug("Completed configuration {}/{}: {} - Processed: {}, Updated: {}".format(
                    config_index + 1, total_configs, config.DisplayName, processed_count, updated_count))
                
            except Exception as e:
                # Roll back the transaction if there was an error
                if t.HasStarted():
                    t.RollBack()
                logger.error("Error processing configuration: {}".format(str(e)))
                self.update_status("Error processing configuration: {}".format(str(e)))
                
        except Exception as e:
            logger.error("Error in process_single_configuration: {}".format(str(e)))
        
        return (processed_count, updated_count)

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
                
        logger.debug("Found {} elements with parameter '{}'".format(len(all_elements), param_name))
        return all_elements
        
    except Exception as e:
        logger.error("Error finding elements with parameter '{}': {}".format(param_name, str(e)))
        return []

# Main execution
if __name__ == '__main__':
    # Show the Batch Import UI
    BatchImportUI().ShowDialog()

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.