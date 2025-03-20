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
    sys.path.append(lib_path)

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
from Autodesk.Revit.DB import *

from System import EventHandler
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import MessageBox, MessageBoxButton, MessageBoxResult, Visibility
from System.Windows.Controls import SelectionMode

from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Import extensible storage
from extensible_storage import BaseSchema, simple_field

# Initialize logger
logger = script.get_logger()

def get_all_mapping_storages(doc):
    """Get all data storage elements with our mapping schema."""
    logger.debug("This function is deprecated and will be removed in future versions.")
    return []

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

class ConfigEditorUI(forms.WPFWindow):
    """Configuration Editor UI implementation."""
    
    def __init__(self):
        """Initialize the Configuration Editor UI."""
        # Initialize WPF window
        forms.WPFWindow.__init__(self, 'ConfigEditor.xaml')
        
        # Initialize data collections
        self.configs = ObservableCollection[object]()
        self.mappings = ObservableCollection[object]()
        self.current_config = None
        
        # Set up event handlers
        self.configsListView.SelectionChanged += self.config_selection_changed
        self.configsListView.SelectionMode = SelectionMode.Single
        self.editButton.Click += self.edit_button_click
        self.saveButton.Click += self.save_button_click
        self.deleteButton.Click += self.delete_button_click
        self.backButton.Click += self.back_button_click
        self.addNewRowButton.Click += self.add_new_row_button_click
        self.removeRowButton.Click += self.remove_row_button_click
        self.enableMappingCheckBox.Checked += self.enable_mapping_checked
        self.enableMappingCheckBox.Unchecked += self.enable_mapping_unchecked
        
        # Initialize UI
        self.configsListView.ItemsSource = self.configs
        self.mappingDataGrid.ItemsSource = self.mappings
        
        # Initial UI state
        self.tabControl.SelectedItem = self.configsTab
        self.editConfigTab.IsEnabled = False
        
        # Load configurations
        self.load_configurations()
    
    def load_configurations(self):
        """Load all mapping configurations from storage."""
        try:
            # Clear existing configurations
            self.configs.Clear()
            logger.debug("Cleared existing configurations")
            
            # Load configurations from consolidated storage
            loaded_configs = load_configs_with_pickle(revit.doc)
            logger.debug("Loaded {} configurations from storage".format(len(loaded_configs)))
            
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
                    logger.debug("Added config: {} -> {}".format(
                        config_item.streambim_property, 
                        config_item.revit_parameter
                    ))
                
                self.update_status("Loaded {} configurations".format(len(self.configs)))
            else:
                self.update_status("No configurations found")
                
            # Force the UI to refresh the configurations list
            self.refresh_ui()
            
        except Exception as e:
            logger.error("Error loading configurations: {}".format(str(e)))
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))
            self.update_status("Error loading configurations: {}".format(str(e)))
    
    def config_selection_changed(self, sender, args):
        """Handle configuration selection changed."""
        self.editButton.IsEnabled = self.configsListView.SelectedItem is not None
    
    def edit_button_click(self, sender, args):
        """Handle edit button click."""
        selected_config = self.configsListView.SelectedItem
        
        if not selected_config:
            self.update_status("No configuration selected")
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
        self.update_status("Editing configuration: {}".format(selected_config.DisplayName))
    
    def save_button_click(self, sender, args):
        """Handle save button click."""
        if not self.current_config:
            self.update_status("No configuration to save")
            return
        
        # Collect mapping data
        mapping_data = []
        for mapping in self.mappings:
            if mapping.ChecklistValue and mapping.RevitValue:
                mapping_data.append({
                    'ChecklistValue': mapping.ChecklistValue,
                    'RevitValue': mapping.RevitValue
                })
        
        mapping_json = json.dumps(mapping_data)
        
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
                
                # Update UI
                self.update_status("Configuration saved successfully")
                
                # Go back to configs tab
                self.back_button_click(None, None)
            else:
                self.update_status("Error saving configuration")
            
        except Exception as e:
            logger.error("Error saving configuration: {}".format(str(e)))
            self.update_status("Error saving configuration: {}".format(str(e)))
    
    def refresh_ui(self):
        """Force UI refresh."""
        logger.debug("Forcing UI refresh")
        # First clear and re-set the ItemsSource
        self.configsListView.ItemsSource = None
        self.configsListView.ItemsSource = self.configs
        # Force layout update
        self.configsListView.UpdateLayout()
    
    def delete_button_click(self, sender, args):
        """Handle delete button click."""
        if not self.current_config:
            self.update_status("No configuration to delete")
            return
        
        # Log details about the current config
        logger.debug("Current config for deletion:")
        logger.debug("- Type: {}".format(type(self.current_config)))
        logger.debug("- DisplayName: {}".format(self.current_config.DisplayName))
        
        # Confirm deletion
        result = MessageBox.Show(
            "Are you sure you want to delete this configuration?\n\n{}".format(self.current_config.DisplayName),
            "Confirm Deletion",
            MessageBoxButton.YesNo
        )
        
        # Properly check the result - MessageBox.Show returns MessageBoxResult
        logger.debug("MessageBox result: {}".format(result))
        if result != MessageBoxResult.Yes:
            logger.debug("Deletion cancelled by user")
            return
            
        logger.debug("User confirmed deletion - proceeding to delete")
        
        try:
            # Get all configurations
            all_configs = load_configs_with_pickle(revit.doc)
            
            # Log configuration being deleted
            logger.debug("Deleting configuration:")
            logger.debug("- Checklist ID: {}".format(self.current_config.checklist_id))
            logger.debug("- Checklist name: {}".format(self.current_config.checklist_name))
            logger.debug("- StreamBIM property: {}".format(self.current_config.streambim_property))
            logger.debug("- Revit parameter: {}".format(self.current_config.revit_parameter))
            
            # Count configurations before filtering
            logger.debug("Total configurations before delete: {}".format(len(all_configs)))
            
            # Filter out the configuration to delete with exact matching
            updated_configs = []
            found_match = False
            for config in all_configs:
                # Log each config we're checking
                checklist_id_match = config.get('checklist_id') == self.current_config.checklist_id
                property_match = config.get('streambim_property') == self.current_config.streambim_property
                parameter_match = config.get('revit_parameter') == self.current_config.revit_parameter
                
                logger.debug("Checking config: {}/{}/{} - Match: {}/{}/{}".format(
                    config.get('checklist_id'), 
                    config.get('streambim_property'),
                    config.get('revit_parameter'),
                    checklist_id_match,
                    property_match,
                    parameter_match
                ))
                
                # Only keep configurations that don't match ALL criteria
                if not (checklist_id_match and property_match and parameter_match):
                    updated_configs.append(config)
                else:
                    found_match = True
                    logger.debug("Found matching config to delete")
            
            if not found_match:
                logger.debug("WARNING: No matching configuration found to delete!")
            
            # Count configurations after filtering
            logger.debug("Total configurations after delete: {}".format(len(updated_configs)))
            
            # Save the updated configurations
            logger.debug("Saving updated configurations...")
            success = save_configs_with_pickle(revit.doc, updated_configs)
            
            if success:
                logger.debug("Configurations saved successfully, updating UI...")
                # Go back to configs tab first to release the current config
                self.current_config = None
                self.editConfigTab.IsEnabled = False
                self.tabControl.SelectedItem = self.configsTab
                
                # Reload configurations to refresh the list
                self.load_configurations()
                
                # Force UI refresh
                self.refresh_ui()
                
                # Update UI
                self.update_status("Configuration deleted successfully")
            else:
                self.update_status("Error deleting configuration")
            
        except Exception as e:
            logger.error("Error deleting configuration: {}".format(str(e)))
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))
            self.update_status("Error deleting configuration: {}".format(str(e)))
    
    def back_button_click(self, sender, args):
        """Handle back button click."""
        # Reset current configuration
        self.current_config = None
        
        # Clear mappings
        self.mappings.Clear()
        
        # Update UI
        self.editConfigTab.IsEnabled = False
        self.tabControl.SelectedItem = self.configsTab
        self.update_status("Returned to configurations list")
    
    def add_new_row_button_click(self, sender, args):
        """Add a new empty row to the mapping DataGrid."""
        self.mappings.Add(MappingEntry(ChecklistValue="", RevitValue=""))
        self.update_status("Added new mapping row")
    
    def remove_row_button_click(self, sender, args):
        """Remove the selected row from the mapping DataGrid."""
        if self.mappingDataGrid.SelectedItem:
            self.mappings.Remove(self.mappingDataGrid.SelectedItem)
            self.update_status("Removed selected mapping row")
    
    def enable_mapping_checked(self, sender, args):
        """Handle enabling the mapping feature."""
        self.mappingGrid.Visibility = Visibility.Visible
        
    def enable_mapping_unchecked(self, sender, args):
        """Handle disabling the mapping feature."""
        self.mappingGrid.Visibility = Visibility.Collapsed
    
    def update_status(self, message):
        """Update status text."""
        self.statusTextBlock.Text = message
        logger.debug(message)

# Main execution
if __name__ == '__main__':
    # Log information about the streambim_api module
    logger.debug("StreamBIM API module location: {}".format(streambim_api.__file__))
    logger.debug("save_configs_with_pickle function: {}".format(save_configs_with_pickle))
    
    # Show the Configuration Editor UI
    ConfigEditorUI().ShowDialog()