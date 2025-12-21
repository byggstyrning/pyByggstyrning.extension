# -*- coding: utf-8 -*-
"""Import StreamBIM data into Revit parameters.

This tool connects to StreamBIM to import parameter values from StreamBIM checklists
into Revit parameters for elements in the current view.
"""
__title__ = "Checklist\nImporter"
__author__ = "Byggstyrning AB"
__doc__ = "Import StreamBIM checklist items data into Revit instance parameters"

## todo:
# - kunna edita konfigen
# - kör-knapp för att köra alla importer
# - köra suync på events


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
panel_dir = op.dirname(script_path)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(op.dirname(tab_dir))
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
clr.AddReference('RevitAPIUI')

# Try to add Xceed WPF Toolkit reference
# The DLL should be placed in the lib folder or system path
try:
    # Try to load from lib folder first
    xceed_dll_path = op.join(lib_path, 'Xceed.Wpf.Toolkit.dll')
    if op.exists(xceed_dll_path):
        clr.AddReferenceToFileAndPath(xceed_dll_path)
        logger.debug("Loaded Xceed Toolkit from lib folder")
    else:
        # Try to load from system (if installed globally)
        clr.AddReference("Xceed.Wpf.Toolkit")
        logger.debug("Loaded Xceed Toolkit from system")
except Exception as e:
    logger.warning("Could not load Xceed WPF Toolkit: {}. Some features may not be available.".format(str(e)))

from Autodesk.Revit.DB import *

from System import EventHandler
from System.Collections.ObjectModel import ObservableCollection
from System.Dynamic import ExpandoObject
from System.Windows import MessageBox, MessageBoxButton, Visibility
from System.Windows.Media import SolidColorBrush
from System.Windows.Media.Animation import DoubleAnimation
from System.Windows.Media import Colors

from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Import custom modules from the extension lib
from streambim import streambim_api
from revit import revit_utils

# Import extensible storage
from extensible_storage import BaseSchema, simple_field

# Import StreamBIMSettingsSchema and related functions directly from the module
from streambim.streambim_api import StreamBIMSettingsSchema
from streambim.streambim_api import get_or_create_settings_storage
from streambim.streambim_api import load_configs_with_pickle
from streambim.streambim_api import save_configs_with_pickle
from streambim.streambim_api import get_saved_project_id

# Initialize logger
logger = script.get_logger()

# Define data classes using namedtuple for IronPython
Project = namedtuple('Project', ['Id', 'Name', 'Description'])
Checklist = namedtuple('Checklist', ['Id', 'Name'])
PropertyValue = namedtuple('PropertyValue', ['Name', 'Sample'])

# Replace the MappingEntry namedtuple with a proper class
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

    def __str__(self):
        return "MappingEntry(ChecklistValue='{}', RevitValue='{}')".format(
            self._checklist_value, self._revit_value
        )

class MappingSchema(BaseSchema):
    """Schema for storing parameter mapping configurations"""
    
    guid = "f810d3a5-abe3-46e5-989b-b4da078b356b"
    
    @simple_field(value_type="string")
    def checklist_id():
        """The selected checklist ID"""

    @simple_field(value_type="string")
    def checklist_name():
        """The selected checklist name"""
        
    @simple_field(value_type="string")
    def streambim_property():
        """The selected StreamBIM property"""
        
    @simple_field(value_type="string")
    def revit_parameter():
        """The selected Revit parameter"""

    @simple_field(value_type="string")
    def enableMapping():
        """Whether to enable parameter mapping (stored as 'True' or 'False' string)"""

    @simple_field(value_type="string")
    def mapping_config():
        """The mapping configuration as JSON string"""

def get_or_create_data_storage(doc):
    """Get existing or create new data storage element."""
    data_storage = FilteredElementCollector(doc)\
        .OfClass(ExtensibleStorage.DataStorage)\
        .ToElements()
    
    # Look for our storage with the schema
    for ds in data_storage:
        # Check if this storage has our schema
        entity = ds.GetEntity(StreamBIMSettingsSchema.schema)
        if entity.IsValid():
            return ds
    
    # If not found, create a new one
    with revit.Transaction("Create StreamBIM Storage", doc):
        return ExtensibleStorage.DataStorage.Create(doc)

def get_or_create_mapping_storage(doc):
    """Get existing or create new data storage element."""
    if not doc:
        logger.error("No active document available")
        return None
        
    try:
        logger.debug("Searching for existing mapping storage...")
        data_storage = FilteredElementCollector(doc)\
            .OfClass(ExtensibleStorage.DataStorage)\
            .ToElements()
        
        # Look for our storage with the schema
        for ds in data_storage:
            try:
                # Check if this storage has our schema
                entity = ds.GetEntity(MappingSchema.schema)
                if entity.IsValid():
                    logger.debug("Found existing mapping storage")
                    return ds
            except Exception as e:
                logger.debug("Error checking storage entity: {}".format(str(e)))
                continue
        
        logger.debug("No existing mapping storage found, creating new one...")
        # If not found, create a new one
        with revit.Transaction("Create StreamBIM Mapping Storage", doc):
            new_storage = ExtensibleStorage.DataStorage.Create(doc)
            logger.debug("Created new mapping storage")
            return new_storage
            
    except Exception as e:
        logger.error("Error in get_or_create_mapping_storage: {}".format(str(e)))
        return None

class StreamBIMImporterUI(forms.WPFWindow):
    """StreamBIM Importer UI implementation."""
    
    def __init__(self):
        """Initialize the StreamBIM Importer UI."""
        # Initialize WPF window
        forms.WPFWindow.__init__(self, 'ChecklistImporter.xaml')
        
        # Initialize StreamBIM API client
        self.streambim_client = streambim_api.StreamBIMClient()
                
        # Initialize data collections
        self.projects = ObservableCollection[object]()
        self.checklists = ObservableCollection[object]()
        self.all_checklists = []  # Store all checklists for filtering
        self.checklist_items = []
        self.streambim_properties = []
        self.updated_elements = []  # Store updated elements for isolation
        self.saved_project_id = None
        
        # Initialize mapping data
        self.mappings = ObservableCollection[object]()
        self.mapping_storage = None
        self.is_loading_configuration = False  # Flag to prevent recursive loading
        
        # Initialize grouped checklist metadata (runtime only, not persisted)
        self.selected_checklist_group_by = None
        self.selected_checklist_building_id = None
        self.all_checklist_records = {}  # Cache checklist records by ID for metadata lookup

        # Set up event handlers
        self.loginButton.Click += self.login_button_click
        self.logoutButton.Click += self.logout_button_click
        self.selectProjectButton.Click += self.select_project_button_click
        self.selectChecklistButton.Click += self.select_checklist_button_click
        self.streamBIMPropertiesComboBox.SelectionChanged += self.streambim_property_selected
        self.revitParametersComboBox.SelectionChanged += self.revit_parameter_selected
        self.importButton.Click += self.import_button_click
        self.isolateButton.Click += self.isolate_button_click
        
        # Add new event handlers for mapping
        self.enableMappingCheckBox.Checked += self.enable_mapping_checked
        self.enableMappingCheckBox.Unchecked += self.enable_mapping_unchecked
        self.saveMappingButton.Click += self.save_mapping_button_click
        self.addNewRowButton.Click += self.add_new_row_button_click
        
        # Add search text changed handler
        self.checklistSearchBox.TextChanged += self.checklist_search_changed
        
        # Initialize UI
        self.projectsListView.ItemsSource = self.projects
        self.checklistsListView.ItemsSource = self.checklists
        self.mappingDataGrid.ItemsSource = self.mappings

        # Load saved project ID
        self.saved_project_id = self.load_saved_project_id()

        # Try automatic login
        
        self.try_automatic_login()

    def try_automatic_login(self):
        """Attempt to automatically log in using saved tokens."""
        if self.streambim_client.idToken:
            self.update_status("Found saved login...")
            # If we have a saved username, display it in the username field
            try:
                if hasattr(self.streambim_client, 'username') and self.streambim_client.username:
                    self.usernameTextBox.Text = self.streambim_client.username
            except:
                # If username attribute doesn't exist, just continue
                pass
            self.on_login_success()
            return True
        return False
    
    def on_login_success(self):
        """Handle successful login."""
        # Update login status
        try:
            if hasattr(self.streambim_client, 'username') and self.streambim_client.username:
                self.loginStatusTextBlock.Text = "Logged in as: " + self.streambim_client.username
            else:
                self.loginStatusTextBlock.Text = "Logged in"
        except:
            self.loginStatusTextBlock.Text = "Logged in"
            
        # Disable login fields
        self.usernameTextBox.IsEnabled = False
        self.passwordBox.IsEnabled = False
        self.serverUrlTextBox.IsEnabled = False
        
        # Update buttons
        self.loginButton.IsEnabled = False
        self.logoutButton.IsEnabled = True
        self.passwordBox.IsEnabled = False
        
        # Enable project tab and switch to it
        self.projectTab.IsEnabled = True
        self.tabControl.SelectedItem = self.projectTab
        
        # Get projects
        self.update_status("Retrieving projects...")
        # Show busy indicator during project retrieval
        try:
            self.busyIndicator.IsBusy = True
            self.busyIndicator.BusyContent = "Retrieving projects..."
        except:
            pass
        
        projects = self.streambim_client.get_projects()
        
        # Hide busy indicator
        try:
            self.busyIndicator.IsBusy = False
        except:
            pass
        if projects:
            self.projects.Clear()
            saved_project = None

            for project in projects:
                # Handle new project-links data structure
                attrs = project.get('attributes', {})
                project_obj = Project(
                    Id=str(project.get('id')),
                    Name=attrs.get('name', 'Unknown'),
                    Description=attrs.get('description', '')
                )
                self.projects.Add(project_obj)
                
                # Check if this is our saved project
                if self.saved_project_id and str(project.get('id')) == self.saved_project_id:
                    saved_project = project_obj
            
            self.update_status("Retrieved {} projects".format(len(projects)))
            
            # If we have a saved project, automatically select it
            if saved_project:
                self.projectsListView.SelectedItem = saved_project
                self.select_project_button_click(None, None)
        else:
            error_msg = self.streambim_client.last_error or "No projects found"
            self.update_status(error_msg)
    
    def login_button_click(self, sender, args):
        """Handle login button click event."""
        username = self.usernameTextBox.Text
        password = self.passwordBox.Password
        server_url = self.serverUrlTextBox.Text
        
        if not username or not password:
            self.update_status("Please enter username and password")
            return
        
        # Update UI
        self.loginButton.IsEnabled = False
        self.update_status("Logging in to StreamBIM...")
        
        # Show busy indicator during login
        try:
            self.busyIndicator.IsBusy = True
            self.busyIndicator.BusyContent = "Logging in..."
        except:
            pass  # BusyIndicator might not be available if Xceed DLL not loaded
        
        # Set server URL and login
        self.streambim_client.base_url = server_url
        success = self.streambim_client.login(username, password)
        
        # Hide busy indicator
        try:
            self.busyIndicator.IsBusy = False
        except:
            pass
        
        if success:
            self.on_login_success()
        else:
            error_msg = self.streambim_client.last_error or "Login failed. Please check your credentials."
            self.update_status(error_msg)
            self.loginStatusTextBlock.Text = "Login failed: " + error_msg
            self.loginButton.IsEnabled = True
    
    def logout_button_click(self, sender, args):
        """Handle logout button click event."""
        # Clear tokens
        try:
            # Use the clear_tokens method if it exists
            self.streambim_client.clear_tokens()
        except:
            # Fallback to manually clearing tokens
            self.streambim_client.idToken = None
            self.streambim_client.accessToken = None
            if hasattr(self.streambim_client, 'username'):
                self.streambim_client.username = None
        
        # Reset UI
        self.loginButton.IsEnabled = True
        self.logoutButton.IsEnabled = False
        self.projectTab.IsEnabled = False
        self.checklistTab.IsEnabled = False
        self.parameterTab.IsEnabled = False
        self.tabControl.SelectedItem = self.loginTab
        
        # Enable login fields
        self.usernameTextBox.IsEnabled = True
        self.passwordBox.IsEnabled = True
        self.serverUrlTextBox.IsEnabled = True
        
        # Clear data
        self.projects.Clear()
        self.checklists.Clear()
        
        # Update status
        self.loginStatusTextBlock.Text = "Logged out"
        self.update_status("Logged out successfully")
    
    def select_project_button_click(self, sender, args):
        """Handle project selection button click."""
        selected_project = self.projectsListView.SelectedItem
        
        if not selected_project:
            self.update_status("Please select a project")
            return
        
        # Update UI
        self.selectProjectButton.IsEnabled = False
        self.update_status("Selecting project: " + selected_project.Name)
        
        # Show busy indicator during checklist retrieval
        try:
            self.busyIndicator.IsBusy = True
            self.busyIndicator.BusyContent = "Retrieving checklists..."
        except:
            pass
        
        # Save project ID
        if self.save_project_id(selected_project.Id):
            self.saved_project_id = selected_project.Id
        
        # Set current project
        self.streambim_client.set_current_project(selected_project.Id)
        
        # Get checklists
        checklists = self.streambim_client.get_checklists()
        
        # Hide busy indicator
        try:
            self.busyIndicator.IsBusy = False
        except:
            pass
        if checklists:
            self.checklists.Clear()
            self.all_checklists = []  # Clear all checklists list
            self.all_checklist_records = {}  # Cache checklist records for metadata lookup
            
            for checklist in checklists:
                checklist_id = checklist.get('id')
                checklist_obj = Checklist(
                    Id=checklist_id,
                    Name=checklist.get('attributes', {}).get('name', 'Unknown')
                )
                self.checklists.Add(checklist_obj)
                self.all_checklists.append(checklist_obj)  # Store in all checklists list for filtering
                self.all_checklist_records[checklist_id] = checklist  # Cache full record
            
            # Enable checklist tab
            self.checklistTab.IsEnabled = True
            self.tabControl.SelectedItem = self.checklistTab
            self.update_status("Retrieved {} checklists".format(len(checklists)))
        else:
            error_msg = self.streambim_client.last_error or "No checklists found for this project"
            self.update_status(error_msg)
        
        self.selectProjectButton.IsEnabled = True
    
    def select_checklist_button_click(self, sender, args):
        """Handle checklist selection button click."""
        selected_checklist = self.checklistsListView.SelectedItem
        
        if not selected_checklist:
            self.update_status("Please select a checklist")
            return
        
        # Update UI
        self.selectChecklistButton.IsEnabled = False
        self.update_status("Loading checklist preview...")
        
        # Show busy indicator during checklist preview loading
        try:
            self.busyIndicator.IsBusy = True
            self.busyIndicator.BusyContent = "Loading checklist preview..."
        except:
            pass
        
        # Update selected checklist text
        self.selectedChecklistTextBlock.Text = selected_checklist.Name
        
        # Get checklist items (limited for preview)
        checklist_items = self.streambim_client.get_checklist_items(selected_checklist.Id, limit=5)
        
        # Hide busy indicator
        try:
            self.busyIndicator.IsBusy = False
        except:
            pass
        if checklist_items:
            self.checklist_items = checklist_items
            self.selected_checklist_id = selected_checklist.Id  # Store for later use
            self.selected_checklist_name = selected_checklist.Name
            
            # Store grouped checklist metadata (runtime only, not persisted)
            checklist_record = self.all_checklist_records.get(selected_checklist.Id)
            if checklist_record:
                attrs = checklist_record.get('attributes', {})
                self.selected_checklist_group_by = attrs.get('group-by', '') or ''
                
                # Get building ID from relationships
                relationships = checklist_record.get('relationships', {})
                buildings = relationships.get('buildings', {}).get('data', [])
                if buildings and len(buildings) > 0:
                    self.selected_checklist_building_id = buildings[0].get('id')
                else:
                    self.selected_checklist_building_id = None
                    logger.warning("No building ID found for checklist {}".format(selected_checklist.Id))
            else:
                self.selected_checklist_group_by = ''
                self.selected_checklist_building_id = None

            # Extract available properties from checklist items
            self.extract_available_properties()
            
            # Get available Revit parameters
            revit_parameters = revit_utils.get_available_parameters()
            self.revitParametersComboBox.ItemsSource = revit_parameters
            
            # Enable parameter tab
            self.parameterTab.IsEnabled = True
            self.tabControl.SelectedItem = self.parameterTab
            self.update_status("Retrieved checklist preview")
            
            # Now that parameters are loaded, try to load mapping configuration
            self.load_mapping_configuration()
        else:
            error_msg = self.streambim_client.last_error or "No items found in this checklist"
            self.update_status(error_msg)
        
        self.selectChecklistButton.IsEnabled = True
    
    def extract_available_properties(self):
        """Extract available properties from checklist items."""
        if not self.checklist_items:
            return
        
        # Get first item to extract property names
        property_sets = set()
        
        # Look for items properties
        for item in self.checklist_items:
            if 'items' in item and item['items']:
                for prop_name in item['items'].keys():
                    property_sets.add(prop_name)
        
        # Create property list with sample values
        self.streambim_properties = []
        
        # Extract sample values for each property
        for prop_name in sorted(property_sets):
            sample_values = []
            for item in self.checklist_items:
                if 'items' in item and item['items'] and prop_name in item['items']:
                    sample_value = item['items'][prop_name]
                    if sample_value and sample_value not in sample_values:
                        sample_values.append(sample_value)
                    if len(sample_values) >= 3:
                        break
            
            # Create property value object
            self.streambim_properties.append(PropertyValue(
                Name=self.streambim_client._decode_utf8(prop_name),
                Sample=', '.join(sample_values[:3]) + ('...' if len(sample_values) > 3 else '')
            ))
        
        # Update UI
        self.streamBIMPropertiesComboBox.ItemsSource = [prop.Name for prop in self.streambim_properties]
        
        if self.streambim_properties:
            self.streamBIMPropertiesComboBox.SelectedIndex = 0
    
    def streambim_property_selected(self, sender, args):
        """Handle StreamBIM property selection."""
        selected_property = self.streamBIMPropertiesComboBox.SelectedItem
        
        if not selected_property:
            return
        
        # Find the property in the list
        for prop in self.streambim_properties:
            if prop.Name == selected_property:
                self.previewValuesTextBlock.Text = prop.Sample
                break
                
        self.update_summary()
    
    def update_summary(self):
        """Update the summary text."""
        streambim_prop = self.streamBIMPropertiesComboBox.SelectedItem
        revit_param = self.revitParametersComboBox.SelectedItem
        
        if streambim_prop and revit_param:
            self.summaryTextBlock.Text = "Import '{}' from StreamBIM to Revit parameter '{}'".format(
                streambim_prop, revit_param
            )
        else:
            self.summaryTextBlock.Text = "Please select both StreamBIM property and Revit parameter."
    
    def import_button_click(self, sender, args):
        """Handle import button click."""
        streambim_prop = self.streamBIMPropertiesComboBox.SelectedItem
        revit_param = self.revitParametersComboBox.SelectedItem
        
        if not streambim_prop or not revit_param:
            self.update_status("Please select both StreamBIM property and Revit parameter")
            return
            
        # Save current mapping configuration if mapping is enabled
        if self.enableMappingCheckBox.IsChecked:
            self.save_current_mapping()
        
        # Update UI
        self.importButton.IsEnabled = False
        self.isolateButton.IsEnabled = False
        self.progressBar.Visibility = Visibility.Visible
        self.progressText.Visibility = Visibility.Visible
        self.progressBar.Value = 0
        self.progressText.Text = "Fetching checklist items..."
        self.update_status("Loading all checklist items...")
        
        # Show busy indicator during initial loading
        try:
            self.busyIndicator.IsBusy = True
            self.busyIndicator.BusyContent = "Importing checklist, please wait..."
        except:
            pass
        
        # Get all checklist items
        all_checklist_items = self.streambim_client.get_checklist_items(self.selected_checklist_id, streambim_prop, limit=0)
        if not all_checklist_items:
            error_msg = self.streambim_client.last_error or "Failed to retrieve all checklist items"
            self.update_status(error_msg)
            self.importButton.IsEnabled = True
            self.progressBar.Visibility = Visibility.Collapsed
            self.progressText.Visibility = Visibility.Collapsed
            # Hide busy indicator
            try:
                self.busyIndicator.IsBusy = False
            except:
                pass
            return
            
        self.progressBar.Value = 25
        self.progressText.Text = "Building element lookup dictionary..."
        self.update_status("Building IFC GUID lookup dictionary...")
        
        logger.debug("Building IFC GUID element lookup dictionary...")
        ifc_guid_dict = {}
        
        # Collect elements to update
        elements = []
        if self.onlyVisibleElementsCheckBox.IsChecked:
            elements = revit_utils.get_visible_elements()
        else:
            # Use all elements in the document
            elements = FilteredElementCollector(revit.doc).WhereElementIsNotElementType().ToElements()
        
        # Track how many elements have IFC GUIDs
        ifc_guid_count = 0
        
        # Build dictionary mapping IFC GUIDs to elements
        for element in elements:
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
        
        logger.debug("Found {} elements with IFC GUIDs".format(ifc_guid_count))
        
        self.progressBar.Value = 40
        self.progressText.Text = "Processing checklist items..."
        
        # Create value mapping dictionary if enabled
        value_mapping = {}
        if self.enableMappingCheckBox.IsChecked and self.mappings:
            for mapping in self.mappings:
                if mapping.ChecklistValue and mapping.RevitValue:
                    value_mapping[mapping.ChecklistValue] = mapping.RevitValue
        
        # Check if this is a grouped checklist
        is_grouped = self.selected_checklist_group_by and len(self.selected_checklist_group_by) > 0
        
        if is_grouped:
            # Grouped checklist: resolve group keys to IFC GUIDs
            if not self.selected_checklist_building_id:
                error_msg = "Cannot import grouped checklist: no building ID found"
                self.update_status(error_msg)
                logger.error(error_msg)
                self.importButton.IsEnabled = True
                self.progressBar.Visibility = Visibility.Collapsed
                self.progressText.Visibility = Visibility.Collapsed
                return
            
            self.update_status("Resolving grouped checklist items...")
            logger.info("Processing grouped checklist with group-by: {}".format(self.selected_checklist_group_by))
            
            # Build mapping from IFC GUID to property value
            guid_to_value = {}
            total_groups = len(all_checklist_items)
            group_count = 0
            
            for item in all_checklist_items:
                group_count += 1
                
                # Update progress
                if group_count % 5 == 0:
                    progress = 40 + (group_count / total_groups * 10)
                    self.progressBar.Value = progress
                    self.progressText.Text = "Resolving groups... ({}/{})".format(group_count, total_groups)
                
                try:
                    # Get group key (object value)
                    group_key = item.get('object')
                    if not group_key:
                        continue
                    
                    # Get property value from this group
                    if 'items' not in item or streambim_prop not in item['items']:
                        continue
                    
                    property_value = item['items'][streambim_prop]
                    if not property_value:
                        continue
                    
                    # Apply value mapping if enabled
                    if value_mapping:
                        if property_value in value_mapping:
                            property_value = value_mapping[property_value]
                        else:
                            # Skip if mapping enabled but value not in mapping
                            continue
                    
                    # Resolve group key to IFC GUIDs
                    ifc_guids = self.streambim_client.resolve_group_key_to_ifc_guids(
                        self.selected_checklist_id,
                        self.selected_checklist_building_id,
                        group_key
                    )
                    
                    # Map all resolved GUIDs to this property value
                    for guid in ifc_guids:
                        guid_to_value[guid] = property_value
                        
                except Exception as e:
                    logger.error("Error resolving group key: {}".format(str(e)))
                    continue
            
            logger.info("Resolved {} groups to {} IFC GUID mappings".format(total_groups, len(guid_to_value)))
            
        else:
            # Non-grouped checklist: direct GUID matching (existing behavior)
            guid_to_value = {}
            for item in all_checklist_items:
                if 'object' not in item:
                    continue
                
                guid = item['object']
                
                # Get property value
                if 'items' not in item or streambim_prop not in item['items']:
                    continue
                
                property_value = item['items'][streambim_prop]
                if not property_value:
                    continue
                
                # Apply value mapping if enabled
                if value_mapping:
                    if property_value in value_mapping:
                        property_value = value_mapping[property_value]
                    else:
                        continue
                
                guid_to_value[guid] = property_value
        
        self.progressBar.Value = 50
        self.progressText.Text = "Updating element parameters..."
        self.update_status("Importing values...")
        
        # Track progress
        processed = 0
        updated = 0
        self.updated_elements = []  # Reset updated elements list
        
        # Start a group transaction for all parameter changes
        t = Transaction(revit.doc, 'Import StreamBIM Values')
        t.Start()
        
        try:
            # Process GUID to value mappings
            total_items = len(guid_to_value)
            item_count = 0
            
            for guid, revit_value in guid_to_value.items():
                item_count += 1
                processed += 1
                
                try:
                    # Get the element directly from our dictionary (FAST LOOKUP)
                    element = ifc_guid_dict.get(guid)
                    if not element:
                        continue
                    
                    # Get the parameter
                    param = element.LookupParameter(revit_param)
                    if not param:
                        continue
                        
                    # Skip read-only parameters
                    if param.IsReadOnly:
                        continue
                        
                    # Set parameter value directly within the transaction
                    try:
                        if param.StorageType == StorageType.String:
                            param.Set(str(revit_value))
                        elif param.StorageType == StorageType.Double:
                            param.Set(float(revit_value))
                        elif param.StorageType == StorageType.Integer:
                            param.Set(int(revit_value))
                        updated += 1
                        self.updated_elements.append(element)
                    except Exception as e:
                        logger.error("Error setting parameter value for element {}: {}".format(element.Id, str(e)))
                        
                except Exception as e:
                    logger.error("Error processing element: {}".format(str(e)))
                
                # Update progress
                if item_count % 10 == 0:  # Update progress every 10 items
                    progress = 50 + (item_count / total_items * 50)
                    self.progressBar.Value = progress
                    self.progressText.Text = "Processing items... ({}/{})".format(item_count, total_items)
            
            # Commit all changes
            t.Commit()
            
            except Exception as e:
                t.RollBack()
                logger.error("Error during import: {}".format(str(e)))
                self.update_status("Error during import: {}".format(str(e)))
                # Hide busy indicator on error
                try:
                    self.busyIndicator.IsBusy = False
                except:
                    pass
                return
        
        # Update UI
        self.update_status("Import complete. Updated {}/{} elements.".format(updated, processed))
        self.importButton.IsEnabled = True
        self.isolateButton.IsEnabled = updated > 0  # Enable isolate button if elements were updated
        self.progressBar.Visibility = Visibility.Collapsed
        self.progressText.Visibility = Visibility.Collapsed
        
        # Hide busy indicator
        try:
            self.busyIndicator.IsBusy = False
        except:
            pass
    
    def isolate_button_click(self, sender, args):
        """Handle isolate button click."""
        if self.updated_elements:
            if revit_utils.isolate_elements(self.updated_elements):
                self.update_status("Isolated {} updated elements".format(len(self.updated_elements)))
            else:
                self.update_status("Failed to isolate elements")
    
    def update_status(self, message):
        """Update status text."""
        self.statusTextBlock.Text = message
        logger.debug(message)

    def projects_list_double_click(self, sender, args):
        """Handle double-click on projects list view."""
        if self.projectsListView.SelectedItem and self.selectProjectButton.IsEnabled:
            self.select_project_button_click(sender, args)
    
    def checklists_list_double_click(self, sender, args):
        """Handle double-click on checklists list view."""
        if self.checklistsListView.SelectedItem and self.selectChecklistButton.IsEnabled:
            self.select_checklist_button_click(sender, args)

    def checklist_search_changed(self, sender, args):
        """Filter checklists based on search text."""
        search_text = self.checklistSearchBox.Text.lower()
        self.checklists.Clear()
        
        if not search_text:
            # If search box is empty, show all checklists
            for checklist in self.all_checklists:
                self.checklists.Add(checklist)
        else:
            # Filter checklists by name containing the search text
            for checklist in self.all_checklists:
                if search_text in checklist.Name.lower():
                    self.checklists.Add(checklist)
        
        self.update_status("Filtered to {} checklists".format(len(self.checklists)))

    def save_project_id(self, project_id):
        """Save the project ID to extensible storage."""
        try:
            # Get or create data storage
            data_storage = get_or_create_data_storage(revit.doc)
            
            # Get current stored value (if any)
            schema = StreamBIMSettingsSchema(data_storage)
            current_id = schema.get("project_id")
            
            # Only update if different
            if current_id != project_id:
                with StreamBIMSettingsSchema(data_storage) as entity:
                    entity.set("project_id", project_id)
                self.update_status("Saved project ID: {}".format(project_id))
                
            return True
        except Exception as e:
            logger.error("Failed to save project ID: {}".format(str(e)))
            return False

    def load_saved_project_id(self):
        """Load the saved project ID from extensible storage."""
        try:
            data_storage = get_or_create_data_storage(revit.doc)
            schema = StreamBIMSettingsSchema(data_storage)
            
            if schema.is_valid:
                return schema.get("project_id")            
            
            return None
        except Exception as e:
            logger.error("Failed to load project ID: {}".format(str(e)))
            return None

    def enable_mapping_checked(self, sender, args):
        """Handle enabling the mapping feature."""
        self.mappingGrid.Visibility = Visibility.Visible
        
    def enable_mapping_unchecked(self, sender, args):
        """Handle disabling the mapping feature."""
        self.mappingGrid.Visibility = Visibility.Collapsed
        
    def save_mapping_button_click(self, sender, args):
        """Handle save mapping button click."""
        self.save_current_mapping()
        self.update_status("Mapping configuration saved")
        
    def save_current_mapping(self):
        """Save the current mapping configuration to storage."""
        streambim_prop = self.streamBIMPropertiesComboBox.SelectedItem
        revit_param = self.revitParametersComboBox.SelectedItem
        
        logger.debug("Saving mapping configuration:")
        logger.debug("- StreamBIM property: {}".format(streambim_prop))
        logger.debug("- Revit parameter: {}".format(revit_param))
        logger.debug("- Selected checklist ID: {}".format(self.selected_checklist_id))
        logger.debug("- Selected checklist name: {}".format(self.selected_checklist_name))

        # Debug print each mapping entry
        for mapping in self.mappings:
            logger.debug("- Mapping: {}".format(mapping))
        
        if not streambim_prop or not revit_param or not self.selected_checklist_id:
            self.update_status("Cannot save mapping: missing property or parameter selection")
            return False
            
        mapping_data = []

        if self.enableMappingCheckBox.IsChecked:
            # Convert mappings to JSON
            for mapping in self.mappings:
                mapping_data.append({
                    'ChecklistValue': mapping.ChecklistValue,
                    'RevitValue': mapping.RevitValue
                })
            
        mapping_json = json.dumps(mapping_data, ensure_ascii=False)
        logger.debug("Mapping JSON: {}".format(mapping_json))
        
        try:
            # Create configuration dictionary for the new config
            new_config = {
                'checklist_id': self.selected_checklist_id,
                'checklist_name': self.selected_checklist_name,
                'streambim_property': streambim_prop,
                'revit_parameter': str(revit_param),
                'mapping_enabled': self.enableMappingCheckBox.IsChecked,
                'mapping_config': mapping_json,
                'id': None  # No element ID for new configurations
            }
            
            # Load existing configurations
            existing_configs = load_configs_with_pickle(revit.doc)
            logger.debug("Loaded {} existing configurations".format(len(existing_configs)))
            
            # Check if we already have a config with the same checklist_id and property
            updated_existing = False
            for i, config in enumerate(existing_configs):
                if (config.get('checklist_id') == self.selected_checklist_id and 
                    config.get('streambim_property') == streambim_prop):
                    # Update the existing config
                    logger.debug("Updating existing configuration")
                    existing_configs[i] = new_config
                    updated_existing = True
                    break
            
            # If we didn't update an existing config, append the new one
            if not updated_existing:
                logger.debug("Adding new configuration")
                existing_configs.append(new_config)
            
            # Save all configurations
            success = save_configs_with_pickle(revit.doc, existing_configs)
            
            if success:
                self.update_status("Mapping configuration saved")
                return True
            else:
                self.update_status("Failed to save mapping configuration")
                return False
                
        except Exception as e:
            logger.error("Failed to save mapping configuration: {}".format(str(e)))
            self.update_status("Error saving configuration: {}".format(str(e)))
            return False

    def load_mapping_configuration(self):
        """Load mapping configuration from extensible storage if available."""
        if self.is_loading_configuration:
            logger.debug("Already loading configuration, skipping")
            return False
            
        self.is_loading_configuration = True
        try:
            # Check if we have a valid Revit document
            if not revit.doc:
                logger.error("No active Revit document")
                return False
                
            streambim_prop = self.streamBIMPropertiesComboBox.SelectedItem
            logger.debug("Loading configuration for StreamBIM property: {}".format(streambim_prop))

            # Clear the mappings list
            self.mappings.Clear()
            self.enableMappingCheckBox.IsChecked = False
            self.mappingGrid.Visibility = Visibility.Collapsed

            if not streambim_prop or not self.selected_checklist_id:
                logger.debug("No StreamBIM property or checklist ID selected")
                return False
                
            # Load configurations from new pickle-based storage
            loaded_configs = load_configs_with_pickle(revit.doc)
            logger.debug("Loaded {} configurations from storage".format(len(loaded_configs)))
            
            # Find a matching configuration
            matching_config = None
            for config in loaded_configs:
                logger.debug("Checking config: checklist_id={}, property={}".format(
                    config.get('checklist_id'), config.get('streambim_property')))
                
                if (config.get('checklist_id') == self.selected_checklist_id and 
                    config.get('streambim_property') == streambim_prop):
                    matching_config = config
                    logger.debug("Found matching configuration")
                    break
            
            if matching_config:
                # Found matching configuration
                revit_parameter = matching_config.get('revit_parameter')
                mapping_config = matching_config.get('mapping_config')
                mapping_enabled = matching_config.get('mapping_enabled')

                logger.debug("Found matching configuration:")
                logger.debug("- Checklist ID: {}".format(matching_config.get('checklist_id')))
                logger.debug("- Checklist name: {}".format(matching_config.get('checklist_name')))
                logger.debug("- StreamBIM property: {}".format(matching_config.get('streambim_property')))
                logger.debug("- Revit parameter to restore: {}".format(revit_parameter))
                logger.debug("- Enable mapping: {}".format(mapping_enabled))
                
                # Only try to restore parameter selection if we have parameters loaded
                if revit_parameter and self.revitParametersComboBox.Items and self.revitParametersComboBox.Items.Count > 0:
                    logger.debug("Available Revit parameters:")
                    for index, item in enumerate(self.revitParametersComboBox.Items):
                        logger.debug("- Item: '{}' (type: {})".format(item, type(item)))
                        if str(item).strip() == str(revit_parameter).strip():
                            logger.debug("Found matching parameter: {} at index {}".format(item, index))
                            self.revitParametersComboBox.SelectedIndex = index
                            break
                
                # Restore mapping data
                self.mappings.Clear()
                if mapping_config:
                    try:
                        mapping_data = json.loads(mapping_config)
                        for mapping in mapping_data:
                            self.mappings.Add(MappingEntry(
                                ChecklistValue=mapping.get('ChecklistValue', ''),
                                RevitValue=mapping.get('RevitValue', '')
                            ))
                        logger.debug("Restored {} mapping entries".format(len(mapping_data)))
                    except Exception as e:
                        logger.error("Error parsing mapping JSON: {}".format(str(e)))
                
                # Enable mapping checkbox if it was enabled
                self.enableMappingCheckBox.IsChecked = mapping_enabled if mapping_enabled is not None else False
                self.mappingGrid.Visibility = Visibility.Visible if self.enableMappingCheckBox.IsChecked else Visibility.Collapsed
                self.update_status("Loaded mapping configuration")
                logger.debug("Loaded mapping configuration")

                return True
            else:
                logger.debug("No matching configuration found in new storage")
            
            return False
        except Exception as e:
            logger.error("Failed to load mapping configuration: {}".format(str(e)))
            return False
        finally:
            self.is_loading_configuration = False

    def revit_parameter_selected(self, sender, args):
        """Handle Revit parameter selection."""
        selected_parameter = self.revitParametersComboBox.SelectedItem
        
        if not selected_parameter:
            return
        
        self.update_summary()

    def add_new_row_button_click(self, sender, args):
        """Add a new empty row to the mapping DataGrid."""
        self.mappings.Add(MappingEntry(ChecklistValue="", RevitValue=""))
        self.update_status("Added new mapping row")

    def remove_row_button_click(self, sender, args):
        """Remove the selected row from the mapping DataGrid."""
        if self.mappingDataGrid.SelectedItem:
            self.mappings.Remove(self.mappingDataGrid.SelectedItem)
            self.update_status("Removed selected mapping row")

    def save_button_click(self, sender, args):
        """Save the current configuration when the Save button is clicked."""
        if not self.current_checklist:
            self.update_status("No checklist selected")
            return
            
        if not self.selected_property:
            self.update_status("No property selected")
            return
            
        if not self.selected_parameter:
            self.update_status("No parameter selected")
            return
            
        # Handle mapping data
        mapping_data = []
        for mapping in self.mappings:
            if mapping.ChecklistValue and mapping.RevitValue:
                mapping_data.append({
                    'ChecklistValue': mapping.ChecklistValue,
                    'RevitValue': mapping.RevitValue
                })
        
        self.mapping_config = json.dumps(mapping_data)
        
        try:
            # Create configuration dictionary for the new config
            new_config = {
                'checklist_id': self.current_checklist['id'],
                'checklist_name': self.current_checklist['attributes']['name'],
                'streambim_property': self.selected_property,
                'revit_parameter': self.selected_parameter,
                'mapping_enabled': self.enableMappingCheckBox.IsChecked == True,
                'mapping_config': self.mapping_config
            }
            
            # Load existing configurations
            existing_configs = load_configs_with_pickle(revit.doc)
            logger.debug("Loaded {} existing configurations".format(len(existing_configs)))
            
            # Check if we already have a config with the same checklist_id and property
            updated_existing = False
            for i, config in enumerate(existing_configs):
                if (config.get('checklist_id') == self.current_checklist['id'] and 
                    config.get('streambim_property') == self.selected_property):
                    # Update the existing config
                    logger.debug("Updating existing configuration")
                    existing_configs[i] = new_config
                    updated_existing = True
                    break
            
            # If we didn't update an existing config, append the new one
            if not updated_existing:
                logger.debug("Adding new configuration")
                existing_configs.append(new_config)
            
            # Save all configurations
            success = save_configs_with_pickle(revit.doc, existing_configs)
            
            if success:
                self.update_status("Configuration saved")
            else:
                self.update_status("Error saving configuration")
                
        except Exception as e:
            logger.error("Error saving configuration: {}".format(str(e)))
            self.update_status("Error saving configuration: {}".format(str(e)))

# Main execution
if __name__ == '__main__':
    # Show the StreamBIM Importer UI
    StreamBIMImporterUI().ShowDialog() 