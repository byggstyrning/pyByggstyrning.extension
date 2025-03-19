# -*- coding: utf-8 -*-
"""Import StreamBIM data into Revit parameters.

This tool connects to StreamBIM to import parameter values from StreamBIM checklists
into Revit parameters for elements in the current view.
"""
__title__ = "Checklist\nImporter"
__author__ = "Byggstyrning AB"
__doc__ = "Import StreamBIM checklist items data into Revit instance parameters"

## todo:
# - remember project selection
# - remember checklist selection
# - search on project and checklist name
# - optional value mapping (e.g. "X" -> "500", "No" -> "Nej")


import os
import sys
import clr
import json
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
from Autodesk.Revit.DB import *

from System import EventHandler
from System.Collections.ObjectModel import ObservableCollection
from System.Dynamic import ExpandoObject
from System.Windows import MessageBox, MessageBoxButton, Visibility

from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Import custom modules from the extension lib
from streambim import streambim_api
from revit import revit_utils

# Import extensible storage
from extensible_storage import BaseSchema, simple_field

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

class StreamBIMSchema(BaseSchema):
    """Schema for storing StreamBIM settings"""
    
    guid = "4ad08ded-abdc-4ff6-b601-b7f9b3916f32"
    
    @simple_field(value_type="string")
    def project_id():
        """The selected StreamBIM project ID"""

class MappingSchema(BaseSchema):
    """Schema for storing parameter mapping configurations"""
    
    guid = "8ee65b71-f13c-492a-9cf8-721ac98d4f7b"
    
    @simple_field(value_type="string")
    def checklist_id():
        """The selected checklist ID"""
        
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
        entity = ds.GetEntity(StreamBIMSchema.schema)
        if entity.IsValid():
            return ds
    
    # If not found, create a new one
    with revit.Transaction("Create StreamBIM Storage", doc):
        return ExtensibleStorage.DataStorage.Create(doc)

def get_or_create_mapping_storage(doc):
    """Get existing or create new mapping data storage element."""
    data_storage = FilteredElementCollector(doc)\
        .OfClass(ExtensibleStorage.DataStorage)\
        .ToElements()
    
    # Look for our storage with the schema
    for ds in data_storage:
        # Check if this storage has our schema
        entity = ds.GetEntity(MappingSchema.schema)
        if entity.IsValid():
            return ds
    
    # If not found, create a new one
    with revit.Transaction("Create StreamBIM Mapping Storage", doc):
        return ExtensibleStorage.DataStorage.Create(doc)

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
        self.passwordBox.Password = ""  # Clear password for security
        
        # Enable project tab and switch to it
        self.projectTab.IsEnabled = True
        self.tabControl.SelectedItem = self.projectTab
        
        # Get projects
        self.update_status("Retrieving projects...")
        projects = self.streambim_client.get_projects()
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
        
        # Set server URL and login
        self.streambim_client.base_url = server_url
        success = self.streambim_client.login(username, password)
        
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
        
        # Save project ID
        if self.save_project_id(selected_project.Id):
            self.saved_project_id = selected_project.Id
        
        # Set current project
        self.streambim_client.set_current_project(selected_project.Id)
        
        # Get checklists
        checklists = self.streambim_client.get_checklists()
        if checklists:
            self.checklists.Clear()
            self.all_checklists = []  # Clear all checklists list
            
            for checklist in checklists:
                checklist_obj = Checklist(
                    Id=checklist.get('id'),
                    Name=checklist.get('attributes', {}).get('name', 'Unknown')
                )
                self.checklists.Add(checklist_obj)
                self.all_checklists.append(checklist_obj)  # Store in all checklists list for filtering
            
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
        
        # Update selected checklist text
        self.selectedChecklistTextBlock.Text = selected_checklist.Name
        
        # Get checklist items (limited for preview)
        checklist_items = self.streambim_client.get_checklist_items(selected_checklist.Id, limit=1)
        if checklist_items:
            self.checklist_items = checklist_items
            self.selected_checklist_id = selected_checklist.Id  # Store for later use
            
            # Extract available properties from checklist items
            self.extract_available_properties()
            
            # Get available Revit parameters
            revit_parameters = revit_utils.get_available_parameters()
            self.revitParametersComboBox.ItemsSource = revit_parameters
            
            # Enable parameter tab
            self.parameterTab.IsEnabled = True
            self.tabControl.SelectedItem = self.parameterTab
            self.update_status("Retrieved checklist preview")
        else:
            error_msg = self.streambim_client.last_error or "No items found in this checklist"
            self.update_status(error_msg)

        self.load_mapping_configuration()
        
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
        
        # Get all checklist items
        all_checklist_items = self.streambim_client.get_checklist_items(self.selected_checklist_id, limit=0)
        if not all_checklist_items:
            error_msg = self.streambim_client.last_error or "Failed to retrieve all checklist items"
            self.update_status(error_msg)
            self.importButton.IsEnabled = True
            self.progressBar.Visibility = Visibility.Collapsed
            self.progressText.Visibility = Visibility.Collapsed
            return
            
        self.progressBar.Value = 50
        self.progressText.Text = "Processing elements..."
        self.update_status("Importing values...")
        
        # Collect elements to update
        elements = []
        if self.onlyVisibleElementsCheckBox.IsChecked:
            elements = revit_utils.get_visible_elements()
        else:
            # Use all elements in the document
            elements = FilteredElementCollector(revit.doc).WhereElementIsNotElementType().ToElements()
        
        # Create mapping from IFC GUID to checklist item
        guid_to_item = {}
        for item in all_checklist_items:
            if 'object' in item:
                guid_to_item[item['object']] = item
        
        # Create value mapping dictionary if enabled
        value_mapping = {}
        if self.enableMappingCheckBox.IsChecked and self.mappings:
            for mapping in self.mappings:
                if mapping.ChecklistValue and mapping.RevitValue:
                    value_mapping[mapping.ChecklistValue] = mapping.RevitValue
        
        # Track progress
        processed = 0
        updated = 0
        self.updated_elements = []  # Reset updated elements list
        
        # Start a group transaction for all parameter changes
        t = Transaction(revit.doc, 'Import StreamBIM Values')
        t.Start()
        
        try:
            # Process elements
            total_elements = len(elements)
            for i, element in enumerate(elements):
                processed += 1
                
                try:
                    # Try to get IFC GUID from element - check different parameter names
                    ifc_guid_param = element.LookupParameter("IFCGuid")
                    if not ifc_guid_param:
                        ifc_guid_param = element.LookupParameter("IfcGUID")
                    if not ifc_guid_param:
                        ifc_guid_param = element.LookupParameter("IFC GUID")
                    
                    if not ifc_guid_param:
                        continue
                    
                    ifc_guid = ifc_guid_param.AsString()
                    if not ifc_guid or ifc_guid not in guid_to_item:
                        continue
                    
                    # Get item data
                    item = guid_to_item[ifc_guid]
                    if 'items' not in item or streambim_prop not in item['items']:
                        continue
                    
                    # Get property value
                    checklist_value = item['items'][streambim_prop]
                    
                    # Apply value mapping if enabled and value exists in mapping
                    if value_mapping and checklist_value in value_mapping:
                        revit_value = value_mapping[checklist_value]
                    else:
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
                progress = 50 + (i / total_elements * 50)
                self.progressBar.Value = progress
                self.progressText.Text = "Processing elements... ({}/{})".format(i + 1, total_elements)
            
            # Commit all changes
            t.Commit()
            
        except Exception as e:
            t.RollBack()
            logger.error("Error during import: {}".format(str(e)))
            self.update_status("Error during import: {}".format(str(e)))
            return
        
        # Update UI
        self.update_status("Import complete. Updated {}/{} elements.".format(updated, processed))
        self.importButton.IsEnabled = True
        self.isolateButton.IsEnabled = updated > 0  # Enable isolate button if elements were updated
        self.progressBar.Visibility = Visibility.Collapsed
        self.progressText.Visibility = Visibility.Collapsed
    
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
            schema = StreamBIMSchema(data_storage)
            current_id = schema.get("project_id")
            
            # Only update if different
            if current_id != project_id:
                with StreamBIMSchema(data_storage) as entity:
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
            schema = StreamBIMSchema(data_storage)
            
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
        """Save the current mapping configuration to extensible storage."""
        streambim_prop = self.streamBIMPropertiesComboBox.SelectedItem
        revit_param = self.revitParametersComboBox.SelectedItem
        
        logger.debug("Saving mapping configuration:")
        logger.debug("- StreamBIM property: {}".format(streambim_prop))
        logger.debug("- Revit parameter: {}".format(revit_param))
        logger.debug("- Selected checklist ID: {}".format(self.selected_checklist_id))
        
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
            
        mapping_json = json.dumps(mapping_data)
        logger.debug("Mapping JSON: {}".format(mapping_json))
        
        try:
            # Get or create data storage
            mapping_storage = get_or_create_mapping_storage(revit.doc)
            
            # Save mapping configuration
            with MappingSchema(mapping_storage) as entity:
                entity.set("checklist_id", self.selected_checklist_id)
                entity.set("streambim_property", streambim_prop)
                entity.set("revit_parameter", str(revit_param))
                entity.set("enableMapping", str(self.enableMappingCheckBox.IsChecked))
                entity.set("mapping_config", mapping_json)
                
            self.mapping_storage = mapping_storage
            self.update_status("Mapping configuration saved: {}".format(mapping_json))
            return True
        except Exception as e:
            logger.error("Failed to save mapping configuration: {}".format(str(e)))
            return False
            
    def load_mapping_configuration(self):
        """Load mapping configuration from extensible storage if available."""
        if self.is_loading_configuration:
            return False
            
        self.is_loading_configuration = True
        try:
            streambim_prop = self.streamBIMPropertiesComboBox.SelectedItem
            revit_param = self.revitParametersComboBox.SelectedItem

            # Clear the mappings list
            self.mappings.Clear()
            self.enableMappingCheckBox.IsChecked = False
            self.mappingGrid.Visibility = Visibility.Collapsed

            if not streambim_prop or not self.selected_checklist_id:
                return False
                
            # Get mapping storage
            mapping_storage = get_or_create_mapping_storage(revit.doc)
            
            # Check if we have a matching configuration
            schema = MappingSchema(mapping_storage)
            
            if schema.is_valid:
                checklist_id = schema.get("checklist_id")
                prop = schema.get("streambim_property")
                
                if checklist_id == self.selected_checklist_id and prop == streambim_prop:
                    # Found matching configuration
                    mapping_json = schema.get("mapping_config")
                    revit_parameter = schema.get("revit_parameter")
                    enable_mapping = schema.get("enableMapping")
                    
                    logger.debug("Found matching configuration:")
                    logger.debug("- Checklist ID: {}".format(checklist_id))
                    logger.debug("- StreamBIM property: {}".format(prop))
                    logger.debug("- Revit parameter to restore: {}".format(revit_parameter))
                    logger.debug("- Enable mapping: {}".format(enable_mapping))
                    
                    # Restore property selection
                    if revit_parameter:
                        logger.debug("Available Revit parameters:")
                        for index, item in enumerate(self.revitParametersComboBox.Items):
                            logger.debug("- Item: '{}' (type: {})".format(item, type(item)))
                            if str(item).strip() == str(revit_parameter).strip():
                                logger.debug("Found matching parameter: {} at index {}".format(item, index))
                                self.revitParametersComboBox.SelectedIndex = index
                                break
                    
                    # Restore mapping data
                    self.mappings.Clear()
                    if mapping_json:
                        mapping_data = json.loads(mapping_json)
                        for mapping in mapping_data:
                            self.mappings.Add(MappingEntry(
                                ChecklistValue=mapping.get('ChecklistValue', ''),
                                RevitValue=mapping.get('RevitValue', '')
                            ))
                    
                    # Enable mapping checkbox if it was enabled
                    self.enableMappingCheckBox.IsChecked = enable_mapping.lower() == "true" if enable_mapping else False
                    self.mappingGrid.Visibility = Visibility.Visible if self.enableMappingCheckBox.IsChecked else Visibility.Collapsed
                    self.update_status("Loaded mapping configuration")
                    logger.debug("Loaded mapping configuration: {}".format(mapping_json))

                    return True
            
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

# Main execution
if __name__ == '__main__':
    # Show the StreamBIM Importer UI
    StreamBIMImporterUI().ShowDialog() 