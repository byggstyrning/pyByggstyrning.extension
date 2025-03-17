# -*- coding: utf-8 -*-
"""Import StreamBIM data into Revit parameters.

This tool connects to StreamBIM to import parameter values from StreamBIM checklists
into Revit parameters for elements in the current view.
"""
__title__ = "Checklist\nImporter"
__author__ = "Byggstyrning AB"
__doc__ = "Import StreamBIM checklist items data into Revit instance parameters"

import os
import sys
import clr
import json
from collections import namedtuple

# Add the extension directory to the path
import os.path as op
extension_dir = op.dirname(op.dirname(op.dirname(op.dirname(__file__))))
lib_path = op.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)

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
from lib.streambim import streambim
from lib.revit import revit_utils

# Initialize logger
logger = script.get_logger()

# Define data classes using namedtuple for IronPython
Project = namedtuple('Project', ['Id', 'Name', 'Description'])
Checklist = namedtuple('Checklist', ['Id', 'Name'])
PropertyValue = namedtuple('PropertyValue', ['Name', 'Sample'])

class StreamBIMImporterUI(forms.WPFWindow):
    """StreamBIM Importer UI implementation."""
    
    def __init__(self):
        """Initialize the StreamBIM Importer UI."""
        # Initialize WPF window
        forms.WPFWindow.__init__(self, 'ChecklistImporter.xaml')
        
        # Initialize StreamBIM API client
        self.streambim_client = streambim.StreamBIMClient()
                
        # Initialize data collections
        self.projects = ObservableCollection[object]()
        self.checklists = ObservableCollection[object]()
        self.checklist_items = []
        self.streambim_properties = []
        self.updated_elements = []  # Store updated elements for isolation

        # Set up event handlers
        self.loginButton.Click += self.login_button_click
        self.logoutButton.Click += self.logout_button_click
        self.selectProjectButton.Click += self.select_project_button_click
        self.selectChecklistButton.Click += self.select_checklist_button_click
        self.streamBIMPropertiesComboBox.SelectionChanged += self.streambim_property_selected
        self.importButton.Click += self.import_button_click
        self.isolateButton.Click += self.isolate_button_click
        
        # Initialize UI
        self.projectsListView.ItemsSource = self.projects
        self.checklistsListView.ItemsSource = self.checklists
        
        # Try automatic login if we have saved tokens
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
            for project in projects:
                # Handle new project-links data structure
                attrs = project.get('attributes', {})
                self.projects.Add(Project(
                    Id=str(project.get('id')),
                    Name=attrs.get('name', 'Unknown'),
                    Description=attrs.get('description', '')
                ))
            self.update_status("Retrieved {} projects".format(len(projects)))
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
        
        # Set current project
        self.streambim_client.set_current_project(selected_project.Id)
        
        # Get checklists
        checklists = self.streambim_client.get_checklists()
        if checklists:
            self.checklists.Clear()
            for checklist in checklists:
                self.checklists.Add(Checklist(
                    Id=checklist.get('id'),
                    Name=checklist.get('attributes', {}).get('name', 'Unknown')
                ))
            
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
                    value = item['items'][streambim_prop]
                    
                    # Get the parameter
                    param = element.LookupParameter(revit_param)
                    if not param:
                        continue
                        
                    # Set parameter value directly within the transaction
                    try:
                        if param.StorageType == StorageType.String:
                            param.Set(str(value))
                        elif param.StorageType == StorageType.Double:
                            param.Set(float(value))
                        elif param.StorageType == StorageType.Integer:
                            param.Set(int(value))
                        updated += 1
                        self.updated_elements.append(element)
                    except Exception as e:
                        logger.error("Error setting parameter value: {}".format(str(e)))
                        
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
        
        # Show results
        MessageBox.Show(
            "Import complete.\n\nProcessed: {} elements\nUpdated: {} elements".format(processed, updated),
            "Import Results",
            MessageBoxButton.OK
        )
    
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

# Main execution
if __name__ == '__main__':
    # Show the StreamBIM Importer UI
    StreamBIMImporterUI().ShowDialog() 