# -*- coding: utf-8 -*-
"""Import StreamBIM data into Revit parameters.

This tool connects to StreamBIM to import parameter values from StreamBIM checklists
into Revit parameters for elements in the current view.
"""
__title__ = "StreamBIM\nImporter"
__author__ = "Your Company"
__doc__ = "Import StreamBIM checklist data into Revit parameters"

import os
import sys
import clr
import json
from collections import namedtuple

# Add the extension directory to the path
import os.path as op
extension_dir = op.dirname(op.dirname(op.dirname(op.dirname(__file__))))
lib_path = op.join(extension_dir, 'pyBS.lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)

# Add reference to WPF
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import EventHandler
from System.Collections.ObjectModel import ObservableCollection
from System.Dynamic import ExpandoObject
from System.Windows import MessageBox, MessageBoxButton

from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Import custom modules from the extension lib
from streambim import streambim_api
from streambim import revit_utils

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
        forms.WPFWindow.__init__(self, 'StreamBIMImporter.xaml')
        
        # Initialize StreamBIM API client
        self.streambim_client = streambim_api.StreamBIMClient()
        
        # Initialize data collections
        self.projects = ObservableCollection[object]()
        self.checklists = ObservableCollection[object]()
        self.checklist_items = []
        self.streambim_properties = []
        
        # Set up event handlers
        self.loginButton.Click += self.login_button_click
        self.selectProjectButton.Click += self.select_project_button_click
        self.selectChecklistButton.Click += self.select_checklist_button_click
        self.streamBIMPropertiesComboBox.SelectionChanged += self.streambim_property_selected
        self.importButton.Click += self.import_button_click
        
        # Initialize UI
        self.projectsListView.ItemsSource = self.projects
        self.checklistsListView.ItemsSource = self.checklists
        
        # Set status
        self.update_status("Ready to connect to StreamBIM")
    
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
            self.update_status("Login successful. Retrieving projects...")
            self.loginStatusTextBlock.Text = "Logged in as: " + username
            
            # Get projects
            projects = self.streambim_client.get_projects()
            if projects:
                self.projects.Clear()
                for project in projects:
                    self.projects.Add(Project(
                        Id=str(project.get('id')),
                        Name=project.get('name', 'Unknown'),
                        Description=project.get('description', '')
                    ))
                
                # Enable project tab
                self.projectTab.IsEnabled = True
                self.tabControl.SelectedItem = self.projectTab
                self.update_status("Retrieved {} projects".format(len(projects)))
            else:
                error_msg = self.streambim_client.last_error or "No projects found"
                self.update_status(error_msg)
        else:
            error_msg = self.streambim_client.last_error or "Login failed. Please check your credentials."
            self.update_status(error_msg)
            self.loginStatusTextBlock.Text = "Login failed: " + error_msg
        
        self.loginButton.IsEnabled = True
    
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
        self.update_status("Loading checklist items...")
        
        # Get checklist items
        checklist_items = self.streambim_client.get_checklist_items(selected_checklist.Id)
        if checklist_items:
            self.checklist_items = checklist_items
            
            # Extract available properties from checklist items
            self.extract_available_properties()
            
            # Get available Revit parameters
            revit_parameters = revit_utils.get_available_parameters()
            self.revitParametersComboBox.ItemsSource = revit_parameters
            
            # Enable parameter tab
            self.parameterTab.IsEnabled = True
            self.tabControl.SelectedItem = self.parameterTab
            self.update_status("Retrieved {} checklist items".format(len(checklist_items)))
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
                Name=prop_name,
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
                self.previewValuesTextBlock.Text = "Sample values: " + prop.Sample
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
        for item in self.checklist_items:
            if 'object' in item:
                guid_to_item[item['object']] = item
        
        # Track progress
        processed = 0
        updated = 0
        
        # Process elements
        for element in elements:
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
                
                # Set parameter value
                if revit_utils.set_parameter_value(element, revit_param, value):
                    updated += 1
            except Exception as e:
                logger.error("Error processing element: {}".format(str(e)))
        
        # Update UI
        self.update_status("Import complete. Updated {}/{} elements.".format(updated, processed))
        self.importButton.IsEnabled = True
        
        # Show results
        MessageBox.Show(
            "Import complete.\n\nProcessed: {} elements\nUpdated: {} elements".format(processed, updated),
            "Import Results",
            MessageBoxButton.OK
        )
    
    def update_status(self, message):
        """Update status text."""
        self.statusTextBlock.Text = message
        logger.debug(message)

# Main execution
if __name__ == '__main__':
    # Show the StreamBIM Importer UI
    StreamBIMImporterUI().ShowDialog() 