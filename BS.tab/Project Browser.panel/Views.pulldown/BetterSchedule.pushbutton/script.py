# -*- coding: utf-8 -*-
__title__ = "AI Schedule Processor"
__author__ = "Jonatan Jacobsson"
__doc__ = """
This tool allows you to:
1. Select a Scheduled View in Revit
2. Extract all Revit Types and their properties
3. Process the data with AI in a non-blocking UI
4. Update the Revit model with AI-generated content

Instructions:
- Click the button to launch the tool
- Select a schedule from the list
- Configure AI settings and prompts
- Process data and review AI suggestions
- Save accepted changes back to Revit
"""

# Enable persistent engine to keep the script alive for modeless forms
__persistentengine__ = True

# Import PyRevit and .NET modules
from pyrevit import forms, revit, DB, script
from pyrevit.forms import WPFWindow
import wpf
import clr
import threading
import json
import sys
import os
import time
import System
import traceback
import urllib2
import base64
import re


# Add this class to handle HTTP responses similar to requests library
class HttpResponse:
    def __init__(self, response):
        self.response = response
        self.status_code = response.getcode()
        self.headers = dict(response.info().items())
        self._content = None
        
    @property
    def text(self):
        if self._content is None:
            self._content = self.response.read()
        return self._content
        
    def json(self):
        return json.loads(self.text)

# Add this class to replace the requests library
class HttpClient:
    def post(self, url, json=None, headers=None, timeout=60):
        try:
            # Convert JSON data to string
            data = None
            if json is not None:
                data = json_dumps(json).encode('utf-8')
            
            # Create request
            request = urllib2.Request(url, data)
            
            # Add headers
            if headers:
                for key, value in headers.items():
                    request.add_header(key, value)
            
            # Set content type if not in headers
            if headers is None or 'Content-Type' not in headers:
                request.add_header('Content-Type', 'application/json')
            
            # Open the URL with timeout
            response = urllib2.urlopen(request, timeout=timeout)
            
            # Return wrapped response
            return HttpResponse(response)
        except urllib2.HTTPError as e:
            # Handle HTTP errors (4xx, 5xx)
            return HttpResponse(e)
        except Exception as e:
            # Re-raise other exceptions
            raise

# Helper function to handle JSON serialization
def json_dumps(obj):
    return json.dumps(obj, ensure_ascii=False)


# Import System namespaces
from System import Action, EventHandler, Uri, IO
from System.Windows import Threading
from System.Windows.Controls import DataGrid, DataGridTextColumn, DataGridLength, DataGridLengthUnitType, DataGridCheckBoxColumn, DataGridTemplateColumn, Button
from System.Windows.Data import Binding, BindingMode
from System.Windows.Markup import XamlReader
from System.Windows import FrameworkElement, RoutedEventHandler, Thickness, DataTemplate, PropertyPath, FrameworkElementFactory
from System.Windows.Controls.Primitives import DataGridColumnHeader
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Text import Encoding
from System.Threading import Thread, ThreadStart, AutoResetEvent
from System.Net import WebClient, WebRequest, WebResponse, HttpWebRequest, HttpWebResponse
from System.Windows.Media import Brushes
from System.Windows import Visibility
from System.Dynamic import ExpandoObject
from System.Collections.Generic import Dictionary

# Import Revit API
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Get the current Revit document and application
doc = revit.doc
uidoc = revit.uidoc
# app = revit.host  # This line causes the error

# Set up logging
logger = script.get_logger()


# Define the External Event Handler for Revit API operations
class RevitEventHandler(IExternalEventHandler):
    """External Event Handler for Revit API operations"""
    
    def __init__(self, window):
        """Initialize with a reference to the main window"""
        self.window = window
    
    def Execute(self, uiapp):
        """Execute the Revit API operation"""
        try:
            # Call the window's method to update the Revit document
            self.window.update_revit_document(uiapp)
        except Exception as ex:
            logger.error("Error in Revit event handler: {}".format(ex))
    
    def GetName(self):
        """Get the name of the event handler"""
        return "ScheduleAIProcessor_RevitEventHandler"

# Main window class
class ScheduleAIProcessorWindow(WPFWindow):
    def __init__(self):
        """Initialize the window and set up event handlers"""
        try:
            # Initialize the WPF window
            xaml_file = os.path.join(os.path.dirname(__file__), 'ScheduleAIProcessor.xaml')
            super(ScheduleAIProcessorWindow, self).__init__(xaml_file)
            
            # Initialize the schedule data collection - using a list of dictionaries instead of custom objects
            self.schedule_data = ObservableCollection[Dictionary[str, object]]()
            
            # Initialize state variables
            self.selected_schedule = None
            self.is_type_mode = True  # Default to Types mode
            self.revit_event_handler = RevitEventHandler(self)
            self.external_event = ExternalEvent.Create(self.revit_event_handler)
            
            # Initialize target_parameter_combo reference
            self.target_parameter_combo = None
            
            # Set up the DataGrid programmatically instead of relying on XAML
            self.setup_data_grid()
            
            # Set up event handlers
            self.cboScheduleSelector.SelectionChanged += self.cboScheduleSelector_SelectionChanged
            self.rbTypes.Checked += self.element_selection_changed
            self.rbInstances.Checked += self.element_selection_changed
            
            # Set up event handlers for prompt text changes
            self.txtSystemPrompt.TextChanged += self.prompt_text_changed
            self.txtUserPrompt.TextChanged += self.prompt_text_changed
            
            # Initialize AI endpoints dictionary
            self.ai_endpoints = {
                "OpenAI o3-mini (OpenRouter)": {
                    "url": "https://openrouter.ai/api/v1/chat/completions",
                    "model": "openai/o3-mini"
                },
                "Claude (OpenRouter)": {
                    "url": "https://openrouter.ai/api/v1/chat/completions",
                    "model": "anthropic/claude-3-opus"
                },
                "Gemini (OpenRouter)": {
                    "url": "https://openrouter.ai/api/v1/chat/completions",
                    "model": "google/gemini-pro"
                },
                "Custom API": {
                    "url": "https://api.openai.com/v1/chat/completions",
                    "model": "gpt-4"
                }
            }
            
            # Initialize HTTP client for API requests
            self.http_client = HttpClient()
            logger.info("Initialized custom HTTP client")
            
            # Populate the schedule dropdown
            self.populate_schedule_dropdown()
            
            # Check if active view is a schedule and select it
            self.select_active_schedule()
            
            # Initialize AI configuration
            self.initialize_ai_config()
            
            # Load API key from settings
            self.load_api_key()
            
            # Set initial status
            self.txtStatus.Text = "Ready for schedule selection."
            
            # Disable buttons until we have data
            self.btnProcessAll.IsEnabled = False
            self.btnSave.IsEnabled = False
            self.btnAcceptAll.IsEnabled = False
            
            logger.info("ScheduleAIProcessor window initialized")
            
            # Initialize flags to prevent recursive scroll events
            self._system_prompt_scrolling = False
            self._system_preview_scrolling = False
            self._user_prompt_scrolling = False
            self._user_preview_scrolling = False
            
            # Set up scroll synchronization
            self.setup_scroll_sync()
            
        except Exception as ex:
            logger.error("Error initializing window: {}".format(ex))
            forms.alert("Error initializing window: {}".format(ex), title="Error")
    
    def setup_scroll_sync(self):
        """Set up scroll synchronization between prompt and preview TextBoxes"""
        try:
            # Add TextChanged handlers for system prompt
            self.txtSystemPrompt.TextChanged += self.sync_system_scroll
            
            # Add TextChanged handlers for user prompt
            self.txtUserPrompt.TextChanged += self.sync_user_scroll
            
        except Exception as ex:
            logger.error("Error setting up scroll sync: {}".format(ex))

    def sync_system_scroll(self, sender, e):
        """Synchronize scrolling for system prompt TextBoxes"""
        try:
            # Get ScrollViewer from both TextBoxes
            prompt_viewer = self.get_scroll_viewer(self.txtSystemPrompt)
            preview_viewer = self.get_scroll_viewer(self.txtSystemPromptPreview)
            
            if prompt_viewer and preview_viewer:
                # Sync vertical offset
                if sender == self.txtSystemPrompt:
                    preview_viewer.ScrollToVerticalOffset(prompt_viewer.VerticalOffset)
                else:
                    prompt_viewer.ScrollToVerticalOffset(preview_viewer.VerticalOffset)
        except Exception as ex:
            logger.error("Error syncing system scroll: {}".format(ex))

    def sync_user_scroll(self, sender, e):
        """Synchronize scrolling for user prompt TextBoxes"""
        try:
            # Get ScrollViewer from both TextBoxes
            prompt_viewer = self.get_scroll_viewer(self.txtUserPrompt)
            preview_viewer = self.get_scroll_viewer(self.txtUserPromptPreview)
            
            if prompt_viewer and preview_viewer:
                # Sync vertical offset
                if sender == self.txtUserPrompt:
                    preview_viewer.ScrollToVerticalOffset(prompt_viewer.VerticalOffset)
                else:
                    prompt_viewer.ScrollToVerticalOffset(preview_viewer.VerticalOffset)
        except Exception as ex:
            logger.error("Error syncing user scroll: {}".format(ex))

    def get_scroll_viewer(self, textbox):
        """Helper method to get ScrollViewer from a TextBox"""
        try:
            if textbox and textbox.Template:
                return textbox.Template.FindName("PART_ContentHost", textbox)
        except Exception as ex:
            logger.error("Error getting ScrollViewer: {}".format(ex))
        return None
    
    def save_settings(self):
        """Save current settings to JSON file"""
        try:
            settings = {
                "api_key": self.txtAPIKey.Password
            }
            
            # Get the directory of the current script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            settings_path = os.path.join(script_dir, "settings.json")
            
            with open(settings_path, 'w') as f:
                json.dump(settings, f)
            
            logger.info("Settings saved successfully")
            return True
        except Exception as ex:
            logger.error("Error saving settings: {}".format(ex))
            return False

    def load_api_key(self):
        """Load saved API key from settings"""
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            settings_path = os.path.join(script_dir, "settings.json")
            
            if os.path.exists(settings_path):
                with open(settings_path, 'r') as f:
                    settings = json.load(f)
                    if "api_key" in settings:
                        self.txtAPIKey.Password = settings["api_key"]
                        logger.info("API key loaded from settings")
                        return True
            return False
        except Exception as ex:
            logger.error("Error loading API key: {}".format(ex))
            return False
    
    def initialize_ai_config(self):
        """Initialize the AI configuration panel"""
        try:
            # Set up AI endpoints
            self.ai_endpoints = {
                "OpenAI o3-mini (OpenRouter)": {
                    "url": "https://openrouter.ai/api/v1/chat/completions",
                    "model": "openai/o3-mini"
                },
                "Claude (OpenRouter)": {
                    "url": "https://openrouter.ai/api/v1/chat/completions",
                    "model": "anthropic/claude-3-opus"
                },
                "Gemini (OpenRouter)": {
                    "url": "https://openrouter.ai/api/v1/chat/completions",
                    "model": "google/gemini-pro"
                },
                "Custom API": {
                    "url": "https://api.openai.com/v1/chat/completions",
                    "model": "gpt-4"
                }
            }
            
            # Load prompts from the prompts folder
            self.load_prompts()
            
            # Load API key from settings
            self.load_api_key()
            
            # Set up event handlers for prompt text changes
            self.txtSystemPrompt.TextChanged += self.prompt_text_changed
            self.txtUserPrompt.TextChanged += self.prompt_text_changed
            
            logger.info("AI configuration initialized")
            return True
        except Exception as ex:
            logger.error("Error initializing AI configuration: {}".format(ex))
            return False
            
    def load_prompts(self):
        """Load prompts from markdown files in the prompts folder"""
        try:
            # Get the script directory
            script_dir = os.path.dirname(__file__)
            prompts_dir = os.path.join(script_dir, "prompts")
            
            # Create the prompts directory if it doesn't exist
            if not os.path.exists(prompts_dir):
                os.makedirs(prompts_dir)
                logger.info("Created prompts directory: {}".format(prompts_dir))
                
                # Create a default prompt file
                self.create_default_prompt_file(prompts_dir)
            
            # Clear the dropdown
            self.cboPromptTemplate.Items.Clear()
            
            # Store prompts in a dictionary
            self.prompts = {}
            
            # Get all markdown files in the prompts directory
            md_files = [f for f in os.listdir(prompts_dir) if f.endswith('.md')]
            
            if not md_files:
                logger.warning("No prompt files found in {}".format(prompts_dir))
                # Create a default prompt file if none exist
                self.create_default_prompt_file(prompts_dir)
                md_files = [f for f in os.listdir(prompts_dir) if f.endswith('.md')]
            
            # Load each prompt file
            for md_file in md_files:
                file_path = os.path.join(prompts_dir, md_file)
                prompt_name, system_prompt, user_prompt = self.parse_prompt_file(file_path)
                
                if prompt_name:
                    # Add to dictionary
                    self.prompts[prompt_name] = {
                        "system_prompt": system_prompt,
                        "user_prompt": user_prompt,
                        "file_path": file_path
                    }
                    
                    # Add to dropdown
                    self.cboPromptTemplate.Items.Add(prompt_name)
            
            # Select the first item
            if self.cboPromptTemplate.Items.Count > 0:
                self.cboPromptTemplate.SelectedIndex = 0
                
            logger.info("Loaded {} prompt templates".format(len(self.prompts)))
            return True
        except Exception as ex:
            logger.error("Error loading prompts: {}".format(ex))
            return False
    
    def create_default_prompt_file(self, prompts_dir):
        """Create a default prompt file"""
        try:
            default_file_path = os.path.join(prompts_dir, "default.md")
            
            with open(default_file_path, 'w') as f:
                f.write("# General Prompt\n\n")
                f.write("## System Prompt\n")
                f.write("You are an expert in building design and construction. ")
                f.write("Your task is to provide clear, concise descriptions for building elements ")
                f.write("based on their properties. Focus on the most important characteristics ")
                f.write("that would be relevant for documentation and specifications.\n\n")
                f.write("## User Prompt\n")
                f.write("Please write a concise description for this element. ")
                f.write("Include key information that would be useful ")
                f.write("in construction documentation. Keep it under 100 words.\n")
                f.write("{Properties}")
            
            logger.info("Created default prompt file: {}".format(default_file_path))
            return True
        except Exception as ex:
            logger.error("Error creating default prompt file: {}".format(ex))
            return False
    
    def parse_prompt_file(self, file_path):
        """Parse a markdown prompt file to extract name, system prompt, and user prompt"""
        try:
            with open(file_path, 'r') as f:
                content = f.read()
            
            # Extract the name (first heading)
            name_match = re.search(r'^# (.+)$', content, re.MULTILINE)
            if not name_match:
                logger.warning("No name found in prompt file: {}".format(file_path))
                return None, None, None
            
            prompt_name = name_match.group(1).strip()
            
            # Extract system prompt (including any intermediate headers)
            system_match = re.search(r'## System Prompt\s*\n([\s\S]*?)(?=\n## |$)', content)
            if not system_match:
                logger.warning("No system prompt found in prompt file: {}".format(file_path))
                return prompt_name, "", ""
            
            system_prompt = system_match.group(1).strip()
            
            # Extract user prompt (including any intermediate headers)
            user_match = re.search(r'## User Prompt\s*\n([\s\S]*?)(?=\n## |$)', content)
            if not user_match:
                logger.warning("No user prompt found in prompt file: {}".format(file_path))
                return prompt_name, system_prompt, ""
            
            user_prompt = user_match.group(1).strip()
            
            logger.debug("Parsed prompt file: {} - Name: {}".format(file_path, prompt_name))
            return prompt_name, system_prompt, user_prompt
        except Exception as ex:
            logger.error("Error parsing prompt file {}: {}".format(file_path, ex))
            return None, None, None
    
    def cboPromptTemplate_SelectionChanged(self, sender, e):
        """Handle prompt template selection change"""
        try:
            selected_prompt = self.cboPromptTemplate.SelectedItem
            if selected_prompt and selected_prompt in self.prompts:
                prompt_data = self.prompts[selected_prompt]
                
                # Update the text boxes
                self.txtSystemPrompt.Text = prompt_data["system_prompt"]
                self.txtUserPrompt.Text = prompt_data["user_prompt"]
                
                logger.info("Selected prompt template: {}".format(selected_prompt))
            
            # Update the preview
            self.prompt_text_changed(None, None)
        except Exception as ex:
            logger.error("Error handling prompt template selection: {}".format(ex))
    
    def populate_schedule_dropdown(self):
        """Populate the schedule dropdown with available schedules"""
        try:
            # Clear existing items
            self.cboScheduleSelector.Items.Clear()
            
            # Get all schedule views in the document
            collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
            schedules = [s for s in collector if not s.IsTemplate]
            
            if not schedules:
                logger.info("No schedules found in the current document.")
                return
            
            # Populate the dropdown with schedule names one by one
            self.schedules = [s.Name for s in schedules]
            for schedule_name in self.schedules:
                self.cboScheduleSelector.Items.Add(schedule_name)
            
            logger.info("Populated schedule dropdown with {} schedules".format(len(self.schedules)))
            
        except Exception as ex:
            logger.error("Error populating schedule dropdown: {}".format(ex))
            forms.alert("Error populating schedule dropdown: {}".format(ex), title="Error")
    
    def select_active_schedule(self):
        """Check if active view is a schedule and select it in the dropdown"""
        try:
            # Get the active document and active view
            active_view = doc.ActiveView
            
            # Check if active view is a schedule
            if active_view and active_view.ViewType == DB.ViewType.Schedule:
                logger.info("Active view is a schedule: {}".format(active_view.Name))
                
                # Find the schedule in the dropdown and select it
                for i in range(self.cboScheduleSelector.Items.Count):
                    item = self.cboScheduleSelector.Items[i]
                    if item.ToString() == active_view.Name:
                        logger.info("Setting selected schedule to: {}".format(active_view.Name))
                        self.cboScheduleSelector.SelectedIndex = i
                        
                        # Load the schedule data
                        self.cboScheduleSelector_SelectionChanged(None, None)
                        return True
                
                logger.info("Active schedule not found in dropdown: {}".format(active_view.Name))
            else:
                logger.info("Active view is not a schedule")
                
            return False
        except Exception as ex:
            logger.error("Error selecting active schedule: {}".format(ex))
            return False
    
    def cboScheduleSelector_SelectionChanged(self, sender, e):
        """Handle schedule selection change"""
        try:
            # Check if we have a valid selection
            if self.cboScheduleSelector.SelectedIndex < 0:
                logger.info("No schedule selected")
                return
                
            # Get the selected schedule name
            selected_schedule_name = self.cboScheduleSelector.SelectedItem.ToString()
            if not selected_schedule_name:
                logger.info("No schedule selected")
                return
            
            logger.info("Schedule selected: {}".format(selected_schedule_name))
            
            # Find the selected schedule in the document
            collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
            schedules = [s for s in collector if not s.IsTemplate]
            self.selected_schedule = next((s for s in schedules if s.Name == selected_schedule_name), None)
            
            if not self.selected_schedule:
                logger.error("Selected schedule not found in the document")
                return
                
            # Update UI
            self.txtSelectedSchedule.Text = self.selected_schedule.Name
            self.txtStatus.Text = "Schedule selected. Loading {} data...".format(
                "Types" if self.is_type_mode else "Instances")
            
            # Check if we need to update the schedule's display units
            self.check_schedule_units()
            
            # Automatically load elements when a schedule is selected
            # Use dispatcher to avoid UI thread issues
            self.Dispatcher.BeginInvoke(Action(self.load_schedule_data))
            
        except Exception as ex:
            logger.error("Error handling schedule selection: {}".format(ex))
            forms.alert("Error handling schedule selection: {}".format(ex), title="Error")
            
    def check_schedule_units(self):
        """Check if the schedule's length fields are displayed in millimeters"""
        try:
            if not self.selected_schedule:
                return
                
            # Get the schedule definition
            schedule_definition = self.selected_schedule.Definition
            field_count = schedule_definition.GetFieldCount()
            
            # Check each field
            for i in range(field_count):
                field_id = schedule_definition.GetFieldId(i)
                field = schedule_definition.GetField(field_id)
                
                # Check if it's a length field
                if self.is_length_field(field):
                    logger.info("Found length field: {}".format(field.GetName()))
                    
                    # Check if we can modify the field's format
                    if hasattr(field, 'GetFormatOptions') and hasattr(field, 'SetFormatOptions'):
                        # Get the current format options
                        format_options = field.GetFormatOptions()
                        
                        # Check if we need to update the display units
                        if hasattr(format_options, 'DisplayUnitType'):
                            current_unit = format_options.DisplayUnitType
                            
                            # If not already in millimeters, log a message
                            # Note: We don't actually change the schedule's units here
                            # as that would modify the Revit document
                            if current_unit != DisplayUnitType.DUT_MILLIMETERS:
                                logger.info("Length field '{}' is not displayed in millimeters (current: {})".format(
                                    field.GetName(), current_unit))
                                
                                # Inform the user
                                self.txtStatus.Text = "Note: Some length fields may not be displayed in millimeters"
                        else:
                            logger.info("Could not determine display unit for field: {}".format(field.GetName()))
                    else:
                        logger.info("Cannot modify format options for field: {}".format(field.GetName()))
        except Exception as ex:
            logger.error("Error checking schedule units: {}".format(ex))
            # Don't show an alert to the user, just log the error
    
    def element_selection_changed(self, sender, e):
        """Handle changes in the Type/Instance radio button selection"""
        try:
            # Update the is_type_mode flag based on the current selection
            self.is_type_mode = self.rbTypes.IsChecked
            
            # Clear any existing data
            self.schedule_data.Clear()
            
            # Update status text
            if self.selected_schedule:
                self.txtStatus.Text = "Selected {} mode. Loading elements...".format(
                    "Types" if self.is_type_mode else "Instances")
                
                # Automatically load elements when selection changes
                # Use dispatcher to avoid UI thread issues
                self.Dispatcher.BeginInvoke(Action(self.load_schedule_data))
            else:
                self.txtStatus.Text = "Please select a schedule first."
            
        except Exception as ex:
            logger.error("Error handling element selection change: {}".format(ex))
            self.txtStatus.Text = "Error changing selection: {}".format(ex)
    
    def btnLoadElements_Click(self, sender, e):
        """Handle the Load Elements button click"""
        try:
            # Ensure we have a selected schedule
            if not self.selected_schedule:
                forms.alert("Please select a schedule first.", title="No Schedule Selected")
                return
            
            # Update the UI to show we're loading
            self.txtStatus.Text = "Loading elements from schedule..."
            
            # Set the type/instance mode based on radio button selection
            self.is_type_mode = self.rbTypes.IsChecked
            
            # Load the schedule data
            row_count = self.load_schedule_data()
            
            # Update the UI based on the result
            if row_count > 0:
                self.txtStatus.Text = "Loaded {} {} from schedule.".format(
                    row_count, "types" if self.is_type_mode else "instances")
                
                # Enable buttons
                self.btnProcessAll.IsEnabled = True
                self.btnSave.IsEnabled = True
                self.btnAcceptAll.IsEnabled = True
            else:
                self.txtStatus.Text = "No elements found in the selected schedule."
                
                # Disable buttons
                self.btnProcessAll.IsEnabled = False
                self.btnSave.IsEnabled = False
                self.btnAcceptAll.IsEnabled = False
        except Exception as ex:
            logger.error("Error loading elements: {}".format(ex))
            self.txtStatus.Text = "Error loading elements"
            forms.alert("Error loading elements: {}".format(ex), title="Error")
    
    def load_schedule_data(self):
        """Extract data from the selected schedule based on Types or Instances selection"""
        try:
            # Ensure we're on the UI thread for collection modifications
            if not self.Dispatcher.CheckAccess():
                self.Dispatcher.Invoke(Action(self.load_schedule_data))
                return 0  # Return 0 to indicate no rows were loaded in this call

            # Clear existing data
            self.schedule_data.Clear()
            logger.info("Schedule data cleared")
            
            # Get schedule data
            if not self.selected_schedule:
                logger.error("No schedule selected")
                return 0
                
            schedule_id = self.selected_schedule.Id
            
            logger.info("Loading {} from schedule with ID: {}".format(
                "types" if self.is_type_mode else "instances", schedule_id.IntegerValue))
            
            # Get the schedule definition to understand fields
            schedule_definition = self.selected_schedule.Definition
            field_count = schedule_definition.GetFieldCount()
            
            # Get field names and IDs
            field_ids = [schedule_definition.GetFieldId(i) for i in range(field_count)]
            field_names = [schedule_definition.GetField(field_id).GetName() for field_id in field_ids]
            
            logger.info("Schedule fields: {}".format(", ".join(field_names)))
            
            # Try to get the TableData from the schedule to access cell values directly
            table_data = None
            section_data = None
            has_table_data = False
            
            try:
                table_data = self.selected_schedule.GetTableData()
                section_data = table_data.GetSectionData(SectionType.Body)
                has_table_data = True
                logger.info("Successfully retrieved schedule table data with {} rows".format(section_data.NumberOfRows))
            except Exception as table_ex:
                logger.warning("Could not get schedule table data: {}".format(table_ex))
            
            # Use FilteredElementCollector to get the elements in the schedule
            collector = None
            
            if self.is_type_mode:
                # Collect instances first, then get their unique types
                collector = FilteredElementCollector(doc, self.selected_schedule.Id)
                collector.WhereElementIsNotElementType()
                # Get unique types from the instances
                type_ids = set(elem.GetTypeId() for elem in collector)
                collector = [doc.GetElement(type_id) for type_id in type_ids if type_id.IntegerValue != -1]
                logger.info("Created collector for unique types from instances in schedule")
            else:
                # Collect instances displayed in the schedule
                collector = FilteredElementCollector(doc, self.selected_schedule.Id)
                collector.WhereElementIsNotElementType()
                logger.info("Created collector for instances in schedule")
            
            # Get all elements from the collector
            elements = list(collector)
            logger.info("Found {} {} in schedule".format(len(elements), "types" if self.is_type_mode else "instances"))
            
            # Create a temporary list to hold the data
            temp_data = []
            
            # Create a mapping of element IDs to row indices if we have table data
            element_to_row_map = {}
            
            if has_table_data and section_data.NumberOfRows > 1:
                try:
                    # Get the number of columns to avoid out of bounds errors
                    num_columns = section_data.NumberOfColumns
                    logger.info("Schedule has {} columns".format(num_columns))
                    
                    # First, try to find an ElementId column
                    element_id_col_idx = -1
                    for col_idx in range(num_columns):
                        try:
                            header_text = self.selected_schedule.GetCellText(SectionType.Header, 0, col_idx)
                            if header_text and ("ElementId" in header_text or "Element ID" in header_text or "ID" == header_text):
                                element_id_col_idx = col_idx
                                logger.info("Found ElementId column at index {}".format(col_idx))
                                break
                        except Exception as header_ex:
                            logger.debug("Error getting header text for column {}: {}".format(col_idx, header_ex))
                    
                    # If we found an ElementId column, use it to map elements to rows
                    if element_id_col_idx >= 0 and element_id_col_idx < num_columns:
                        for row_idx in range(1, section_data.NumberOfRows):
                            try:
                                cell_text = self.selected_schedule.GetCellText(SectionType.Body, row_idx, element_id_col_idx)
                                if cell_text and cell_text.isdigit():
                                    element_id = int(cell_text)
                                    element_to_row_map[element_id] = row_idx
                            except Exception as cell_ex:
                                logger.debug("Error getting cell text at row {}, column {}: {}".format(
                                    row_idx, element_id_col_idx, cell_ex))
                    else:
                        # If no ElementId column, try to match based on other properties
                        # Find Family and Type columns if they exist
                        family_col_idx = -1
                        type_col_idx = -1
                        
                        for col_idx in range(num_columns):
                            try:
                                header_text = self.selected_schedule.GetCellText(SectionType.Header, 0, col_idx)
                                if header_text == "Family":
                                    family_col_idx = col_idx
                                elif header_text == "Type":
                                    type_col_idx = col_idx
                            except Exception as header_ex:
                                logger.debug("Error getting header text for column {}: {}".format(col_idx, header_ex))
                        
                        # If we have both Family and Type columns, use them to match elements
                        if family_col_idx >= 0 and family_col_idx < num_columns and type_col_idx >= 0 and type_col_idx < num_columns:
                            # Create a dictionary of elements by their family and type
                            elements_by_family_type = {}
                            for elem in elements:
                                family = ""
                                type_name = ""
                                
                                if isinstance(elem, FamilySymbol) and hasattr(elem, 'Family') and elem.Family:
                                    family = elem.Family.Name
                                    type_name = elem.Name
                                elif hasattr(elem, 'Symbol') and elem.Symbol:
                                    family = elem.Symbol.FamilyName
                                    type_name = elem.Symbol.Name
                                elif not self.is_element_type(elem) and elem.GetTypeId().IntegerValue != -1:
                                    element_type = doc.GetElement(elem.GetTypeId())
                                    if element_type:
                                        if hasattr(element_type, 'FamilyName'):
                                            family = element_type.FamilyName
                                        if hasattr(element_type, 'Name'):
                                            type_name = element_type.Name
                                elif self.is_element_type(elem):
                                    if hasattr(elem, 'FamilyName'):
                                        family = elem.FamilyName
                                    if hasattr(elem, 'Name'):
                                        type_name = elem.Name
                                
                                key = (family, type_name)
                                if key not in elements_by_family_type:
                                    elements_by_family_type[key] = []
                                elements_by_family_type[key].append(elem)
                            
                            # Now match rows to elements
                            for row_idx in range(1, section_data.NumberOfRows):
                                try:
                                    family = self.selected_schedule.GetCellText(SectionType.Body, row_idx, family_col_idx)
                                    type_name = self.selected_schedule.GetCellText(SectionType.Body, row_idx, type_col_idx)
                                    
                                    key = (family, type_name)
                                    if key in elements_by_family_type:
                                        # If there's only one element with this family and type, it's a match
                                        if len(elements_by_family_type[key]) == 1:
                                            elem = elements_by_family_type[key][0]
                                            element_to_row_map[elem.Id.IntegerValue] = row_idx
                                        # If there are multiple elements, we need additional criteria to match
                                        # This is a simplification - in a real scenario, you might need more sophisticated matching
                                        else:
                                            # Just use the first element as a fallback
                                            elem = elements_by_family_type[key][0]
                                            element_to_row_map[elem.Id.IntegerValue] = row_idx
                                except Exception as row_ex:
                                    logger.debug("Error matching row {}: {}".format(row_idx, row_ex))
                        else:
                            # If we don't have Family and Type columns, just try to match by searching all columns
                            # for element IDs (this is less reliable)
                            for row_idx in range(1, section_data.NumberOfRows):
                                for col_idx in range(num_columns):
                                    try:
                                        cell_text = self.selected_schedule.GetCellText(SectionType.Body, row_idx, col_idx)
                                        if cell_text and cell_text.isdigit():
                                            potential_id = int(cell_text)
                                            # Check if this ID matches any of our elements
                                            for elem in elements:
                                                if elem.Id.IntegerValue == potential_id:
                                                    element_to_row_map[elem.Id.IntegerValue] = row_idx
                                                    break
                                    except Exception as cell_ex:
                                        # Skip this cell if there's an error
                                        pass
                    
                    logger.info("Created mapping for {} elements to schedule rows".format(len(element_to_row_map)))
                except Exception as map_ex:
                    logger.error("Error creating element to row mapping: {}".format(map_ex))
            
            # Process each element in the collector
            successful_rows = 0
            
            for element in elements:
                try:
                    # Skip null elements
                    if element is None:
                        logger.warning("Skipping null element")
                        continue
                    
                    # Log that we're processing this element
                    logger.info("Processing element with ID: {}".format(element.Id.IntegerValue))
                    
                    # Create a dictionary to store element properties
                    row_data = Dictionary[str, object]()
                    
                    # Store the element ID
                    row_data["ElementId"] = element.Id
                                                            
                    # Initialize AI-related fields
                    row_data["AIResponse"] = ""
                    row_data["IsProcessing"] = False
                    row_data["IsAccepted"] = False
                                        
                    # Check if we have a row mapping for this element
                    element_row_idx = element_to_row_map.get(element.Id.IntegerValue, -1)
                    
                    # Get values for each schedule field
                    for i, field_name in enumerate(field_names):
                        try:
                            # First try to get the value from the schedule table if we have a row mapping
                            if has_table_data and element_row_idx >= 0:
                                try:
                                    cell_text = self.selected_schedule.GetCellText(SectionType.Body, element_row_idx, i)
                                    if cell_text:
                                        # Check if this is a length field
                                        field_id = field_ids[i]
                                        field = schedule_definition.GetField(field_id)
                                        
                                        # If it's a length field, ensure it's in millimeters
                                        if self.is_length_field(field):
                                            # The cell_text already has the display units from the schedule
                                            # We'll keep it as is since the schedule should already be configured
                                            # to display in the desired units
                                            row_data[field_name] = cell_text
                                        else:
                                            row_data[field_name] = cell_text
                                        continue  # Skip to next field if we got the value
                                except Exception as cell_ex:
                                    logger.debug("Error getting cell text for field {}: {}".format(field_name, cell_ex))
                            
                            # If we couldn't get from table or don't have table data, try parameters
                            param = element.LookupParameter(field_name)
                            
                            if param and param.HasValue:
                                # If parameter exists on element, use it
                                row_data[field_name] = self.sanitize_text(self.get_parameter_value(param))
                            else:
                                # If not found on element, try to get it from the element type
                                if not self.is_element_type(element):
                                    element_type = doc.GetElement(element.GetTypeId())
                                    if element_type:
                                        type_param = element_type.LookupParameter(field_name)
                                        if type_param and type_param.HasValue:
                                            row_data[field_name] = self.sanitize_text(self.get_parameter_value(type_param))
                                        else:
                                            # Try to get the value from a calculated field or special field
                                            try:
                                                # Get the field definition
                                                field_id = field_ids[i]
                                                field = schedule_definition.GetField(field_id)
                                                
                                                # Check if it's a calculated field
                                                if field.IsCalculatedField:
                                                    # For calculated fields, we can only get the value from the schedule
                                                    # If we couldn't get it from the table earlier, set to empty
                                                    row_data[field_name] = ""
                                                    logger.debug("Field {} is a calculated field but couldn't get value from table".format(field_name))
                                                else:
                                                    # For other fields, try to get the value using the field's parameter ID
                                                    param_id = field.ParameterId
                                                    if param_id.IntegerValue != -1:
                                                        # Try to get the parameter from the element or its type
                                                        param = None
                                                        try:
                                                            # Try element first
                                                            params = element.GetParameters(field_name)
                                                            if params and len(params) > 0:
                                                                param = params[0]
                                                            
                                                            # If not found, try element type
                                                            if (param is None or not param.HasValue) and element_type:
                                                                type_params = element_type.GetParameters(field_name)
                                                                if type_params and len(type_params) > 0:
                                                                    param = type_params[0]
                                                            
                                                            # If found, get the value
                                                            if param and param.HasValue:
                                                                row_data[field_name] = self.sanitize_text(self.get_parameter_value(param))
                                                            else:
                                                                row_data[field_name] = ""
                                                        except Exception as param_ex:
                                                            logger.debug("Error getting parameter for field {}: {}".format(field_name, param_ex))
                                                            row_data[field_name] = ""
                                                    else:
                                                        row_data[field_name] = ""
                                            except Exception as field_ex:
                                                logger.debug("Error getting field definition for {}: {}".format(field_name, field_ex))
                                                row_data[field_name] = ""
                                    else:
                                        row_data[field_name] = ""
                                else:
                                    # If it's already a type and parameter not found, set to empty string
                                    row_data[field_name] = ""
                        except Exception as field_ex:
                            logger.error("Error getting field {}: {}".format(field_name, field_ex))
                            row_data[field_name] = ""
                    
                    # Add to temporary list
                    temp_data.append(row_data)
                    successful_rows += 1
                except Exception as elem_ex:
                    logger.error("Error processing element: {}".format(elem_ex))
            
            # Add all rows to the ObservableCollection
            for row in temp_data:
                self.schedule_data.Add(row)
            
            # Update the UI
            self.update_data_grid_columns()
            
            # Enable buttons if we have data
            self.btnProcessAll.IsEnabled = (self.schedule_data.Count > 0)
            self.btnSave.IsEnabled = (self.schedule_data.Count > 0)
            self.btnAcceptAll.IsEnabled = (self.schedule_data.Count > 0)
            
            # Populate parameter dropdown based on Types/Instances selection
            # Only call if the ComboBox is initialized
            if self.target_parameter_combo:
                self.populate_parameter_dropdown()
            
            # Update prompt previews with the first row of data
            if successful_rows > 0:
                self.prompt_text_changed(None, None)
            
            # Log the results
            logger.info("Loaded {} rows from schedule".format(successful_rows))
            self.txtStatus.Text = "Loaded {} elements from schedule".format(successful_rows)
            
            return successful_rows
        except Exception as ex:
            logger.error("Error loading schedule data: {}".format(ex))
            self.txtStatus.Text = "Error loading schedule data"
            forms.alert("Error loading schedule data: {}".format(ex), title="Error")
            return 0
    
    def populate_parameter_dropdown(self):
        """Populate the parameter dropdown with available parameters"""
        try:
            if not self.target_parameter_combo:
                logger.warning("Parameter ComboBox not initialized")
                return

            # Clear existing items
            self.target_parameter_combo.Items.Clear()

            # Get the first element from the schedule to check available parameters
            if not self.schedule_data or len(self.schedule_data) == 0:
                logger.warning("No schedule data available")
                return

            # Get the first element
            first_element_id = self.schedule_data[0]["ElementId"]
            element = doc.GetElement(first_element_id)
            
            if not element:
                logger.warning("Could not get element from schedule")
                return

            # Get parameters that can be modified
            parameters = [p for p in element.Parameters if not p.IsReadOnly]
            
            # Add parameters to combo box
            for param in parameters:
                # Only add text parameters
                if param.StorageType == StorageType.String:
                    self.target_parameter_combo.Items.Add(param.Definition.Name)

            # Select the first item if available
            if self.target_parameter_combo.Items.Count > 0:
                self.target_parameter_combo.SelectedIndex = 0

        except Exception as ex:
            logger.error("Error populating parameter dropdown: {}".format(ex))

    def get_selected_parameter(self):
        """Get the currently selected parameter name"""
        try:
            if self.target_parameter_combo and self.target_parameter_combo.SelectedItem:
                return self.target_parameter_combo.SelectedItem.ToString()
            return None
        except Exception as ex:
            logger.error("Error getting selected parameter: {}".format(ex))
            return None

    def sanitize_text(self, text):
        """Sanitize text to ensure proper UTF-8 handling"""
        if text is None:
            return ""
        
        try:
            # Handle ElementId objects
            if isinstance(text, ElementId):
                return str(text.IntegerValue)
            
            # Try to handle any encoding issues
            if isinstance(text, str):
                # For Python 2.x, ensure we have a unicode string
                return text.decode('utf-8', errors='replace') if hasattr(text, 'decode') else text
            return str(text)
        except Exception as ex:
            logger.error("Error sanitizing text: {}".format(ex))
            # Return a safe version of the text
            return str(text).encode('ascii', errors='replace').decode('ascii')
    
    def get_parameter_value(self, param):
        """Get the value of a parameter in the appropriate format"""
        if not param.HasValue:
            return None
            
        storage_type = param.StorageType
        
        if storage_type == StorageType.String:
            return param.AsString()
        elif storage_type == StorageType.Integer:
            return param.AsInteger()
        elif storage_type == StorageType.Double:
            # Check if this is a length parameter
            if self.is_length_parameter(param):
                # Get the value in internal units (feet)
                value_in_feet = param.AsDouble()
                try:
                    # Convert from internal units (feet) to millimeters
                    units = UnitTypeId.Millimeters
                    value_in_mm = UnitUtils.ConvertFromInternalUnits(value_in_feet, units)
                    # Format to 2 decimal places
                    return "{:.2f} mm".format(value_in_mm)
                except Exception as ex:
                    logger.error("Error converting length parameter: {}".format(ex))
                    # Fallback to just returning the value
                    return value_in_feet
            else:
                return param.AsDouble()
        elif storage_type == StorageType.ElementId:
            return param.AsElementId().IntegerValue
        else:
            return str(param.AsValueString())
            
    def is_length_parameter(self, param):
        """Check if a parameter is a length parameter"""
        try:
            # Get the parameter definition
            param_def = param.Definition
            
            # Check if it has a unit type
            if hasattr(param_def, 'UnitType'):
                unit_type = param_def.UnitType
                
                # Check if it's a length parameter
                return unit_type == UnitType.UT_Length or unit_type == UnitType.UT_LinearVelocity or unit_type == UnitType.UT_SheetLength
            
            # For older Revit versions or if UnitType is not available
            param_name = param_def.Name.lower()
            length_keywords = ['length', 'width', 'height', 'depth', 'thickness', 'radius', 'diameter', 'offset', 'distance', 'size', 'dimension']
            
            # Check if the parameter name contains any length-related keywords
            for keyword in length_keywords:
                if keyword in param_name:
                    return True
                    
            return False
        except Exception as ex:
            logger.error("Error checking if parameter is length: {}".format(ex))
            return False
    
    def update_data_grid_columns(self):
        """Update DataGrid columns based on the loaded data"""
        try:
            # Only proceed if we have data
            if self.schedule_data.Count == 0:
                logger.info("No data to update DataGrid columns")
                return
            
            # Call setup_data_grid to recreate all columns
            self.setup_data_grid()
            
            # Enable buttons if we have data
            self.btnProcessAll.IsEnabled = (self.schedule_data.Count > 0)
            self.btnSave.IsEnabled = (self.schedule_data.Count > 0)
            self.btnAcceptAll.IsEnabled = (self.schedule_data.Count > 0)
            
            # Populate parameter dropdown based on Types/Instances selection
            # Only call if the ComboBox is initialized
            if self.target_parameter_combo:
                self.populate_parameter_dropdown()
            
            logger.info("DataGrid columns updated successfully")
        except Exception as ex:
            logger.error("Error updating DataGrid columns: {}".format(ex))
            forms.alert("Error updating DataGrid columns: {}".format(ex), title="Error")
    
    def btnProcessRow_Click(self, sender, e):
        """Handle the Process button click for a row"""
        try:
            # Get the button that was clicked
            button = sender
            
            # Get the row data from the button's Tag property
            row_data = button.Tag
            
            # Ensure we have a valid row
            if not row_data:
                logger.error("No row data found for button click")
                return
            
            # Check if we're already processing this row
            if row_data["IsProcessing"]:
                logger.info("Row is already being processed")
                return
            
            # Get the AI configuration
            agent = self.cboAIAgent.Text
            system_prompt = self.txtSystemPrompt.Text
            user_prompt_template = self.txtUserPrompt.Text
            api_key = self.txtAPIKey.Password
            
            # Validate inputs
            if not api_key:
                forms.alert("Please enter an API key in the AI Configuration section.", title="Missing API Key")
                return
            
            if not system_prompt:
                forms.alert("Please enter a system prompt in the AI Configuration section.", title="Missing System Prompt")
                return
            
            if not user_prompt_template:
                forms.alert("Please enter a user prompt template in the AI Configuration section.", title="Missing User Prompt")
                return
            
            # Update UI to show processing
            self.txtStatus.Text = "Processing row..."
            
            # Process the row in a background thread
            def process_in_background():
                try:
                    # Process the row with AI
                    self.process_row_with_ai(row_data, agent, system_prompt, user_prompt_template, api_key)
                    
                    # Update UI on success
                    def update_success_status():
                        self.txtStatus.Text = "Processing completed for row"
                    self.Dispatcher.Invoke(Action(update_success_status))
                    
                except Exception as thread_ex:
                    # Update UI on failure
                    def update_failure_status():
                        self.txtStatus.Text = "Error processing row: {}".format(thread_ex)
                    self.Dispatcher.Invoke(Action(update_failure_status))
                    
                    # Log the error
                    logger.error("Error in background thread: {}".format(thread_ex))
            
            # Start the background thread
            thread = Thread(ThreadStart(process_in_background))
            thread.IsBackground = True
            thread.Start()
            
        except Exception as ex:
            # Update UI on error
            def update_error():
                self.txtStatus.Text = "Error: {}".format(ex)
            self.Dispatcher.Invoke(Action(update_error))
            
            # Log the error
            logger.error("Error processing row: {}".format(ex))
    
    def process_row_with_ai(self, row_data, agent, system_prompt, user_prompt_template, api_key=None):
        """Process a single row with AI"""
        try:
            # Use provided API key or get from textbox
            if not api_key:
                api_key = self.txtAPIKey.Password
            
            if not api_key:
                raise ValueError("API key is required")
                
            # Log row data and agent before processing
            logger.info("Processing row data: {}".format(row_data))
            logger.info("Using AI agent: {}".format(agent))

            # Mark the row as processing
            def mark_processing():
                row_data["IsProcessing"] = True
                row_data["AIResponse"] = "Processing..."
                self.dataGrid.Items.Refresh()
            self.Dispatcher.Invoke(Action(mark_processing))
            
            # Format the user prompt with the row data properties
            user_prompt = self.format_prompt_with_properties(user_prompt_template, row_data)
            
            # Make the API request
            ai_text = self.make_api_request(agent, system_prompt, user_prompt, api_key)
            
            # Update the row with the AI response
            if ai_text:
                self.update_row_with_ai_response(row_data, ai_text)
                return True
            else:
                # Update with error
                def update_error():
                    row_data["AIResponse"] = "Error: Failed to get a response from the AI"
                    row_data["IsProcessing"] = False
                    self.dataGrid.Items.Refresh()
                
                self.Dispatcher.Invoke(Action(update_error))
                return False
                
        except Exception as ex:
            logger.error("Error processing row with AI: {}".format(ex))
            
            # Update with error
            def update_error():
                row_data["AIResponse"] = "Error: {}".format(ex)
                row_data["IsProcessing"] = False
                self.dataGrid.Items.Refresh()
            
            self.Dispatcher.Invoke(Action(update_error))
            return False
        
    def btnProcessAll_Click(self, sender, e):
        """Process all rows with AI"""
        try:
            # Check if we have a schedule selected
            if not self.selected_schedule:
                forms.alert("Please select a schedule first.", title="No Schedule Selected")
                return
            
            # Check if we have an API key
            api_key = self.txtAPIKey.Password.strip()
            if not api_key:
                forms.alert("Please enter an API key.", title="No API Key")
                return
            
            # Update status
            self.txtStatus.Text = "Starting parallel processing of all rows..."
            
            # Get AI configuration
            agent = self.cboAIAgent.SelectedItem.ToString()
            system_prompt = self.txtSystemPrompt.Text
            user_prompt_template = self.txtUserPrompt.Text
            
            # Start processing in a background thread
            thread = Thread(ThreadStart(lambda: self.process_all_rows_parallel(
                agent, system_prompt, user_prompt_template, api_key)))
            thread.IsBackground = True
            thread.Start()
            
        except Exception as ex:
            logger.error("Error processing all rows: {}".format(ex))
            forms.alert("Error processing all rows: {}".format(ex), title="Error")
    
    def btnAcceptAll_Click(self, sender, e):
        """Accept all rows that have AI responses"""
        try:
            # Count how many rows have AI responses
            rows_with_responses = [row for row in self.schedule_data if row["AIResponse"] and not row["AIResponse"].startswith("Error:") and not row["AIResponse"] == "Processing..."]
            
            if not rows_with_responses:
                forms.alert("No AI responses to accept.", title="No Responses")
                return
            
            # Confirm with the user
            result = forms.alert(
                "Accept all {} rows with AI responses?".format(len(rows_with_responses)),
                title="Confirm Accept All",
                ok=False,
                yes=True,
                no=True
            )
            
            if not result:
                return
            
            # Set IsAccepted to True for all rows with responses
            for row in rows_with_responses:
                row["IsAccepted"] = True
            
            # Refresh the DataGrid
            self.dataGrid.Items.Refresh()
            
            # Update status
            self.txtStatus.Text = "Accepted {} rows with AI responses.".format(len(rows_with_responses))
            
        except Exception as ex:
            logger.error("Error accepting all rows: {}".format(ex))
            forms.alert("Error accepting all rows: {}".format(ex), title="Error")
    
    def btnSave_Click(self, sender, e):
        """Save accepted AI responses to Revit"""
        try:
            # Get the selected parameter
            selected_param = self.get_selected_parameter()
            if not selected_param:
                self.txtStatus.Text = "Please select a parameter to save to."
                return

            # Count how many rows are accepted
            accepted_rows = [row for row in self.schedule_data 
                           if "IsAccepted" in row and "AIResponse" in row 
                           and row["IsAccepted"] and row["AIResponse"]]
            
            if not accepted_rows:
                self.txtStatus.Text = "No accepted AI responses to save."
                return
            
            # Update status
            self.txtStatus.Text = "Saving {} responses to Revit...".format(len(accepted_rows))
            
            # Use the external event to update Revit
            self.external_event.Raise()

        except Exception as ex:
            logger.error("Error saving to Revit: {}".format(ex))
            self.txtStatus.Text = "Error saving to Revit: {}".format(ex)
    
    def update_revit_document(self, uiapp):
        """Update the Revit document with accepted AI responses (called by External Event)"""
        try:
            # Get accepted rows
            accepted_rows = [row for row in self.schedule_data 
                           if "IsAccepted" in row and "AIResponse" in row 
                           and row["IsAccepted"] and row["AIResponse"]]
            
            if not accepted_rows:
                return
            
            # Get the selected parameter name using the new method
            selected_param_name = self.get_selected_parameter()
            
            if not selected_param_name:
                logger.error("No parameter selected for saving")
                return
            
            # Log information about what we're trying to update
            logger.info("Updating parameter '{}' for {} elements".format(selected_param_name, len(accepted_rows)))
            
            # Start a transaction
            t = Transaction(doc, "Update Elements with AI Data")
            t.Start()
            
            success_count = 0
            failure_count = 0
            
            try:
                # Update each element
                for row in accepted_rows:
                    try:
                        element = doc.GetElement(row["ElementId"])
                        if element:
                            param = element.LookupParameter(selected_param_name)
                            if param and not param.IsReadOnly:
                                param.Set(row["AIResponse"])
                                success_count += 1
                            else:
                                logger.warning("Parameter not found or read-only: {} for element {}".format(
                                    selected_param_name, row["ElementId"]))
                                failure_count += 1
                    except Exception as ex:
                        logger.error("Error updating element {}: {}".format(row["ElementId"], ex))
                        failure_count += 1
                
                t.Commit()
                
                # Update the UI with results
                self.on_revit_update_completed(success_count, failure_count)
                
            except Exception as ex:
                if t.HasStarted():
                    t.RollBack()
                raise
                
        except Exception as ex:
            logger.error("Error updating Revit document: {}".format(ex))
            if t and t.HasStarted():
                t.RollBack()
            raise
    
    def on_revit_update_completed(self, success_count, failure_count=0):
        """Called when Revit update is completed"""
        # Ensure we're on the UI thread
        if not self.Dispatcher.CheckAccess():
            # If we're not on the UI thread, invoke this method on the UI thread
            def update_ui_with_completion():
                self.on_revit_update_completed(success_count, failure_count)
            self.Dispatcher.Invoke(Action(update_ui_with_completion))
            return
            
        # We're on the UI thread now, update the UI
        if failure_count > 0:
            self.txtStatus.Text = "Updated {} elements in Revit. {} elements failed to update.".format(
                success_count, failure_count)
        else:
            self.txtStatus.Text = "Successfully updated {} elements in Revit.".format(success_count)
        
        # Refresh the view to show changes
        uidoc.RefreshActiveView()
    
    def btnClose_Click(self, sender, e):
        """Close the window"""
        # Save settings before closing
        self.save_settings()
        
        # Close the window
        self.Close()

    def setup_data_grid(self):
        """Set up the DataGrid with columns and event handlers"""
        try:
            # Import required WPF types
            from System.Windows.Controls import DataGridTemplateColumn, DataGridTextColumn, DataGridCheckBoxColumn
            from System.Windows.Controls import DataGridSelectionMode, DataGridSelectionUnit
            from System.Windows.Data import Binding
            from System.Windows import DataTemplate, FrameworkElementFactory
            from System.Windows import Thickness
            
            # Clear existing columns
            self.dataGrid.Columns.Clear()
            
            # Enable multiple selection
            self.dataGrid.SelectionMode = DataGridSelectionMode.Extended
            self.dataGrid.SelectionUnit = DataGridSelectionUnit.FullRow
            
            # Paste binding
            paste_command = System.Windows.Input.RoutedCommand()
            paste_binding = System.Windows.Input.KeyBinding(
                paste_command,
                System.Windows.Input.Key.V,
                System.Windows.Input.ModifierKeys.Control
            )
            paste_binding.Command = System.Windows.Input.ApplicationCommands.Paste
            paste_binding.CommandTarget = self.dataGrid
            self.dataGrid.InputBindings.Add(paste_binding)

            # Add command bindings
            self.dataGrid.CommandBindings.Add(
                System.Windows.Input.CommandBinding(
                    System.Windows.Input.ApplicationCommands.Paste,
                    System.Windows.Input.ExecutedRoutedEventHandler(self.handle_paste)
                )
            )
            
            # Set the ItemsSource directly to the schedule_data collection
            self.dataGrid.ItemsSource = self.schedule_data
            
            # Set selection mode
            self.dataGrid.SelectionMode = DataGridSelectionMode.Extended
            self.dataGrid.SelectionUnit = DataGridSelectionUnit.FullRow
            
            # Get the schedule definition to understand columns
            field_names = []
            if hasattr(self, 'selected_schedule') and self.selected_schedule:
                try:
                    schedule_definition = self.selected_schedule.Definition
                    field_count = schedule_definition.GetFieldCount()
                    
                    # Get field names
                    field_ids = [schedule_definition.GetFieldId(i) for i in range(field_count)]
                    field_names = [schedule_definition.GetField(field_id).GetName() for field_id in field_ids]
                    
                    # Log the field names
                    logger.info("Schedule fields: {}".format(", ".join(field_names)))
                    
                    # Create ElementId column first
                    element_id_column = DataGridTextColumn()
                    element_id_column.Header = "ElementId"
                    element_id_column.Binding = Binding("[ElementId].IntegerValue")
                    element_id_column.IsReadOnly = True
                    element_id_column.Width = System.Windows.Controls.DataGridLength(80)
                    self.dataGrid.Columns.Add(element_id_column)
                    
                    # Add columns for each schedule field
                    for field_name in field_names:
                        column = DataGridTextColumn()
                        column.Header = field_name
                        column.Binding = Binding("[{}]".format(field_name))
                        column.IsReadOnly = True
                        column.Width = System.Windows.Controls.DataGridLength(1, System.Windows.Controls.DataGridLengthUnitType.Star)
                        self.dataGrid.Columns.Add(column)
                except Exception as ex:
                    logger.error("Error getting schedule fields: {}".format(ex))
            else:
                # Create a default ElementId column if no schedule is selected
                element_id_column = DataGridTextColumn()
                element_id_column.Header = "ElementId"
                element_id_column.Binding = Binding("[ElementId].IntegerValue")
                element_id_column.IsReadOnly = True
                element_id_column.Width = System.Windows.Controls.DataGridLength(80)
                self.dataGrid.Columns.Add(element_id_column)
                logger.info("No schedule selected, creating default columns only")
            
            # Add Process button column
            process_column = DataGridTemplateColumn()
            process_column.Header = "Process"
            process_column.Width = System.Windows.Controls.DataGridLength(80)
            
            # Create template for the button
            template = DataTemplate()
            factory = FrameworkElementFactory(System.Windows.Controls.Button)
            factory.SetValue(System.Windows.Controls.Button.ContentProperty, "Process")
            factory.SetValue(System.Windows.Controls.Button.MarginProperty, System.Windows.Thickness(2))
            factory.SetValue(System.Windows.Controls.Button.PaddingProperty, System.Windows.Thickness(5, 2, 5, 2))
            factory.AddHandler(System.Windows.Controls.Button.ClickEvent, System.Windows.RoutedEventHandler(self.btnProcessRow_Click))
            factory.SetBinding(System.Windows.Controls.Button.TagProperty, Binding())
            template.VisualTree = factory
            
            process_column.CellTemplate = template
            self.dataGrid.Columns.Add(process_column)

            # Add AI Response column
            ai_response_column = System.Windows.Controls.DataGridTemplateColumn()
            
            # Create header template with parameter selector
            header_template = DataTemplate()
            header_panel = FrameworkElementFactory(System.Windows.Controls.StackPanel)
            header_panel.SetValue(System.Windows.Controls.StackPanel.OrientationProperty, 
                                System.Windows.Controls.Orientation.Vertical)
            
            # Add the header text
            header_text = FrameworkElementFactory(System.Windows.Controls.TextBlock)
            header_text.SetValue(System.Windows.Controls.TextBlock.TextProperty, "AI Response")
            header_text.SetValue(System.Windows.Controls.TextBlock.MarginProperty, 
                               System.Windows.Thickness(0, 0, 0, 5))
            header_panel.AppendChild(header_text)
                        
            param_combo = FrameworkElementFactory(System.Windows.Controls.ComboBox)
            param_combo.SetValue(System.Windows.Controls.ComboBox.NameProperty, "cboTargetParameter")
            param_combo.SetValue(System.Windows.Controls.ComboBox.MinWidthProperty, 100.0)
            param_combo.SetValue(System.Windows.Controls.ComboBox.MarginProperty,
                               System.Windows.Thickness(0))
            
            # Store reference to combo box for later population
            self.target_parameter_combo = None
            param_combo.AddHandler(
                System.Windows.Controls.ComboBox.LoadedEvent,
                System.Windows.RoutedEventHandler(self.on_parameter_combo_loaded)
            )
            
            header_panel.AppendChild(param_combo)
            header_template.VisualTree = header_panel
            
            ai_response_column.HeaderTemplate = header_template
            
            # Create cell template for the AI Response column
            template = System.Windows.DataTemplate()
            factory = System.Windows.FrameworkElementFactory(System.Windows.Controls.TextBox)
            factory.SetValue(System.Windows.Controls.TextBox.TextProperty, System.Windows.Data.Binding("[AIResponse]"))
            factory.SetValue(System.Windows.Controls.TextBox.StyleProperty, self.dataGrid.Resources["MultilineTextBoxStyle"])
            
            # Add GotFocus event handler to select all text
            def textbox_got_focus(sender, e):
                textbox = sender
                textbox.Dispatcher.BeginInvoke(System.Action(lambda: textbox.SelectAll()))
            
            # Add key event handler for the TextBox
            def textbox_key_down(sender, e):
                if (e.Key == System.Windows.Input.Key.V and 
                    (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) == System.Windows.Input.ModifierKeys.Control):
                    self.handle_paste(sender, e)
                    e.Handled = True
            
            factory.AddHandler(
                System.Windows.Controls.TextBox.GotFocusEvent,
                System.Windows.RoutedEventHandler(textbox_got_focus)
            )
            
            factory.AddHandler(
                System.Windows.Controls.TextBox.KeyDownEvent,
                System.Windows.Input.KeyEventHandler(textbox_key_down)
            )
            
            template.VisualTree = factory
            ai_response_column.CellTemplate = template
            ai_response_column.Width = System.Windows.Controls.DataGridLength(2, System.Windows.Controls.DataGridLengthUnitType.Star)
            self.dataGrid.Columns.Add(ai_response_column)
            
            # Add Accept checkbox column
            accept_column = DataGridCheckBoxColumn()
            accept_column.Header = "Accept"
            accept_column.Binding = Binding("[IsAccepted]")
            accept_column.Width = System.Windows.Controls.DataGridLength(60)
            self.dataGrid.Columns.Add(accept_column)
            
            # Ensure the DataGrid is visible
            self.dataGrid.Visibility = System.Windows.Visibility.Visible
            
            # Force a refresh of the DataGrid
            self.dataGrid.Items.Refresh()
            
            logger.info("DataGrid setup complete with {} columns".format(self.dataGrid.Columns.Count))
            return True
            
        except Exception as ex:
            logger.error("Error setting up DataGrid: {}".format(ex))
            return False

    def create_data_grid_columns(self):
        """Programmatically create the DataGrid columns"""
        # This method is now replaced by setup_data_grid
        # We'll just call setup_data_grid to ensure backward compatibility
        return self.setup_data_grid()

    def format_prompt_with_properties(self, prompt_template, row_data):
        """Format the prompt template with row data properties"""
        try:
            # Start with the template
            formatted_prompt = prompt_template
                        
            # Special handling for ElementId placeholder - do this first as it's critical
            if "{ElementId}" in formatted_prompt:
                try:
                    # Log only the ElementId value, not the entire row_data dictionary
                    if "ElementId" in row_data:
                        element_id_obj = row_data["ElementId"]
                        
                        # Check if it's None
                        if element_id_obj is None:
                            logger.warning("ElementId is None")
                            element_id_value = "None"
                        else:
                            # Try to get the integer value safely
                            try:
                                if hasattr(element_id_obj, "IntegerValue"):
                                    element_id_value = str(element_id_obj.IntegerValue)
                                    logger.info("Processing ElementId: {}".format(element_id_value))
                                else:
                                    # If it's not a Revit ElementId, try to convert it directly
                                    element_id_value = str(element_id_obj)
                                    logger.info("Processing non-Revit ElementId: {}".format(element_id_value))
                            except Exception as id_ex:
                                logger.debug("Error getting ElementId integer value: {}".format(id_ex))
                                # Fallback to string representation
                                try:
                                    element_id_value = str(element_id_obj)
                                    logger.info("Using string representation of ElementId: {}".format(element_id_value))
                                except:
                                    logger.warning("Failed to convert ElementId to string")
                                    element_id_value = "Invalid ID"
                    else:
                        # If ElementId is not in the dictionary, use a placeholder
                        logger.warning("ElementId key not found in row_data")
                        element_id_value = "Unknown ID"
                    
                    # Replace the placeholder
                    formatted_prompt = formatted_prompt.replace("{ElementId}", element_id_value)
                    logger.debug("Replaced {ElementId} with {}".format(element_id_value))
                except Exception as ex:
                    logger.error("Error handling ElementId placeholder: {}".format(ex))
                    # Provide a safe fallback
                    formatted_prompt = formatted_prompt.replace("{ElementId}", "Unknown ID")
            
            # If the template contains {Properties}, format all properties
            if "{Properties}" in formatted_prompt:
                # Format the properties for the prompt in a more structured way
                properties_lines = []
                
                # Add Element ID safely
                try:
                    if "ElementId" in row_data and hasattr(row_data["ElementId"], "IntegerValue"):
                        properties_lines.append("Element ID: {}".format(row_data["ElementId"].IntegerValue))
                    elif "ElementId" in row_data:
                        properties_lines.append("Element ID: {}".format(row_data["ElementId"]))
                    else:
                        properties_lines.append("Element ID: Unknown")
                except Exception as prop_id_ex:
                    logger.debug("Error adding ElementId to properties: {}".format(prop_id_ex))
                    properties_lines.append("Element ID: Error retrieving ID")
                
                # Get the schedule definition to understand fields
                if self.selected_schedule:
                    schedule_definition = self.selected_schedule.Definition
                    field_count = schedule_definition.GetFieldCount()
                    
                    # Get field names
                    field_ids = [schedule_definition.GetFieldId(i) for i in range(field_count)]
                    field_names = [schedule_definition.GetField(field_id).GetName() for field_id in field_ids]
                    
                    # Add all schedule fields
                    properties_lines.append("\nSchedule Fields:")
                    for field_name in field_names:
                        # Check if the field exists in the dictionary and get its value
                        if field_name in row_data:
                            field_value = row_data[field_name]
                            if field_value:  # Only include non-empty fields
                                properties_lines.append("- {}: {}".format(field_name, field_value))
                
                # Join all properties into a single string
                properties_text = "\n".join(properties_lines)
                
                # Replace the {Properties} placeholder with the formatted text
                formatted_prompt = formatted_prompt.replace("{Properties}", properties_text)
            
            # Special handling for ElementId placeholder
            if "{ElementId}" in formatted_prompt and "ElementId" in row_data:
                element_id_value = str(row_data["ElementId"].IntegerValue)
                formatted_prompt = formatted_prompt.replace("{ElementId}", element_id_value)
                logger.debug("Replaced {ElementId} with {}".format(element_id_value))
            
            # Replace individual field placeholders like {Width} with their values
            if self.selected_schedule:
                schedule_definition = self.selected_schedule.Definition
                field_count = schedule_definition.GetFieldCount()
                
                # Get field names
                field_ids = [schedule_definition.GetFieldId(i) for i in range(field_count)]
                field_names = [schedule_definition.GetField(field_id).GetName() for field_id in field_ids]
                
                # Replace each field placeholder with its value
                for field_name in field_names:
                    placeholder = "{" + field_name + "}"
                    if placeholder in formatted_prompt:
                        if field_name in row_data:
                            try:
                                field_value = str(row_data[field_name])
                                formatted_prompt = formatted_prompt.replace(placeholder, field_value)
                            except Exception as field_ex:
                                logger.debug("Error converting field {} to string: {}".format(field_name, field_ex))
                                formatted_prompt = formatted_prompt.replace(placeholder, "Error")
                        else:
                            # If field doesn't exist, replace with empty string
                            formatted_prompt = formatted_prompt.replace(placeholder, "")
                            logger.debug("Field {} not found in row data, replaced with empty string".format(field_name))
            
            return formatted_prompt
        except Exception as ex:
            logger.error("Error formatting prompt: {}".format(ex))
            # Return a safe fallback prompt
            if "{ElementId}" in prompt_template:
                return prompt_template.replace("{ElementId}", "Unknown ID")
            return prompt_template
    
    def make_api_request(self, agent, system_prompt, user_prompt, api_key):
        """Make an API request to the AI service"""
        try:
            # Get endpoint configuration
            endpoint_config = self.ai_endpoints.get(agent, self.ai_endpoints["OpenAI o3-mini (OpenRouter)"])
            url = endpoint_config["url"]
            model = endpoint_config["model"]
            
            # Prepare request payload
            payload = {
                "model": model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 1000
            }
            
            # Log detailed request information
            logger.info("=== API Request Details ===")
            logger.info("URL: {}".format(url))
            logger.info("Model: {}".format(model))
            logger.info("System prompt length: {} chars".format(len(system_prompt)))
            logger.info("User prompt length: {} chars".format(len(user_prompt)))
            logger.info("Temperature: {}".format(payload['temperature']))
            logger.info("Max tokens: {}".format(payload['max_tokens']))
            
            # Check if API key is valid
            if not api_key or len(api_key.strip()) == 0:
                logger.error("API key is empty or invalid")
                return None
                
            # Set up headers
            headers = {
                'Content-Type': 'application/json',
                'Authorization': 'Bearer ' + api_key
            }
            logger.info("Headers set (Authorization header value hidden)")
            
            # Log the start of the request
            logger.info("Making API request...")
            start_time = time.time()
            
            try:
                # Make the API call using our custom HTTP client
                response = self.http_client.post(url, json=payload, headers=headers, timeout=60)
                
                # Calculate request duration
                duration = time.time() - start_time
                logger.info("Request completed in {:.2f} seconds".format(duration))
                
                # Log response details
                logger.info("=== API Response Details ===")
                logger.info("Status code: {}".format(response.status_code))
                logger.info("Response headers: {}".format(response.headers))
                
                # Check if the request was successful
                if response.status_code == 200:
                    # Parse the response
                    try:
                        response_json = response.json()
                        logger.info("Successfully parsed JSON response")
                    except Exception as json_ex:
                        logger.error("Failed to parse JSON response: {}".format(json_ex))
                        logger.error("Raw response: {}".format(response.text[:500]))  # Log first 500 chars
                        return None
                    
                    # Extract the AI response text
                    choices = response_json.get("choices", [])
                    if not choices:
                        logger.error("API response contains empty choices array")
                        logger.error("Full response: {}".format(response_json))
                        return None
                    
                    first_choice = choices[0]
                    logger.info("Successfully extracted first choice from response")
                    
                    # Handle different response formats
                    if "message" in first_choice:
                        # Standard OpenAI format
                        message = first_choice.get("message", {})
                        ai_text = message.get("content", "")
                        logger.info("Found response in OpenAI format (message.content)")
                    elif "content" in first_choice:
                        # Alternative format
                        ai_text = first_choice.get("content", "")
                        logger.info("Found response in alternative format (content)")
                    elif "text" in first_choice:
                        # Another alternative format
                        ai_text = first_choice.get("text", "")
                        logger.info("Found response in alternative format (text)")
                    else:
                        logger.error("API response choice does not contain recognized content field")
                        logger.error("Available fields in choice: {}".format(list(first_choice.keys())))
                        return None
                    
                    if not ai_text:
                        logger.error("API response does not contain message content")
                        logger.error("Full choice object: {}".format(first_choice))
                        return None
                    
                    logger.info("Successfully extracted AI response (length: {} chars)".format(len(ai_text)))
                    return ai_text
                else:
                    # Log detailed error information
                    logger.error("API request failed with status code: {}".format(response.status_code))
                    logger.error("Response headers: {}".format(response.headers))
                    try:
                        error_json = response.json()
                        logger.error("Error response JSON: {}".format(error_json))
                    except:
                        logger.error("Raw error response: {}".format(response.text[:1000]))  # Log first 1000 chars
                    return None
            except Exception as http_ex:
                logger.error("HTTP request failed: {}".format(http_ex))
                logger.error("Request URL: {}".format(url))
                logger.error("Request timeout: 60 seconds")
                return None
                
        except Exception as ex:
            logger.error("Error making API request: {}".format(str(ex)))
            logger.error("Exception type: {}".format(type(ex).__name__))
            logger.error("Exception traceback: {}".format(traceback.format_exc()))
            return None
    
    def update_row_with_ai_response(self, row_data, ai_text):
        """Update a row with the AI response"""
        try:
            # Update the row on the UI thread
            def update_ui():
                # Update the row data
                row_data["AIResponse"] = ai_text
                row_data["IsProcessing"] = False
                
                # Refresh the DataGrid
                self.dataGrid.Items.Refresh()
                
                # Update status
                self.txtStatus.Text = "AI processing completed for {}".format(row_data["ElementId"].IntegerValue)
            
            # Invoke the update on the UI thread
            self.Dispatcher.Invoke(Action(update_ui))
            
            return True
        except Exception as ex:
            logger.error("Error updating row with AI response: {}".format(ex))
            return False

    def process_all_rows_parallel(self, agent, system_prompt, user_prompt_template, api_key):
        """Process all rows in parallel with AI"""
        try:
            # Create a thread-safe counter for progress tracking
            progress_lock = threading.Lock()
            completed_count = [0]  # Use a list for mutable reference
            total_rows = len(self.schedule_data)
            
            # Function to process a single row in a separate thread
            def process_row_thread(row_data, row_index):
                try:
                    # Mark as processing - must be done on UI thread
                    def mark_processing():
                        row_data["IsProcessing"] = True
                        row_data["AIResponse"] = "Processing..."
                        self.dataGrid.Items.Refresh()
                    self.Dispatcher.Invoke(Action(mark_processing))
                    
                    # Process the row with AI
                    success = self.process_row_with_ai(row_data, agent, system_prompt, user_prompt_template, api_key)
                    
                    # Update progress counter
                    with progress_lock:
                        completed_count[0] += 1
                        current_count = completed_count[0]
                    
                    # Update status on UI thread
                    def update_progress():
                        self.txtStatus.Text = "Processed {} of {} rows with AI...".format(current_count, total_rows)
                    self.Dispatcher.Invoke(Action(update_progress))
                    
                except Exception as ex:
                    logger.error("Error processing row {}: {}".format(row_index, ex))
                    # Update row with error on UI thread
                    def update_error():
                        row_data["AIResponse"] = "Error: {}".format(ex)
                        row_data["IsProcessing"] = False
                        self.dataGrid.Items.Refresh()
                    self.Dispatcher.Invoke(Action(update_error))
                    
                    # Still update progress counter for failed rows
                    with progress_lock:
                        completed_count[0] += 1
                        current_count = completed_count[0]
                    
                    # Update status on UI thread
                    def update_progress():
                        self.txtStatus.Text = "Processed {} of {} rows with AI...".format(current_count, total_rows)
                    self.Dispatcher.Invoke(Action(update_progress))
            
            # Create a list to keep track of all threads
            threads = []
            
            # Maximum number of concurrent threads (adjust based on your system capabilities)
            max_concurrent_threads = 5
            
            # Create and start threads for each row, but limit concurrency
            active_threads = 0
            thread_index = 0
            
            while thread_index < total_rows or active_threads > 0:
                # Start new threads if we're under the limit and have more rows to process
                while active_threads < max_concurrent_threads and thread_index < total_rows:
                    row_data = self.schedule_data[thread_index]
                    thread = threading.Thread(
                        target=process_row_thread, 
                        args=(row_data, thread_index)
                    )
                    thread.daemon = True
                    thread.start()
                    threads.append(thread)
                    thread_index += 1
                    active_threads += 1
                
                # Check for completed threads
                for thread in list(threads):
                    if not thread.is_alive():
                        threads.remove(thread)
                        active_threads -= 1
                
                # Small delay to prevent CPU hogging
                time.sleep(0.1)
            
            # Final status update
            def final_status():
                self.txtStatus.Text = "Completed processing all {} rows.".format(total_rows)
            self.Dispatcher.Invoke(Action(final_status))
            
        except Exception as thread_ex:
            logger.error("Error in parallel processing: {}".format(thread_ex))
            
            # Update status on UI thread
            def update_error_status():
                self.txtStatus.Text = "Error processing rows: {}".format(thread_ex)
            self.Dispatcher.Invoke(Action(update_error_status))

    def is_length_field(self, field):
        """Check if a schedule field is a length field"""
        try:
            # Check if the field has a unit type
            if hasattr(field, 'UnitType'):
                unit_type = field.UnitType
                # Check if it's a length field
                return unit_type == UnitType.UT_Length or unit_type == UnitType.UT_LinearVelocity or unit_type == UnitType.UT_SheetLength
            
            # For older Revit versions or if UnitType is not available
            # Try to get the parameter ID and check if it's a length parameter
            param_id = field.ParameterId
            if param_id.IntegerValue != -1:
                # Try to get the parameter definition
                param_element = doc.GetElement(param_id)
                if param_element and hasattr(param_element, 'GetDefinition'):
                    param_def = param_element.GetDefinition()
                    if hasattr(param_def, 'UnitType'):
                        unit_type = param_def.UnitType
                        return unit_type == UnitType.UT_Length or unit_type == UnitType.UT_LinearVelocity or unit_type == UnitType.UT_SheetLength
            
            # If we can't determine from the parameter, check the field name
            field_name = field.GetName().lower()
            length_keywords = ['length', 'width', 'height', 'depth', 'thickness', 'radius', 'diameter', 'offset', 'distance', 'size', 'dimension']
            
            # Check if the field name contains any length-related keywords
            for keyword in length_keywords:
                if keyword in field_name:
                    return True
                    
            return False
        except Exception as ex:
            logger.error("Error checking if field is length: {}".format(ex))
            return False

    def is_element_type(self, element):
        """Check if an element is a type element (not an instance)"""
        try:
            # First check if it's a known type class
            if isinstance(element, FamilySymbol) or isinstance(element, ElementType):
                return True
                
            # Then try other methods
            if hasattr(element, 'IsElementType') and not isinstance(element, FamilySymbol):
                return element.IsElementType
            elif hasattr(element, 'GetTypeId'):
                # If it has a GetTypeId method and the ID is invalid, it's likely a type
                type_id = element.GetTypeId()
                return type_id.IntegerValue == -1
            else:
                # Try to determine by category
                if hasattr(element, 'Category') and element.Category:
                    # Types usually have categories that end with "Types"
                    return element.Category.Name.endswith("Types")
                return False
        except Exception as ex:
            logger.error("Error checking if element is type: {}".format(ex))
            # Default to False if we can't determine
            return False

    def prompt_text_changed(self, sender, e):
        """Update the preview textboxes when the system prompt or user prompt changes"""
        try:
            # Only update if we have data
            if self.schedule_data and len(self.schedule_data) > 0:
                try:
                    # Get the first row of data for the preview
                    first_row = self.schedule_data[0]
                    
                    # Log the first row keys to help with debugging
                    keys = [key for key in first_row.Keys]
                    
                    # Update system prompt preview
                    system_prompt = self.txtSystemPrompt.Text
                    try:
                        self.txtSystemPromptPreview.Text = self.format_prompt_with_properties(system_prompt, first_row)
                    except Exception as sys_ex:
                        logger.error("Error formatting system prompt: {}".format(sys_ex))
                        self.txtSystemPromptPreview.Text = "Error generating preview: {}".format(sys_ex)
                    
                    # Update user prompt preview
                    user_prompt_template = self.txtUserPrompt.Text
                    try:
                        formatted_user_prompt = self.format_prompt_with_properties(user_prompt_template, first_row)
                        self.txtUserPromptPreview.Text = formatted_user_prompt
                    except Exception as user_ex:
                        logger.error("Error formatting user prompt: {}".format(user_ex))
                        self.txtUserPromptPreview.Text = "Error generating preview: {}".format(user_ex)
                except Exception as row_ex:
                    logger.error("Error accessing first row data: {}".format(row_ex))
                    self.txtSystemPromptPreview.Text = "Error accessing row data: {}".format(row_ex)
                    self.txtUserPromptPreview.Text = "Error accessing row data: {}".format(row_ex)
            else:
                # If no data is loaded, show a message
                self.txtSystemPromptPreview.Text = self.txtSystemPrompt.Text
                self.txtUserPromptPreview.Text = "Load schedule data to see a preview with actual values."
        except Exception as ex:
            logger.error("Error updating prompt previews: {}".format(ex))
            self.txtUserPromptPreview.Text = "Error generating preview: {}".format(ex)
            self.txtSystemPromptPreview.Text = "Error generating preview: {}".format(ex)

    def SystemPrompt_ScrollChanged(self, sender, e):
        """Synchronize scrolling between system prompt and its preview"""
        if not self._system_preview_scrolling:
            self._system_prompt_scrolling = True
            scroll_viewer = sender.Template.FindName("PART_ContentHost", sender)
            if scroll_viewer:
                preview_scroll = self.txtSystemPromptPreview.Template.FindName("PART_ContentHost", self.txtSystemPromptPreview)
                if preview_scroll:
                    preview_scroll.ScrollToVerticalOffset(scroll_viewer.VerticalOffset)
            self._system_prompt_scrolling = False

    def SystemPromptPreview_ScrollChanged(self, sender, e):
        """Synchronize scrolling between system preview and its prompt"""
        if not self._system_prompt_scrolling:
            self._system_preview_scrolling = True
            scroll_viewer = sender.Template.FindName("PART_ContentHost", sender)
            if scroll_viewer:
                prompt_scroll = self.txtSystemPrompt.Template.FindName("PART_ContentHost", self.txtSystemPrompt)
                if prompt_scroll:
                    prompt_scroll.ScrollToVerticalOffset(scroll_viewer.VerticalOffset)
            self._system_preview_scrolling = False

    def UserPrompt_ScrollChanged(self, sender, e):
        """Synchronize scrolling between user prompt and its preview"""
        if not self._user_preview_scrolling:
            self._user_prompt_scrolling = True
            scroll_viewer = sender.Template.FindName("PART_ContentHost", sender)
            if scroll_viewer:
                preview_scroll = self.txtUserPromptPreview.Template.FindName("PART_ContentHost", self.txtUserPromptPreview)
                if preview_scroll:
                    preview_scroll.ScrollToVerticalOffset(scroll_viewer.VerticalOffset)
            self._user_prompt_scrolling = False

    def UserPromptPreview_ScrollChanged(self, sender, e):
        """Synchronize scrolling between user preview and its prompt"""
        if not self._user_prompt_scrolling:
            self._user_preview_scrolling = True
            scroll_viewer = sender.Template.FindName("PART_ContentHost", sender)
            if scroll_viewer:
                prompt_scroll = self.txtUserPrompt.Template.FindName("PART_ContentHost", self.txtUserPrompt)
                if prompt_scroll:
                    prompt_scroll.ScrollToVerticalOffset(scroll_viewer.VerticalOffset)
            self._user_preview_scrolling = False

    def update_cell_value(self, row_data, column_name, value):
        """Update a specific cell value in the row data."""
        try:
            if column_name in row_data:
                row_data[column_name] = value
                return True
            return False
        except Exception as ex:
            logger.error("Error updating cell value: {}".format(ex))
            return False

    def handle_paste(self, sender, e):
        """Handle paste operation for cells."""
        try:
            # Get clipboard content
            clipboard_text = System.Windows.Clipboard.GetText()
            if not clipboard_text:
                return

            # If pasting from a TextBox
            if isinstance(sender, System.Windows.Controls.TextBox):
                row_data = sender.DataContext
                if row_data:
                    self.update_cell_value(row_data, "AIResponse", clipboard_text)
                    row_data["IsAccepted"] = True
                return

            # Get selected cells
            selected_cells = self.dataGrid.SelectedCells
            if not selected_cells or len(selected_cells) == 0:
                return

            # For each selected cell
            for cell in selected_cells:
                if cell.Column.Header.ToString() == "AI Response":
                    self.update_cell_value(cell.Item, "AIResponse", clipboard_text)
                    cell.Item["IsAccepted"] = True

            # Refresh the DataGrid
            self.dataGrid.Items.Refresh()
            
        except Exception as ex:
            logger.error("Error handling paste operation: {}".format(ex))

    def on_parameter_combo_loaded(self, sender, e):
        """Handle the ComboBox Loaded event"""
        try:
            # Store the ComboBox reference
            self.target_parameter_combo = sender
            
            # Populate the ComboBox with available parameters
            # Only populate if we have schedule data
            if self.schedule_data and self.schedule_data.Count > 0:
                self.populate_parameter_dropdown()
        except Exception as ex:
            logger.error("Error handling ComboBox Loaded event: {}".format(ex))

# Main entry point
if __name__ == '__main__':
    # Create and show the window
    window = ScheduleAIProcessorWindow()
    window.Show()  # Show the window without parameters 