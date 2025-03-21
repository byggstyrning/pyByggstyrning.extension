# -*- coding: utf-8 -*-
__title__ = "Run\nEverything"
__author__ = "Byggstyrning AB"
__doc__ = """Run all saved StreamBIM checklist configurations automatically.

This tool applies all saved mapping configurations to all elements
in the model that have an IfcGUID parameter without any user interaction."""

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

class RunEverythingProcessor:
    """Run Everything processor that runs all checks without UI."""
    
    def __init__(self):
        """Initialize the processor."""
        # Initialize StreamBIM API client
        self.api_client = streambim_api.StreamBIMClient()
        
        # Initialize config list
        self.configs = []
        
        # Load configurations
        self.load_configurations()
        
        # Log status
        logger.info("Found {} configurations to process".format(len(self.configs)))
    
    def load_configurations(self):
        """Load all mapping configurations from storage."""
        # Load configurations from consolidated storage
        logger.info("Loading configurations from storage...")
        loaded_configs = load_configs_with_pickle(revit.doc)
        
        if loaded_configs:
            logger.info("Found {} configurations in storage".format(len(loaded_configs)))
            # Convert to ConfigItem objects
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
                self.configs.append(config)
            
            logger.info("Loaded {} configurations".format(len(self.configs)))
        else:
            logger.info("No configurations found in storage")
    
    def try_automatic_login(self):
        """Attempt to automatically log in using saved tokens."""
        # Load tokens from file first
        self.api_client.load_tokens()
        
        # Check if token exists
        if self.api_client.idToken:
            logger.info("Found saved StreamBIM login...")
            
            # Try to load saved project ID
            saved_project_id = get_saved_project_id(revit.doc)
            if saved_project_id:
                self.api_client.set_current_project(saved_project_id)
                logger.info("Using saved project ID: {}".format(saved_project_id))
            
            return True
        else:
            logger.error("No saved StreamBIM login found. Please log in using the ChecklistImporter first.")
            return False
    
    def run_import_configurations(self):
        """Run import for all configurations."""
        if not self.configs:
            logger.info("No configurations to process. Exiting.")
            return
            
        logger.info("Starting batch import process for {} configurations".format(len(self.configs)))
        
        try:
            # Track total elements processed and updated
            total_processed = 0
            total_updated = 0
            
            # Process each configuration separately
            for i, config in enumerate(self.configs):
                logger.info("==== Processing configuration {}/{}: {} ====".format(
                    i + 1, len(self.configs), config.DisplayName
                ))
                logger.info("Checklist: {} (ID: {})".format(config.ChecklistName, config.checklist_id))
                logger.info("Property: {} -> Parameter: {}".format(config.streambim_property, config.revit_parameter))
                logger.info("Mapping enabled: {}".format(config.mapping_enabled))
                
                # Skip configurations without checklist ID
                if not config.checklist_id:
                    logger.info("Skipping configuration - no checklist ID")
                    config.elements_processed = 0
                    config.elements_updated = 0
                    continue
                
                # Process this configuration in its own transaction
                processed_count, updated_count = self.process_single_configuration(config, i, len(self.configs))
                
                # Update totals
                total_processed += processed_count
                total_updated += updated_count
            
            logger.info("Batch import completed. Processed {} configurations. Updated {}/{} elements.".format(
                len(self.configs), total_updated, total_processed))
            
        except Exception as e:
            logger.error("Error running batch import: {}".format(str(e)))
            # Log detailed stack trace
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))
    
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
                logger.info("Retrieving checklist items for checklist ID: {}".format(config.checklist_id))
                checklist_items = self.api_client.get_checklist_items(config.checklist_id, limit=0)
                if not checklist_items:
                    logger.info("No checklist items found for checklist ID: {}".format(config.checklist_id))
                    config.elements_processed = 0
                    config.elements_updated = 0
                    return (0, 0)
                    
                logger.info("Retrieved {} checklist items for {}".format(len(checklist_items), config.ChecklistName))
                
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
            
            # Optimize by pre-building element lookup dict for the checklist items
            # First extract all the element IDs from the checklist items
            element_ids = set()
            for item in checklist_items:
                element_id = item.get('object')
                if not element_id:
                    element_id = item.get('attributes', {}).get('elementId')
                if element_id:
                    element_ids.add(element_id)
            
            # If no elements found, skip
            if not element_ids:
                logger.info("No element IDs found in checklist items")
                return (0, 0)
                
            logger.info("Found {} unique element IDs in checklist items".format(len(element_ids)))
            
            # Build a lookup dictionary for elements by IfcGUID
            element_lookup = {}
            logger.info("Building element lookup dictionary...")
            
            # Get all elements in the document
            all_elements = FilteredElementCollector(revit.doc).WhereElementIsNotElementType().ToElements()
            
            # Filter for elements with IfcGUID parameter that match our IDs
            count = 0
            for element in all_elements:
                try:
                    # Check for different variations of the parameter name
                    ifc_guid_param = element.LookupParameter("IFCGuid")
                    if not ifc_guid_param:
                        ifc_guid_param = element.LookupParameter("IfcGUID")
                    if not ifc_guid_param:
                        ifc_guid_param = element.LookupParameter("IFC GUID")
                    
                    if ifc_guid_param and ifc_guid_param.HasValue:
                        guid_value = ifc_guid_param.AsString()
                        if guid_value in element_ids:
                            element_lookup[guid_value] = element
                            count += 1
                except:
                    continue
                    
            logger.info("Found {} elements with matching GUIDs".format(count))
            
            # Start a transaction for this configuration
            t = Transaction(revit.doc, "Batch Import: " + config.DisplayName)
            t.Start()
            
            try:
                logger.info("Starting element processing")
                
                # Set an estimated number of elements for the progress tracking
                config.elements_total = len(checklist_items)
                
                # Process each checklist item directly
                for idx, item in enumerate(checklist_items):                                        
                    processed_count += 1
                    
                    try:
                        # Get the element ID from the checklist item
                        element_id = item.get('object')
                        if not element_id:
                            element_id = item.get('attributes', {}).get('elementId')
                        
                        if not element_id:
                            continue
                        
                        # Find the element using our lookup dictionary
                        element = element_lookup.get(element_id)
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
                    
                    # Log progress periodically
                    if idx % 100 == 0:
                        logger.debug("Processed {}/{} items, updated {} so far".format(idx, len(checklist_items), updated_count))
                
                # Commit the transaction
                t.Commit()
                
                # Update the final progress
                config.elements_processed = processed_count
                config.elements_updated = updated_count
                logger.info("Completed configuration {}/{}: {} - Processed: {}, Updated: {}".format(
                    config_index + 1, total_configs, config.DisplayName, processed_count, updated_count))
                
            except Exception as e:
                # Roll back the transaction if there was an error
                if t.HasStarted():
                    t.RollBack()
                logger.error("Error processing configuration: {}".format(str(e)))
                
        except Exception as e:
            logger.error("Error in process_single_configuration: {}".format(str(e)))
        
        return (processed_count, updated_count)

# Main execution
if __name__ == '__main__':
    logger.info("Starting Run Everything script...")
    
    # Create processor and run without UI
    processor = RunEverythingProcessor()
    
    # Check if we have valid login
    if processor.try_automatic_login():
        # Run import process
        processor.run_import_configurations()
    else:
        logger.error("Cannot proceed without StreamBIM login. Please run the ChecklistImporter tool first to log in.")
