# -*- coding: utf-8 -*-
"""Core functions for MMI parameter operations."""

from Autodesk.Revit.DB import Transaction, ElementId, FilteredElementCollector
from Autodesk.Revit.DB import ExtensibleStorage, StorageType
from pyrevit import revit, forms, script
import datetime

# Import the MMI schema
from mmi.schema import MMIParameterSchema
from mmi.config import CONFIG_KEYS

# Initialize logger
logger = script.get_logger()

def get_mmi_parameter_name(doc):
    """Get the configured MMI parameter name from extensible storage.
    
    Args:
        doc: The active Revit document
        
    Returns:
        str: The MMI parameter name or 'MMI' as fallback
    """
    try:
        # Look for existing storage with our schema
        data_storage = FilteredElementCollector(doc) \
            .OfClass(ExtensibleStorage.DataStorage) \
            .ToElements()
        
        for ds in data_storage:
            try:
                entity = ds.GetEntity(MMIParameterSchema.schema)
                if entity.IsValid():
                    param_name = entity.Get[str]("mmi_parameter_name")
                    if param_name:
                        return param_name
            except Exception as e:
                logger.debug("Error checking storage entity: {}".format(str(e)))
                continue
        
        return "MMI"  # Default fallback if no stored value
    except Exception as e:
        logger.error("Error getting MMI parameter name: {}".format(str(e)))
        return "MMI"  # Default fallback

def set_mmi_value(doc, elements, value, param_name=None):
    """Set MMI parameter value on the given elements.
    
    Args:
        doc: The active Revit document
        elements: A list of elements or element_ids
        value: The MMI value to set
        param_name: Optional parameter name, will use stored value or fallback if None
        
    Returns:
        tuple: (success_count, failed_elements_ids)
    """
    if not elements:
        return 0, []
    
    # Get parameter name if not provided
    if not param_name:
        param_name = get_mmi_parameter_name(doc)
    
    # Process elements if they're element_ids
    element_list = []
    for elem in elements:
        if isinstance(elem, ElementId):
            element = doc.GetElement(elem)
            if element:
                element_list.append(element)
        else:
            element_list.append(elem)
    
    if not element_list:
        return 0, []
    
    # Set MMI parameter on each element
    success_count = 0
    failed_elements = []
    
    for element in element_list:
        # Try to set the parameter value
        param = element.LookupParameter(param_name)
        if param and not param.IsReadOnly and param.StorageType == StorageType.String:
            param.Set(str(value))
            success_count += 1
        else:
            failed_elements.append(element.Id)
    
    return success_count, failed_elements

def set_selection_mmi_value(doc, value, show_results=False):
    """Set MMI parameter value on the current selection.
    
    Args:
        doc: The active Revit document
        value: The MMI value to set
        show_results: Whether to show result message
        
    Returns:
        bool: True if operation was successful
    """
    # Get the current selection
    selection = revit.get_selection()
    
    if not selection:
        forms.alert('Please select at least one element.', exitscript=True)
        return False
    
    # Start a transaction
    t = Transaction(doc, 'Set MMI Parameter to {}'.format(value))
    t.Start()
    
    try:
        # Set MMI parameter for selected elements
        success_count, failed_elements = set_mmi_value(doc, selection.element_ids, value)
        
        # Show results if requested
        if show_results and success_count > 0:
            forms.alert(
                'Successfully set MMI parameter to {} on {} elements.'.format(
                    value, success_count
                ),
                title='Success'
            )
        
        # Show failure message if any
        if failed_elements:
            print("Could not set MMI parameter on {} elements.".format(len(failed_elements)))
        
        return success_count > 0
        
    except Exception as e:
        print("Error: {}".format(e))
        forms.alert('Error: {}'.format(e), title='Error')
        return False
    finally:
        # Commit the transaction
        t.Commit()

def get_or_create_mmi_storage(doc):
    """Get existing or create new MMI settings storage element."""
    try:
        schema = MMIParameterSchema.schema # Get the schema definition once
        if not schema:
            logger.error("Could not get MMIParameterSchema definition.")
            return None
            
        # Look for existing storage with our schema
        data_storage_elements = FilteredElementCollector(doc)\
            .OfClass(ExtensibleStorage.DataStorage)\
            .ToElements()
        
        for ds in data_storage_elements:
            try:
                entity = ds.GetEntity(schema)
                # Check both validity and matching schema GUID to be safe
                if entity.IsValid() and entity.Schema.GUID == schema.GUID:
                    logger.debug("Found existing MMI settings storage (ElementId: {})".format(ds.Id))
                    return ds
            except Exception as e:
                # Log potential errors during checking but continue searching
                logger.debug("Error checking storage entity (ElementId: {}): {}".format(ds.Id if ds else 'None', str(e)))
                continue
        
        # If not found, create a new one and initialize it
        logger.debug("No existing MMI storage found. Creating and initializing new one...")
        with revit.Transaction("Create MMI Settings Storage", doc):
            new_storage = ExtensibleStorage.DataStorage.Create(doc)
            # Create a default entity from the schema and set it immediately
            initial_entity = MMIParameterSchema.entity # Creates a new Entity(schema)
            new_storage.SetEntity(initial_entity) 
            logger.debug("Created and initialized new MMI settings storage (ElementId: {})".format(new_storage.Id))
            return new_storage
            
    except Exception as e:
        logger.error("Error in get_or_create_mmi_storage: {}".format(str(e)))
        return None

def save_mmi_parameter(doc, parameter_name):
    """Save the MMI parameter name to extensible storage."""
    try:
        # Get or create data storage
        data_storage = get_or_create_mmi_storage(doc)
        
        if not data_storage:
            logger.error("Could not get or create MMI storage")
            return False
            
        # Get current value
        schema = MMIParameterSchema(data_storage)
        current_param = schema.get("mmi_parameter_name")
        
        # Only update if different
        if current_param != parameter_name:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with revit.Transaction("Save MMI Parameter Setting", doc):
                with MMIParameterSchema(data_storage) as entity:
                    entity.set("mmi_parameter_name", parameter_name)
                    entity.set("last_used_date", timestamp)
                    entity.set("is_validated", True)
            
            logger.debug("Saved MMI parameter name: {}".format(parameter_name))
            return True
        
        return True
    except Exception as e:
        logger.error("Failed to save MMI parameter: {}".format(str(e)))
        return False
    
def save_monitor_config(doc, selected_config):
    """Save the MMI monitor configuration to extensible storage."""
    try:
        logger.debug("Entering save_monitor_config with selected_config: {}".format(selected_config))
        data_storage = get_or_create_mmi_storage(doc)
        if not data_storage:
            logger.error("Could not get or create MMI storage")
            return False
            
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        changes_made = False
        
        # Use the BaseSchema context manager directly to handle the transaction
        with MMIParameterSchema(data_storage, update=True) as entity: 
            logger.debug("Opened MMIParameterSchema context manager.")
            for display_name, new_value in selected_config.items():
                schema_key = CONFIG_KEYS.get(display_name)
                if schema_key:
                    # Get current value *before* setting
                    current_value = entity.get(schema_key)
                    logger.debug("Comparing for key '{}': current='{}' (type: {}), new='{}' (type: {})".format(
                        schema_key, current_value, type(current_value), new_value, type(new_value)
                    ))
                    
                    # Explicitly handle potential None values from initial storage reads
                    if current_value is None:
                         current_value = False # Default to False if not set

                    # Ensure comparison is boolean vs boolean
                    if bool(current_value) != bool(new_value):
                        logger.debug("Change detected for '{}'! Setting to {}".format(schema_key, new_value))
                        entity.set(schema_key, bool(new_value)) # Ensure boolean type is set
                        # Use the correct field name from the schema
                        entity.set("last_used_date", timestamp) # Update timestamp on change
                        changes_made = True
                    else:
                        logger.debug("No change needed for '{}'".format(schema_key))
            
            logger.debug("Exiting MMIParameterSchema context manager. Changes made flag: {}".format(changes_made))
        
        # Log message based on whether changes were made
        if changes_made:
             logger.debug("MMI Monitor configuration changes saved.")
        else:
            logger.debug("No changes detected in MMI Monitor configuration.")
        return True
    except Exception as e:
        logger.error("Failed to save MMI monitor config: {}".format(str(e)))
        return False

def load_monitor_config(doc, use_display_names=False):

    """Load the MMI monitor configuration from extensible storage."""
    config = {}
    try:
        data_storage = get_or_create_mmi_storage(doc)
        if not data_storage:
            # Return defaults with appropriate keys based on use_display_names
            if use_display_names:
                return {key: False for key in CONFIG_KEYS}
            else:
                return {schema_key: False for _, schema_key in CONFIG_KEYS.items()}
            
        schema = MMIParameterSchema(data_storage)
        if schema.is_valid:
            for display_name, schema_key in CONFIG_KEYS.items():
                value = schema.get(schema_key) or False  # Default to False if None
                # Store with either display name or schema key based on parameter
                if use_display_names:
                    config[display_name] = value
                else:
                    config[schema_key] = value
                logger.debug("Loaded config for '{}': {}".format(display_name, value))
            return config
        else:
            # Return defaults with appropriate keys based on use_display_names
            if use_display_names:
                return {key: False for key in CONFIG_KEYS}
            else:
                return {schema_key: False for _, schema_key in CONFIG_KEYS.items()}
            
    except Exception as e:
        logger.error("Failed to load MMI monitor config: {}".format(str(e)))
        # Return defaults with appropriate keys based on use_display_names
        if use_display_names:
            return {key: False for key in CONFIG_KEYS}
        else:
            return {schema_key: False for _, schema_key in CONFIG_KEYS.items()}
    