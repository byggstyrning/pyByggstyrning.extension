# -*- coding: utf-8 -*-
"""Configuration management for 3D Zone parameter mapping."""

import base64
import pickle
import uuid
from Autodesk.Revit.DB import FilteredElementCollector, ExtensibleStorage, BuiltInCategory
from pyrevit import revit, script
from zone3d.schema import Zone3DConfigSchema

# Initialize logger
logger = script.get_logger()

def serialize_config(config_dict):
    """Convert BuiltInCategory enums to integers for pickling.
    
    Args:
        config_dict: Configuration dictionary
        
    Returns:
        dict: Serialized configuration dictionary
    """
    serialized = config_dict.copy()
    
    # Convert source_categories from BuiltInCategory to int
    if "source_categories" in serialized:
        serialized["source_categories"] = [
            int(cat) if isinstance(cat, BuiltInCategory) else cat
            for cat in serialized["source_categories"]
        ]
    
    # Convert target_filter_categories from BuiltInCategory to int
    if "target_filter_categories" in serialized:
        serialized["target_filter_categories"] = [
            int(cat) if isinstance(cat, BuiltInCategory) else cat
            for cat in serialized["target_filter_categories"]
        ]
    
    return serialized

def deserialize_config(config_dict):
    """Convert integer category values back to BuiltInCategory enums.
    
    Args:
        config_dict: Serialized configuration dictionary
        
    Returns:
        dict: Deserialized configuration dictionary
    """
    deserialized = config_dict.copy()
    
    # Convert source_categories from int to BuiltInCategory
    if "source_categories" in deserialized:
        converted_cats = []
        for cat in deserialized["source_categories"]:
            # If it's already a BuiltInCategory, keep it
            if isinstance(cat, BuiltInCategory):
                converted_cats.append(cat)
            # Otherwise, try to convert from int
            else:
                try:
                    converted_cats.append(BuiltInCategory(int(cat)))
                except:
                    logger.warning("Could not convert category value {} to BuiltInCategory".format(cat))
                    converted_cats.append(cat)
        deserialized["source_categories"] = converted_cats
    
    # Convert target_filter_categories from int to BuiltInCategory
    if "target_filter_categories" in deserialized:
        converted_cats = []
        for cat in deserialized["target_filter_categories"]:
            # If it's already a BuiltInCategory, keep it
            if isinstance(cat, BuiltInCategory):
                converted_cats.append(cat)
            # Otherwise, try to convert from int
            else:
                try:
                    converted_cats.append(BuiltInCategory(int(cat)))
                except:
                    logger.warning("Could not convert category value {} to BuiltInCategory".format(cat))
                    converted_cats.append(cat)
        deserialized["target_filter_categories"] = converted_cats
    
    return deserialized

def get_or_create_storage(doc):
    """Get existing or create new 3D Zone settings storage element."""
    if not doc:
        logger.error("No active document available")
        return None
        
    try:
        logger.info("Searching for 3D Zone settings storage...")
        data_storages = FilteredElementCollector(doc)\
            .OfClass(ExtensibleStorage.DataStorage)\
            .ToElements()
        
        # Look for our storage with the schema
        for ds in data_storages:
            try:
                # Check if this storage has our schema
                entity = ds.GetEntity(Zone3DConfigSchema.schema)
                if entity.IsValid():
                    logger.info("Found existing 3D Zone settings storage")
                    return ds
            except Exception as e:
                logger.info("Error checking storage entity: {}".format(str(e)))
                continue
        
        logger.info("No existing 3D Zone settings storage found, creating new one...")
        # If not found, create a new one
        with revit.Transaction("Create 3D Zone Settings Storage", doc):
            new_storage = ExtensibleStorage.DataStorage.Create(doc)
            logger.info("Created new 3D Zone settings storage")
            return new_storage
            
    except Exception as e:
        logger.error("Error in get_or_create_storage: {}".format(str(e)))
        return None

def load_configs(doc):
    """Load configurations from 3D Zone storage using pickle serialization."""
    if not doc:
        logger.error("No active document available")
        return []
        
    try:
        # Get storage
        storage = get_or_create_storage(doc)
        if not storage:
            logger.info("No storage found, returning empty config list")
            return []
            
        # Load configurations
        schema = Zone3DConfigSchema(storage)
        if not schema.is_valid:
            logger.info("Invalid 3D Zone schema")
            return []
            
        pickled_configs = schema.get("pickled_configs")
        if not pickled_configs:
            logger.info("No configurations found in storage")
            return []
            
        # Decode and unpickle
        try:
            decoded_data = base64.b64decode(pickled_configs)
            raw_configs = pickle.loads(decoded_data)
            
            # Check if configs contain BuiltInCategory objects (old format)
            # If so, we need to serialize them first before deserializing
            needs_serialization = False
            for cfg in raw_configs:
                if isinstance(cfg, dict):
                    for key in ["source_categories", "target_filter_categories"]:
                        if key in cfg:
                            for item in cfg[key]:
                                if isinstance(item, BuiltInCategory):
                                    needs_serialization = True
                                    break
                            if needs_serialization:
                                break
                if needs_serialization:
                    break
            
            # If old format detected, serialize then deserialize
            if needs_serialization:
                logger.info("Detected old format with BuiltInCategory objects, converting...")
                raw_configs = [serialize_config(cfg) for cfg in raw_configs]
            
            # Deserialize BuiltInCategory enums (convert int back to BuiltInCategory)
            configs = [deserialize_config(cfg) for cfg in raw_configs]
            logger.info("Successfully loaded {} configurations".format(len(configs)))
            
            # Pretty print configs for debugging
            for i, cfg in enumerate(configs):
                logger.info("Configuration {}:".format(i + 1))
                logger.info("  ID: {}".format(cfg.get("id", "N/A")))
                logger.info("  Name: {}".format(cfg.get("name", "N/A")))
                logger.info("  Order: {}".format(cfg.get("order", "N/A")))
                logger.info("  Enabled: {}".format(cfg.get("enabled", False)))
                source_cats = cfg.get("source_categories", [])
                logger.info("  Source Categories: {}".format([str(cat) for cat in source_cats]))
                logger.info("  Source Params: {}".format(cfg.get("source_params", [])))
                logger.info("  Target Params: {}".format(cfg.get("target_params", [])))
                target_filter = cfg.get("target_filter_categories", [])
                if target_filter:
                    logger.info("  Target Filter Categories: {}".format([str(cat) for cat in target_filter]))
                logger.info("")
            
            return configs
        except Exception as e:
            error_msg = str(e)
            logger.error("Error unpickling configurations: {}".format(error_msg))
            
            # Check if it's the BuiltInCategory serialization error
            if "BuiltInCategory" in error_msg or "unknown serialization format" in error_msg.lower():
                logger.info("Detected corrupted configuration data with BuiltInCategory objects.")
                logger.info("Attempting to clear corrupted configuration data...")
                try:
                    with revit.Transaction("Clear Corrupted 3D Zone Configurations", doc):
                        with Zone3DConfigSchema(storage) as entity:
                            entity.set("pickled_configs", "")
                    logger.info("Cleared corrupted configuration data. Please recreate configurations.")
                except Exception as clear_error:
                    logger.error("Failed to clear corrupted data: {}".format(str(clear_error)))
            return []
    except Exception as e:
        logger.error("Error loading configurations: {}".format(str(e)))
        return []

def save_configs(doc, config_list):
    """Save configurations to 3D Zone storage using pickle serialization.
    
    Args:
        doc: The Revit document
        config_list: List of configuration dictionaries
        
    Returns:
        bool: True if successful, False otherwise
    """
    if not doc:
        logger.error("No active document available")
        return False
        
    try:
        # Get or create data storage
        storage = get_or_create_storage(doc)
        if not storage:
            logger.error("Failed to get or create 3D Zone storage")
            return False
        
        # Validate configs have required fields
        for config in config_list:
            required_fields = ["id", "name", "order", "enabled", "source_categories", 
                            "source_params", "target_params"]
            for field in required_fields:
                if field not in config:
                    logger.error("Configuration missing required field: {}".format(field))
                    return False
        
        # Serialize BuiltInCategory enums before pickling
        serialized_configs = [serialize_config(cfg) for cfg in config_list]
        
        # Pickle and encode the data
        try:
            pickled_data = pickle.dumps(serialized_configs)
            encoded_data = base64.b64encode(pickled_data)
            
            # Save to storage
            with revit.Transaction("Save 3D Zone Configurations", doc):
                with Zone3DConfigSchema(storage) as entity:
                    # Save configurations
                    entity.set("pickled_configs", encoded_data)
                    
            logger.info("Saved {} configurations to 3D Zone storage".format(len(config_list)))
            return True
        except Exception as e:
            logger.error("Error pickling configurations: {}".format(str(e)))
            return False
    except Exception as e:
        logger.error("Error saving configurations: {}".format(str(e)))
        return False

def get_next_order(doc):
    """Get the next order number for a new configuration.
    
    Args:
        doc: The Revit document
        
    Returns:
        int: Next order number (1 if no configs exist)
    """
    configs = load_configs(doc)
    if not configs:
        return 1
    
    # Convert generator to list for IronPython 2.7 compatibility (max() doesn't support default keyword)
    orders = [config.get("order", 0) for config in configs]
    max_order = max(orders) if orders else 0
    return max_order + 1

def generate_config_id():
    """Generate a unique ID for a new configuration.
    
    Returns:
        str: Unique ID string
    """
    return str(uuid.uuid4())

def get_config_by_id(doc, config_id):
    """Get a configuration by its ID.
    
    Args:
        doc: The Revit document
        config_id: The configuration ID
        
    Returns:
        dict: Configuration dictionary or None if not found
    """
    configs = load_configs(doc)
    for config in configs:
        if config.get("id") == config_id:
            return config
    return None

def delete_config(doc, config_id):
    """Delete a configuration by its ID.
    
    Args:
        doc: The Revit document
        config_id: The configuration ID to delete
        
    Returns:
        bool: True if successful, False otherwise
    """
    configs = load_configs(doc)
    original_count = len(configs)
    configs = [c for c in configs if c.get("id") != config_id]
    
    if len(configs) == original_count:
        logger.warning("Configuration with ID {} not found".format(config_id))
        return False
    
    return save_configs(doc, configs)

def get_enabled_configs(doc):
    """Get all enabled configurations sorted by order.
    
    Args:
        doc: The Revit document
        
    Returns:
        list: List of enabled configurations sorted by order
    """
    configs = load_configs(doc)
    enabled = [c for c in configs if c.get("enabled", False)]
    return sorted(enabled, key=lambda x: x.get("order", 999))

