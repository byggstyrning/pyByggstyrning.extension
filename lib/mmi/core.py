# -*- coding: utf-8 -*-
"""Core functions for MMI parameter operations."""

from Autodesk.Revit.DB import Transaction, ElementId, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.DB import ExtensibleStorage, StorageType
from pyrevit import revit, forms, script
import datetime
import System

# Import the MMI schema
from mmi.schema import MMIParameterSchema
from mmi.config import CONFIG_KEYS, MONITOR_LIMIT_DEFAULTS
from mmi.changed_date import format_changed_date, should_stamp_changed_date
from mmi.threshold import normalize_limit
try:
    from revit.compat import get_element_id_value
except ImportError:
    def get_element_id_value(item):
        if hasattr(item, 'Value'):
            return item.Value
        return item.IntegerValue

# Initialize logger
logger = script.get_logger()

# Schemas that predate per-warning limits. Newest previous version first.
PREVIOUS_SCHEMA_GUIDS = (
    "8844cb2d-4234-4bf0-8361-b3da4d64234e",
    "8844cb2d-4234-4bf0-8361-b3da4d64234d",
    "8844cb2d-4234-4bf0-8361-b3da4d64234c",
)

_LEGACY_STRING_FIELDS = (
    "mmi_parameter_name",
    "last_used_date",
    "pm_id",
    "default_mmi",
    "changed_date_parameter",
)

_LEGACY_BOOL_FIELDS = (
    "is_validated",
    "validate_mmi",
    "pin_elements",
    "warn_on_move",
    "check_mmi_after_sync",
    "default_on_new_instances",
    "warn_on_type_change",
    "warn_on_instance_params",
    "write_changed_date",
)

_LEGACY_INT_FIELDS = (
    "pin_mmi_limit",
    "move_mmi_limit",
    "type_instance_limit",
    "instance_param_limit",
)

# Categories that should not have MMI values set
# Using integer values directly since some BuiltInCategory enum values may not exist
EXCLUDED_CATEGORIES = [
    int(BuiltInCategory.OST_CurtainGrids),  # Curtain Wall Grids (2034483)
    2451555,                                 # Wall Sweeps (no BuiltInCategory enum available)
]

def _find_current_mmi_storage(doc):
    """Return the current-schema storage element, or None. Does not create one."""
    try:
        schema = MMIParameterSchema.schema
        if not schema:
            return None
        data_storage = FilteredElementCollector(doc) \
            .OfClass(ExtensibleStorage.DataStorage) \
            .ToElements()
        for ds in data_storage:
            try:
                entity = ds.GetEntity(schema)
                if entity.IsValid():
                    return ds
            except Exception as e:
                logger.debug("Error checking storage entity: {}".format(str(e)))
                continue
        return None
    except Exception as e:
        logger.error("Error finding MMI storage: {}".format(str(e)))
        return None


def get_mmi_parameter_name(doc):
    """Get the configured MMI parameter name from extensible storage.
    
    Args:
        doc: The active Revit document
        
    Returns:
        str: The MMI parameter name or 'MMI' as fallback
    """
    try:
        ds = _find_current_mmi_storage(doc)
        if ds is not None:
            entity = ds.GetEntity(MMIParameterSchema.schema)
            param_name = entity.Get[str]("mmi_parameter_name")
            if param_name:
                return param_name
        return "MMI"  # Default fallback if no stored value
    except Exception as e:
        logger.error("Error getting MMI parameter name: {}".format(str(e)))
        return "MMI"  # Default fallback


def read_changed_date_setting(doc):
    """Read the optional changed-date switch and text parameter.

    Does not create storage. Missing storage or fields mean the feature is off.

    Returns:
        tuple: (enabled, parameter_name)
    """
    try:
        ds = _find_current_mmi_storage(doc)
        if ds is None:
            return False, ""
        schema = MMIParameterSchema(ds)
        if not schema.is_valid:
            return False, ""
        enabled = bool(schema.get("write_changed_date"))
        raw = schema.get("changed_date_parameter")
        parameter_name = "" if raw is None else str(raw).strip()
        return enabled, parameter_name
    except Exception as e:
        logger.debug("Could not read changed date setting: {}".format(e))
        return False, ""


def get_changed_date_parameter_name(doc):
    """Return the stored changed-date text parameter name, or \"\"."""
    _enabled, parameter_name = read_changed_date_setting(doc)
    return parameter_name


def _button_changed_date(doc):
    """Parameter name and YYYY-MM-dd to stamp, or (None, None)."""
    enabled, date_param = read_changed_date_setting(doc)
    mmi_param = get_mmi_parameter_name(doc)
    if not should_stamp_changed_date(enabled, date_param, mmi_param):
        return None, None
    return date_param, format_changed_date(datetime.date.today())

def _try_set_text_parameter(element, param_name, value):
    """Set a writable text parameter. Returns True when the value is stored."""
    if not param_name or value is None:
        return False
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly and param.StorageType == StorageType.String:
        param.Set(str(value))
        return True
    return False


def set_mmi_value(doc, elements, value, param_name=None, extra_param_name=None, extra_value=None):
    """Set MMI parameter value on the given elements.
    
    Args:
        doc: The active Revit document
        elements: A list of elements or element_ids
        value: The MMI value to set
        param_name: Optional parameter name, will use stored value or fallback if None
        extra_param_name: Optional text parameter written only after MMI is set
        extra_value: Value written to extra_param_name
        
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
        try:
            # Skip excluded categories
            if element.Category and element.Category.Id:
                category_id = get_element_id_value(element.Category.Id)
                if category_id in EXCLUDED_CATEGORIES:
                    # Silently skip excluded categories (don't add to failed_elements)
                    logger.debug("Skipping excluded category: {} ({})".format(
                        element.Category.Name, category_id))
                    continue
            
            if _try_set_text_parameter(element, param_name, value):
                success_count += 1
                if extra_param_name and extra_value is not None:
                    if not _try_set_text_parameter(element, extra_param_name, extra_value):
                        logger.debug(
                            "Could not set changed date parameter '{}' on element {}".format(
                                extra_param_name, element.Id
                            )
                        )
            else:
                failed_elements.append(element.Id)
        except Exception as ex:
            # Catch any Revit API exceptions during parameter setting
            logger.debug("Error setting MMI parameter on element {}: {}".format(element.Id, ex))
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
    
    # Read before the transaction so storage lookup does not nest transactions.
    extra_name, extra_value = _button_changed_date(doc)
    logger.info(
        "MMI button write: value={0} selection={1} changed_date_param={2} date={3}".format(
            value,
            len(selection.element_ids) if selection.element_ids else 0,
            extra_name or "",
            extra_value or "",
        )
    )

    # Start a transaction
    t = Transaction(doc, 'Set MMI Parameter to {}'.format(value))
    t.Start()
    
    try:
        # Set MMI parameter for selected elements. The date is written in the
        # same pass, and only on elements whose MMI value was set.
        success_count, failed_elements = set_mmi_value(
            doc,
            selection.element_ids,
            value,
            extra_param_name=extra_name,
            extra_value=extra_value,
        )
        
        # Show results if requested
        if show_results and success_count > 0:
            forms.alert(
                'Successfully set MMI parameter to {} on {} elements.'.format(
                    value, success_count
                ),
                title='Success'
            )
        
        # Missing or read-only parameters are expected on some categories.
        # Keep this in the debug log so it does not open a user warning.
        if failed_elements:
            failed_ids = []
            for element_id in failed_elements:
                failed_ids.append(str(get_element_id_value(element_id)))
            logger.debug(
                "Could not set MMI parameter on {} element(s): {}".format(
                    len(failed_elements), ", ".join(failed_ids)
                )
            )
        
        return success_count > 0
        
    except Exception as e:
        logger.error("Error setting MMI on selection: {0}".format(e))
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
        
        # Check for current schema version first
        for ds in data_storage_elements:
            try:
                entity = ds.GetEntity(schema)
                # Check both validity and matching schema GUID to be safe
                if entity.IsValid() and entity.Schema.GUID == schema.GUID:
                    logger.debug("Found existing MMI settings storage with current schema (ElementId: {})".format(ds.Id))
                    return ds
            except Exception as e:
                # Log potential errors during checking but continue searching
                logger.debug("Error checking storage entity (ElementId: {}): {}".format(ds.Id if ds else 'None', str(e)))
                continue
        
        # Check for old schema versions and migrate if found
        for guid_text in PREVIOUS_SCHEMA_GUIDS:
            old_schema_guid = System.Guid(guid_text)
            for ds in data_storage_elements:
                try:
                    old_schemas = ds.GetEntitySchemaGuids()
                    if old_schema_guid in old_schemas:
                        logger.debug("Found old schema storage, migrating to new version (ElementId: {})".format(ds.Id))
                        return migrate_mmi_storage(doc, ds, old_schema_guid)
                except Exception as e:
                    logger.debug("Error checking for old schema (ElementId: {}): {}".format(ds.Id if ds else 'None', str(e)))
                    continue
        
        # If not found, create a new one and initialize it
        logger.debug("No existing MMI storage found. Creating and initializing new one...")
        with revit.Transaction("Create MMI Settings Storage", doc):
            new_storage = ExtensibleStorage.DataStorage.Create(doc)
            # Create a default entity from the schema and set it immediately
            initial_entity = MMIParameterSchema.entity # Creates a new Entity(schema)
            new_storage.SetEntity(initial_entity) 
            logger.debug("Created and initialized new MMI settings storage (ElementId: {})".format(new_storage.Id))
        _seed_monitor_limit_defaults(new_storage)
        return new_storage
            
    except Exception as e:
        logger.error("Error in get_or_create_mmi_storage: {}".format(str(e)))
        return None

def _read_legacy_string(old_entity, name):
    try:
        value = old_entity.Get[str](name)
        if value is None:
            return ""
        return value
    except Exception:
        return ""


def _read_legacy_bool(old_entity, name):
    try:
        return bool(old_entity.Get[bool](name))
    except Exception:
        return False


def _read_legacy_int(old_entity, name):
    try:
        return int(old_entity.Get[System.Int32](name))
    except Exception:
        return None


def _extract_legacy_monitor_data(old_entity):
    old_data = {}
    for name in _LEGACY_STRING_FIELDS:
        old_data[name] = _read_legacy_string(old_entity, name)
    for name in _LEGACY_BOOL_FIELDS:
        old_data[name] = _read_legacy_bool(old_entity, name)
    for name in _LEGACY_INT_FIELDS:
        old_data[name] = _read_legacy_int(old_entity, name)
    return old_data


def _apply_limit_defaults(entity):
    """Write limit defaults and leave newer optional toggles off."""
    entity.set("schema_version", "1.5")
    for key, default in MONITOR_LIMIT_DEFAULTS.items():
        entity.set(key, int(default))
    entity.set("warn_on_type_change", False)
    entity.set("warn_on_instance_params", False)
    entity.set("write_changed_date", False)
    entity.set("changed_date_parameter", "")


def _seed_monitor_limit_defaults(data_storage):
    """Fill a newly created storage element with limit defaults."""
    try:
        with MMIParameterSchema(data_storage, update=True) as entity:
            _apply_limit_defaults(entity)
    except Exception as e:
        logger.warning("Could not seed monitor limit defaults: {}".format(e))


def migrate_mmi_storage(doc, old_storage, old_schema_guid):
    """Migrate old MMI storage to the current schema version."""
    try:
        old_data = {}
        try:
            old_schemas = old_storage.GetEntitySchemaGuids()
            if old_schema_guid in old_schemas:
                old_entity = old_storage.GetEntity(old_schema_guid)
                if old_entity.IsValid():
                    old_data = _extract_legacy_monitor_data(old_entity)
                    logger.debug("Extracted old data: {}".format(old_data))
        except Exception as e:
            logger.warning("Could not extract old data, using defaults: {}".format(e))
        
        # Create new storage with migrated data
        with revit.Transaction("Migrate MMI Settings Storage", doc):
            # Delete old storage
            doc.Delete(old_storage.Id)
            
            # Create new storage
            new_storage = ExtensibleStorage.DataStorage.Create(doc)
            
            # Create entity with new schema and migrated data
            with MMIParameterSchema(new_storage, update=True) as entity:
                entity.set("mmi_parameter_name", old_data.get("mmi_parameter_name", ""))
                entity.set("last_used_date", old_data.get("last_used_date", ""))
                entity.set("pm_id", old_data.get("pm_id", ""))
                entity.set("default_mmi", old_data.get("default_mmi", ""))
                entity.set("is_validated", old_data.get("is_validated", False))
                entity.set("validate_mmi", old_data.get("validate_mmi", False))
                entity.set("pin_elements", old_data.get("pin_elements", False))
                entity.set("warn_on_move", old_data.get("warn_on_move", False))
                entity.set("check_mmi_after_sync", old_data.get("check_mmi_after_sync", False))
                entity.set("default_on_new_instances", old_data.get("default_on_new_instances", False))
                _apply_limit_defaults(entity)
                entity.set("warn_on_type_change", bool(old_data.get("warn_on_type_change", False)))
                entity.set("warn_on_instance_params", bool(old_data.get("warn_on_instance_params", False)))
                entity.set("write_changed_date", bool(old_data.get("write_changed_date", False)))
                date_parameter = old_data.get("changed_date_parameter", "")
                if not date_parameter:
                    date_parameter = ""
                entity.set("changed_date_parameter", str(date_parameter))
                for key, default in MONITOR_LIMIT_DEFAULTS.items():
                    raw_limit = old_data.get(key, None)
                    if raw_limit is None:
                        continue
                    entity.set(key, normalize_limit(raw_limit, default))
                
            logger.debug("Successfully migrated MMI storage to new schema (ElementId: {})".format(new_storage.Id))
            return new_storage
            
    except Exception as e:
        logger.error("Error migrating MMI storage: {}".format(str(e)))
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


def get_default_mmi(doc):
    """Read project default MMI for new instances from extensible storage.

    Returns:
        str: Stored value or \"\" if unset / invalid.
    """
    try:
        data_storage = get_or_create_mmi_storage(doc)
        if not data_storage:
            return ""
        schema = MMIParameterSchema(data_storage)
        if not schema.is_valid:
            return ""
        try:
            raw = schema.get("default_mmi")
        except Exception:
            return ""
        if raw is None:
            return ""
        s = str(raw).strip()
        return s
    except Exception as e:
        logger.error("Error reading default_mmi: {}".format(str(e)))
        return ""


def save_default_mmi(doc, value):
    """Persist default MMI for new instances (empty string disables).

    Returns:
        bool: True on success.
    """
    try:
        data_storage = get_or_create_mmi_storage(doc)
        if not data_storage:
            logger.error("Could not get or create MMI storage for default_mmi")
            return False
        normalized = "" if value is None else str(value).strip()
        schema = MMIParameterSchema(data_storage)
        current = ""
        if schema.is_valid:
            try:
                cur = schema.get("default_mmi")
                if cur is not None:
                    current = str(cur).strip()
            except Exception:
                current = ""
        if current == normalized:
            return True
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with revit.Transaction("Save MMI default on new instances", doc):
            with MMIParameterSchema(data_storage) as entity:
                entity.set("default_mmi", normalized)
                entity.set("last_used_date", timestamp)
        logger.debug("Saved default_mmi: {}".format(repr(normalized)))
        return True
    except Exception as e:
        logger.error("Failed to save default_mmi: {}".format(str(e)))
        return False


def save_changed_date_parameter(doc, parameter_name):
    """Persist the text parameter that receives the changed date.

    Returns:
        bool: True on success.
    """
    try:
        data_storage = get_or_create_mmi_storage(doc)
        if not data_storage:
            logger.error("Could not get or create MMI storage for changed date parameter")
            return False
        normalized = "" if parameter_name is None else str(parameter_name).strip()
        _enabled, current = read_changed_date_setting(doc)
        if current == normalized:
            return True
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with revit.Transaction("Save MMI changed date parameter", doc):
            with MMIParameterSchema(data_storage) as entity:
                entity.set("changed_date_parameter", normalized)
                entity.set("last_used_date", timestamp)
        logger.debug("Saved changed_date_parameter: {}".format(repr(normalized)))
        return True
    except Exception as e:
        logger.error("Failed to save changed date parameter: {}".format(str(e)))
        return False


def _default_limit_config():
    limits = {}
    for key, default in MONITOR_LIMIT_DEFAULTS.items():
        limits[key] = int(default)
    return limits


def _read_stored_limit(schema, key, default):
    if schema is None or not getattr(schema, "is_valid", False):
        return int(default)
    try:
        raw = schema.get(key)
    except Exception:
        return int(default)
    return normalize_limit(raw, default)


def load_monitor_limits(doc):
    """Read the four per-warning integer limits.

    Returns:
        dict: Schema key to integer limit. Missing or invalid values use defaults.
    """
    limits = _default_limit_config()
    try:
        data_storage = get_or_create_mmi_storage(doc)
        if not data_storage:
            return limits
        schema = MMIParameterSchema(data_storage)
        for key, default in MONITOR_LIMIT_DEFAULTS.items():
            limits[key] = _read_stored_limit(schema, key, default)
        return limits
    except Exception as e:
        logger.error("Failed to load MMI monitor limits: {}".format(str(e)))
        return limits


def save_monitor_limits(doc, limits):
    """Persist per-warning integer limits. Invalid values are stored as defaults.

    Returns:
        bool: True on success.
    """
    try:
        data_storage = get_or_create_mmi_storage(doc)
        if not data_storage:
            logger.error("Could not get or create MMI storage for monitor limits")
            return False
        incoming = limits or {}
        normalized = {}
        for key, default in MONITOR_LIMIT_DEFAULTS.items():
            normalized[key] = normalize_limit(incoming.get(key), default)

        schema = MMIParameterSchema(data_storage)
        changes_made = False
        for key, new_value in normalized.items():
            current = _read_stored_limit(schema, key, MONITOR_LIMIT_DEFAULTS[key])
            if current != new_value:
                changes_made = True
                break
        if not changes_made:
            return True

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with revit.Transaction("Save MMI monitor limits", doc):
            with MMIParameterSchema(data_storage) as entity:
                for key, new_value in normalized.items():
                    entity.set(key, int(new_value))
                entity.set("last_used_date", timestamp)
        logger.debug("Saved MMI monitor limits: {}".format(normalized))
        return True
    except Exception as e:
        logger.error("Failed to save MMI monitor limits: {}".format(str(e)))
        return False


def _empty_monitor_config(use_display_names):
    if use_display_names:
        config = {key: False for key in CONFIG_KEYS}
    else:
        config = {schema_key: False for _, schema_key in CONFIG_KEYS.items()}
    config.update(_default_limit_config())
    return config


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
        
        # Build a reverse mapping (schema_key -> schema_key) for when config uses schema keys directly
        valid_schema_keys = set(CONFIG_KEYS.values())
        
        # Use the BaseSchema context manager directly to handle the transaction
        with MMIParameterSchema(data_storage, update=True) as entity: 
            logger.debug("Opened MMIParameterSchema context manager.")
            for key, new_value in selected_config.items():
                # Support both display names and schema keys as input
                # First check if key is a display name that maps to a schema key
                schema_key = CONFIG_KEYS.get(key)
                # If not found, check if the key itself is already a valid schema key
                if not schema_key and key in valid_schema_keys:
                    schema_key = key
                if schema_key:
                    try:
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
                    except Exception as field_error:
                        # Handle missing fields by setting the value (this will create the field)
                        logger.debug("Field '{}' not found in existing storage, creating it with value: {}".format(schema_key, new_value))
                        try:
                            entity.set(schema_key, bool(new_value))
                            entity.set("last_used_date", timestamp)
                            changes_made = True
                        except Exception as set_error:
                            logger.warning("Could not create/set field '{}': {}".format(schema_key, set_error))
            
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
    """Load the MMI monitor configuration from extensible storage.

    Boolean toggles are keyed by display name or schema key. The four integer
    limits are always included under their schema keys.
    """
    config = {}
    try:
        data_storage = get_or_create_mmi_storage(doc)
        if not data_storage:
            return _empty_monitor_config(use_display_names)
            
        schema = MMIParameterSchema(data_storage)
        if schema.is_valid:
            for display_name, schema_key in CONFIG_KEYS.items():
                try:
                    value = schema.get(schema_key) or False  # Default to False if None
                except Exception as field_error:
                    # Handle missing fields gracefully (e.g., new fields in updated schema)
                    logger.debug("Field '{}' not found, using default: {}".format(schema_key, field_error))
                    value = False
                    
                # Store with either display name or schema key based on parameter
                if use_display_names:
                    config[display_name] = value
                else:
                    config[schema_key] = value
                logger.debug("Loaded config for '{}': {}".format(display_name, value))
            for key, default in MONITOR_LIMIT_DEFAULTS.items():
                config[key] = _read_stored_limit(schema, key, default)
            return config
        else:
            return _empty_monitor_config(use_display_names)
            
    except Exception as e:
        logger.error("Failed to load MMI monitor config: {}".format(str(e)))
        return _empty_monitor_config(use_display_names)
    