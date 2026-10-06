# -*- coding: utf-8 -*-
"""Configuration constants and functions for MMI monitor."""

from pyrevit import script
from pyrevit.userconfig import user_config

# Initialize logger
logger = script.get_logger()

# Constants
CONFIG_SECTION = 'MMIMonitor'
CONFIG_KEY_ACTIVE = 'isActive'

# Per-warning defaults. MMI_THRESHOLD remains the fallback when a pin limit is missing.
DEFAULT_PIN_MMI_LIMIT = 400
DEFAULT_MOVE_MMI_LIMIT = 425
DEFAULT_TYPE_INSTANCE_LIMIT = 375
DEFAULT_INSTANCE_PARAM_LIMIT = 375
MMI_THRESHOLD = DEFAULT_PIN_MMI_LIMIT

MONITOR_LIMIT_DEFAULTS = {
    "pin_mmi_limit": DEFAULT_PIN_MMI_LIMIT,
    "move_mmi_limit": DEFAULT_MOVE_MMI_LIMIT,
    "type_instance_limit": DEFAULT_TYPE_INSTANCE_LIMIT,
    "instance_param_limit": DEFAULT_INSTANCE_PARAM_LIMIT,
}

# MMI values offered as defaults (e.g. Settings "Default on new instances" combo)
STANDARD_MMI_VALUES = (
    "100", "125", "150", "175",
    "200", "225", "250", "275", "300", "325", "350", "375",
    "400", "425", "450", "475",
)

# Standard config keys mapping
CONFIG_KEYS = {
    "✅ Attempt to fix MMI values": "validate_mmi",
    "🔒 Pin elements": "pin_elements",
    "⚠️ Warn when moving elements": "warn_on_move",
    "🔄 Check MMI after sync": "check_mmi_after_sync",
    "🆕 Default on new instances": "default_on_new_instances",
    "⚠️ Warn on type changes": "warn_on_type_change",
    "✏️ Warn on instance parameter edits": "warn_on_instance_params",
}

def is_monitor_active():
    """Check if the MMI monitor is currently active based on user config."""
    try:
        # The correct way to use user_config is to directly access sections as attributes
        if not hasattr(user_config, CONFIG_SECTION):
            return False
        
        # Get the section and check the active value
        section = getattr(user_config, CONFIG_SECTION)
        return section.get_option(CONFIG_KEY_ACTIVE, default_value=False)
    except Exception as ex:
        logger.error("Error checking monitor state: {}".format(ex))
        return False

def set_monitor_active(is_active):
    """Set the MMI monitor active state in user config."""
    try:
        # Make sure the section exists
        if not hasattr(user_config, CONFIG_SECTION):
            user_config.add_section(CONFIG_SECTION)
        
        # Get the section and set the value
        section = getattr(user_config, CONFIG_SECTION)
        section.set_option(CONFIG_KEY_ACTIVE, is_active)
        
        # Save the changes
        user_config.save_changes()
        logger.debug("MMI Monitor active state set to: {}".format(is_active))
    except Exception as ex:
        logger.error("Error setting monitor state: {}".format(ex)) 