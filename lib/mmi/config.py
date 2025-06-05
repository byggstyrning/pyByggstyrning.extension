# -*- coding: utf-8 -*-
"""Configuration constants and functions for MMI monitor."""

from pyrevit import script
from pyrevit.userconfig import user_config

# Initialize logger
logger = script.get_logger()

# Constants
CONFIG_SECTION = 'MMIMonitor'
CONFIG_KEY_ACTIVE = 'isActive'
MMI_THRESHOLD = 400

# Standard config keys mapping
CONFIG_KEYS = {
    "✅ Validate MMI": "validate_mmi",
    "🔒 Pin elements >=400": "pin_elements",
    "⚠️ Warn when moving elements >400": "warn_on_move",
    "🔄 Check MMI after sync": "check_mmi_after_sync"
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