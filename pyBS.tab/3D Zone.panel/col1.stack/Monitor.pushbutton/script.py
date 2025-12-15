# -*- coding: utf-8 -*-
"""Toggles the 3D Zone Monitor on and off.

When active, the monitor watches for element movement and automatically
updates parameters based on configured zone mappings.
When inactive, it does nothing.
"""

__title__ = "Monitor"
__author__ = "Byggstyrning AB"
__doc__ = "Toggle 3D Zone Monitor on/off for the current session"
__highlight__ = 'new'

# Import standard libraries
import sys
import os

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit
from pyrevit.coreutils.ribbon import ICON_MEDIUM
from pyrevit.revit import ui
import pyrevit.extensions as exts

# Add the extension directory to the path
import os.path as op
script_path = __file__
pushbutton_dir = op.dirname(script_path)
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Initialize logger
logger = script.get_logger()

# Import zone3d libraries
try:
    from zone3d import config, monitor
except ImportError as e:
    logger.error("Failed to import zone3d libraries: {}".format(e))
    forms.alert("Failed to import required libraries. Check logs for details.")
    script.exit()

# --- Button Initialization ---

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    """Initialize the button icon based on the current active state."""
    try:
        on_icon = ui.resolve_icon_file(script_cmp.directory, exts.DEFAULT_ON_ICON_FILE)
        off_icon = ui.resolve_icon_file(script_cmp.directory, exts.DEFAULT_OFF_ICON_FILE)

        button_icon = script_cmp.get_bundle_file(
            on_icon if monitor.is_monitor_active() else off_icon
        )
        ui_button_cmp.set_icon(button_icon, icon_size=ICON_MEDIUM)
    except Exception as e:
        logger.error("Error initializing 3D Zone Monitor button: {}".format(e))

# --- Main Execution ---

if __name__ == '__main__':
    was_active = monitor.is_monitor_active()
    new_active_state = not was_active

    success = False
    if new_active_state:
        # Activate: Register handlers
        logger.info("Activating 3D Zone Monitor...")
        
        # Check if there are any enabled configurations
        configs = config.get_enabled_configs(revit.doc)
        if not configs:
            forms.alert(
                "No enabled configurations found.\n\n"
                "Please configure mappings using the Config button first.",
                title="No Configurations",
                exitscript=True
            )
        
        if monitor.register_event_handlers(revit.doc):
            monitor.set_monitor_active(True)
            script.toggle_icon(new_active_state)
            
            # Populate initial location cache
            monitor.populate_initial_location_cache(revit.doc, configs)
            
            # Show activation message
            config_names = [cfg.get("name", "Unknown") for cfg in configs]
            config_list = "\n".join(["{}. {}".format(i+1, name) for i, name in enumerate(config_names)])
            
            forms.show_balloon(
                header="3D Zone Monitor",
                text="Monitor activated\n\n{} active configuration(s):\n{}".format(
                    len(configs), config_list
                ),
                is_new=True
            )
            success = True
        else:
            forms.show_balloon(
                header="Error",
                text="Failed to activate 3D Zone Monitor",
                tooltip="Check logs for details",
                is_new=True
            )
    else:
        # Deactivate: Deregister handlers
        logger.info("Deactivating 3D Zone Monitor...")
        
        if monitor.deregister_event_handlers(revit.doc):
            monitor.set_monitor_active(False)
            script.toggle_icon(new_active_state)
            
            # Clear cache
            monitor.clear_location_cache()
            logger.info("Cleared element location cache")
            
            success = True
        else:
            forms.show_balloon(
                header="Error",
                text="Failed to deactivate 3D Zone Monitor",
                tooltip="Check logs for details",
                is_new=True
            )

    if success:
        logger.info("3D Zone Monitor state toggled to: {}".format("ON" if new_active_state else "OFF"))
    else:
        logger.error("Failed to toggle 3D Zone Monitor state.")

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.

