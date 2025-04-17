# -*- coding: utf-8 -*-
"""MMI Settings.

This tool allows users to set the MMI settings for the current project.

"""

__title__ = "Settings"
__author__ = "Byggstyrning AB"
__doc__ = "MMI Settings"
__highlight__ = 'new'

# Import standard libraries
import sys

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Add the extension directory to the path - FIXED PATH RESOLUTION
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

# Try direct import from current directory's parent path
sys.path.append(op.dirname(op.dirname(panel_dir)))

# Initialize logger
logger = script.get_logger()

# Import from MMI library modules
from mmi.config import CONFIG_KEYS
from mmi.core import get_or_create_mmi_storage, get_mmi_parameter_name, save_mmi_parameter, save_monitor_config, load_monitor_config

# Import MMI schema
from mmi.schema import MMIParameterSchema

# Import revit utils
from revit.revit_utils import get_available_parameters

def open_mmi_parameter_selector():
    """Open the MMI parameter selection form."""
    # Get the current MMI parameter
    current_parameter = get_mmi_parameter_name(revit.doc)
    
    # Get all available parameters
    available_parameters = get_available_parameters()
    
    if not available_parameters:
        forms.alert("No parameters found in the project.", title="MMI Parameter")
        return False
    
    # Create parameter selection form
    selected_parameter = forms.ask_for_one_item(
        available_parameters,
        default=current_parameter,
        prompt="Select instance parameter for MMI:",
        title="MMI Parameter"
    )
    
    if selected_parameter:
        # Save the selected parameter
        if save_mmi_parameter(revit.doc, selected_parameter):
            forms.show_balloon(
                header="MMI Parameter Set",
                text="MMI parameter set to: {}".format(selected_parameter),
                tooltip="MMI parameter set to: {}".format(selected_parameter),
                is_new=True
            )
            return True
        else:
            forms.alert(
                "Failed to save MMI parameter setting. See log for details.",
                title="MMI Parameter Error"
            )
            return False
    return False  # User cancelled

if __name__ == '__main__':
    # Load the current configuration
    current_config = load_monitor_config(revit.doc, use_display_names=True)
    
    switches = {name: current_config.get(name, False) for name in CONFIG_KEYS.keys()}
    rops, rswitches = forms.CommandSwitchWindow.show(
        ['⚙️ Set MMI Parameter', '💾 Save Config'],
        switches=switches,
        message="MMI Monitor configuration:",
        recognize_access_key=False
    )

    if rops is not None and 'Set MMI Parameter' in rops:
        # Show the MMI parameter selection form
        open_mmi_parameter_selector()
    elif rswitches is not None:
        # Save the configuration
        if save_monitor_config(revit.doc, rswitches):
            pass
        else:
            forms.warning("Failed to save MMI Monitor configuration. See log for details.", title="MMI Monitor Config Error")
    else:
        pass
