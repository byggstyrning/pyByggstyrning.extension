# -*- coding: utf-8 -*-
"""Write parameters to elements based on configured zone mappings.

Executes all enabled configurations in order, writing parameters from
spatial elements (Rooms, Spaces, Areas, Mass/Generic Model) to contained elements.
"""

__title__ = "Write"
__author__ = "Byggstyrning AB"
__doc__ = "Execute all enabled 3D Zone configurations to write parameters to elements"

# Import standard libraries
import sys
import os

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

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
    from zone3d import config, core
except ImportError as e:
    logger.error("Failed to import zone3d libraries: {}".format(e))
    forms.alert("Failed to import required libraries. Check logs for details.")
    script.exit()

# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Load enabled configurations
    configs = config.get_enabled_configs(doc)
    
    if not configs:
        forms.alert(
            "No enabled configurations found.\n\n"
            "Please configure mappings using the Config button first.",
            title="No Configurations",
            exitscript=True
        )
    
    # Show confirmation
    config_names = [cfg.get("name", "Unknown") for cfg in configs]
    config_list = "\n".join(["{}. {}".format(i+1, name) for i, name in enumerate(config_names)])
    
    result = forms.alert(
        "Execute {} configuration(s)?\n\n{}".format(len(configs), config_list),
        title="Execute 3D Zone Configurations",
        ok=False,
        yes=True,
        no=True
    )
    
    if not result:
        script.exit()
    
    try:
        summary = core.execute_all_configurations(doc)
        
        # Build results message
        results_text = "Execution Complete\n\n"
        results_text += "Total Configurations: {}\n".format(summary["total_configs"])
        results_text += "Total Elements Updated: {}\n".format(summary["total_elements_updated"])
        results_text += "Total Parameters Copied: {}\n\n".format(summary["total_parameters_copied"])
        
        # Add per-configuration results
        if summary["config_results"]:
            results_text += "Per Configuration:\n"
            for result in summary["config_results"]:
                config_name = result.get("config_name", "Unknown")
                elements_updated = result.get("elements_updated", 0)
                params_copied = result.get("parameters_copied", 0)
                errors = result.get("errors", [])
                
                results_text += "\n{}:\n".format(config_name)
                results_text += "  Elements Updated: {}\n".format(elements_updated)
                results_text += "  Parameters Copied: {}\n".format(params_copied)
                
                if errors:
                    results_text += "  Errors: {}\n".format(len(errors))
                    for error in errors[:3]:  # Show first 3 errors
                        results_text += "    - {}\n".format(error[:60])
                    if len(errors) > 3:
                        results_text += "    ... and {} more\n".format(len(errors) - 3)
        
        # Show results
        forms.alert(results_text, title="3D Zone Write Results")
    
    except Exception as e:
        error_msg = "Error executing configurations: {}".format(str(e))
        logger.error(error_msg)
        forms.alert(error_msg, title="Error", exitscript=True)

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.

