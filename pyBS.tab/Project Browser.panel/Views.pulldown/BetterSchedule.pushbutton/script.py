# -*- coding: utf-8 -*-
__title__ = "Better Schedule"
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

from pyrevit import forms, script
import os
import sys

logger = script.get_logger()

_SHIM_FILE = os.path.abspath(__file__)
_SHIM_DIR = os.path.dirname(_SHIM_FILE)
_EXT_LIB = os.path.normpath(os.path.join(_SHIM_DIR, "..", "..", "..", "..", "lib"))
if _EXT_LIB not in sys.path:
    sys.path.insert(0, _EXT_LIB)

from toolbox_probe import find_better_schedule_script

_UNAVAILABLE_MSG = (
    "Better Schedule is not available on this machine.\n\n"
    "Install the restricted pyByggstyrning.toolbox lib extension if you have access, "
    "then reload pyRevit."
)

try:
    script_path = find_better_schedule_script(shim_file=_SHIM_FILE)

    if script_path and os.path.isfile(script_path):
        target_dir = os.path.dirname(script_path)
        if target_dir not in sys.path:
            sys.path.insert(0, target_dir)

        toolbox_root = os.path.normpath(os.path.join(target_dir, ".."))
        toolbox_lib = os.path.join(toolbox_root, "lib")
        if os.path.isdir(toolbox_lib) and toolbox_lib not in sys.path:
            sys.path.insert(0, toolbox_lib)

        globals_dict = globals().copy()
        globals_dict["__file__"] = script_path

        with open(script_path, "r") as f:
            script_content = f.read()

        exec(script_content, globals_dict)
    else:
        forms.alert(_UNAVAILABLE_MSG, title="Better Schedule")

except Exception as ex:
    logger.error("Error attempting to run BetterSchedule: {}".format(ex))
    import traceback

    logger.error(traceback.format_exc())
    forms.alert("Failed to run Better Schedule: {}".format(ex), title="Error")
