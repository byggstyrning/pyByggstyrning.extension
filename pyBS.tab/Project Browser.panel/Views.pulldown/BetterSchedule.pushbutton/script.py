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

try:
    # Find paths to both the current location and target script
    current_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', '..', '..', 'pyByggstyrning.toolbox', 'pyBS.tab', 'Project Browser.panel', 'Views.pulldown', 'BetterSchedule.pushbutton', 'script.py')
    
    if os.path.exists(script_path):
        
        # Make the target script directory the first in sys.path so imports work
        target_dir = os.path.dirname(script_path)
        if target_dir not in sys.path:
            sys.path.insert(0, target_dir)
        
        # Create a globals dict with __file__ pointing to the original script
        # This ensures that when the script looks for associated files, it finds them
        globals_dict = globals().copy()
        globals_dict['__file__'] = script_path
        
        # Read the script content
        with open(script_path, 'r') as f:
            script_content = f.read()
        
        exec(script_content, globals_dict)
        
    else:
        forms.alert("BetterSchedule script could not be found.\n\nChecked path:\n- {}".format(script_path), title="Script Not Found")
        
except Exception as ex:
    logger.error("Error attempting to run BetterSchedule: {}".format(ex))
    import traceback
    logger.error(traceback.format_exc())
    forms.alert("Failed to run BetterSchedule: {}".format(ex), title="Error")

