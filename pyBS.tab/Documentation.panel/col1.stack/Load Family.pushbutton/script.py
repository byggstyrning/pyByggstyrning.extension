# -*- coding: utf-8 -*-
__title__ = "Load Family"
__author__ = "Jonatan Jacobsson"
__doc__ = """Loads the 3D View Reference family into the current project.
No prompts - just click and load.
"""

# Import libraries
import os
from pyrevit import revit, script

# Get the current Revit document
doc = revit.doc

logger = script.get_logger()

# Get the script directory
script_dir = script.get_script_path()
family_name = "3D View Reference.rfa"
family_path = os.path.join(script_dir, family_name)

# Check if the family file exists in the script directory
if not os.path.exists(family_path):
    # If not in the script directory, provide a default path or use a predefined one
    # You may need to adjust this path to where your family file is actually located
    default_path = os.path.join(os.path.dirname(script_dir), "families", family_name)
    
    if os.path.exists(default_path):
        family_path = default_path
    else:
        script.exit()

# Load the family into the project
try:
    with revit.Transaction("Load 3D View Reference Family"):
        family_loaded = doc.LoadFamily(family_path)
        
        if not family_loaded:
            logger.error("Failed to load family: {}".format(family_name))
except Exception as e:
    logger.error("Error loading family: {}".format(str(e)))

