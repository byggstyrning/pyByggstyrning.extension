# -*- coding: utf-8 -*-
"""
Startup script for pyByggstyrning extension.
Clones the MMI panel to the Modify tab when a document is opened.

NOTE: The Switchback API has been moved to pyValidator.extension to avoid conflicts.
"""

import clr
clr.AddReference('AdWindows')
import Autodesk.Windows as AdWindows

from pyrevit import HOST_APP
from System import EventHandler
from Autodesk.Revit.DB import Events
from pyrevit.coreutils import logger

# Create a logger instance for this script
script_logger = logger.get_logger('pyByggstyrning_startup')

def find_and_clone_mmi_panel():
    ribbon = AdWindows.ComponentManager.Ribbon
    if not ribbon:
        return
    
    # Find the Modify tab
    modify_tab = next((tab for tab in ribbon.Tabs 
                      if tab.Title == "Modify" and tab.IsVisible), None)
    
    if not modify_tab:
        return
        
    # Find the source MMI panel
    source_panel = None
    for tab in ribbon.Tabs:
        if not tab.IsVisible:
            continue
        for panel in tab.Panels:
            if panel.Source and "MMI" == panel.Source.Title:
                source_panel = panel
                break
        if source_panel:
            break
            
    if not source_panel:
        return
        
    # Check if panel already exists in Modify tab
    if any(panel.Source and panel.Source.Title == source_panel.Source.Title 
           for panel in modify_tab.Panels):
        return
        
    # Clone panel to Modify tab
    modify_tab.Panels.Add(AdWindows.RibbonPanel())
    new_panel = modify_tab.Panels[modify_tab.Panels.Count - 1]
    new_panel.Source = source_panel.Source.Clone()
    new_panel.IsEnabled = True

# Run on startup / reload
find_and_clone_mmi_panel()

# Run on document open event
def doc_opening_handler(sender, args):
    find_and_clone_mmi_panel()

# Register the event handler to load the MMI panel when a document is opened
HOST_APP.app.DocumentOpening += \
    EventHandler[Events.DocumentOpeningEventArgs](
        doc_opening_handler
    )

# Register IFC export handler for 3D Zone parameter mapping
try:
    import sys
    import os.path as op
    # Add lib path for zone3d import
    extension_dir = op.dirname(op.abspath(__file__))
    lib_path = op.join(extension_dir, 'lib')
    if lib_path not in sys.path:
        sys.path.insert(0, lib_path)
    
    from zone3d import ifc_export
    if ifc_export.register_ifc_export_handler():
        pass
    else:
        script_logger.warning("Failed to register 3D Zone IFC export handler")
except Exception as e:
    script_logger.warning("Could not register 3D Zone IFC export handler: {}".format(str(e)))

