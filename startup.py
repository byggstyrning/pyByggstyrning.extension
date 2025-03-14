# -*- coding: utf-8 -*-
"""
Startup script for pyByggstyrning extension.
Clones the MMI panel to the Modify tab when a document is opened.
"""

import clr
clr.AddReference('AdWindows')
import Autodesk.Windows as AdWindows

from pyrevit import HOST_APP
from System import EventHandler
from Autodesk.Revit.DB import Events

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
            if panel.Source and "MMI" in panel.Source.Title:
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

# Run on document open on event
def doc_opening_handler(sender, args):
    find_and_clone_mmi_panel()

# Register the event handler to load the MMI panel when a document is opened
HOST_APP.app.DocumentOpening += \
    EventHandler[Events.DocumentOpeningEventArgs](
        doc_opening_handler
    ) 