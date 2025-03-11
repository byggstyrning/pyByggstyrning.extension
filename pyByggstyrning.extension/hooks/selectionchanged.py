# -*- coding: utf-8 -*-
"""
Selection changed hook for pyByggstyrning extension.
This script detects when elements are selected in Revit and
dynamically adds MMI buttons to the Manage tab.
"""

import clr
import os
import sys

# Import .NET and Revit API
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('AdWindows')  # Used to access Revit ribbon
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
import Autodesk.Windows as AdWindows
from System import EventHandler, Uri, Object
from System.Windows.Media.Imaging import BitmapImage

# Import pyRevit libraries
from pyrevit import revit, DB, UI, script, forms
from pyrevit.framework import wpf, Media

# Set up logger
logger = script.get_logger()

# Get the active Revit application instance
uiapp = __revit__

def add_mmi_button_to_manage_tab():
    try:
        # Get the current document and selection
        uidoc = uiapp.ActiveUIDocument
        if not uidoc:
            return
            
        doc = uidoc.Document
        selection = uidoc.Selection
        
        # Get selected elements
        selected_ids = selection.GetElementIds()
        
        # Log selection info
        if selected_ids and len(selected_ids) > 0:
            logger.debug("Selection changed: {} elements selected".format(len(selected_ids)))
        
        # Get the Revit ribbon
        ribbon = AdWindows.ComponentManager.Ribbon
        
        # Only proceed if elements are selected and ribbon is available
        if not selected_ids or len(selected_ids) == 0 or not ribbon:
            remove_mmi_button_from_manage_tab()
            return
        
        # Find the Manage tab
        manage_tab = next((tab for tab in ribbon.Tabs 
                          if tab.Title == "Manage" and tab.IsVisible), None)
        
        if not manage_tab:
            logger.warning("Manage tab not found or not visible")
            return
        
        # Check if our panel already exists
        mmi_panel = next((panel for panel in manage_tab.Panels 
                         if panel.Source and panel.Source.Title == "MMI Selection"), None)
        
        # If our panel doesn't exist, create it
        if not mmi_panel:
            # Create a new panel
            manage_tab.Panels.Add(AdWindows.RibbonPanel())
            mmi_panel = manage_tab.Panels[-1]  # Get the last panel (our new one)
            
            # Set up the panel source
            mmi_panel.Source = AdWindows.RibbonPanelSource()
            mmi_panel.Source.Title = "MMI Selection"
            mmi_panel.Source.Name = "MMISelectionPanel"
            
            # Find our 200 button's information from the original panel
            bs_tab = next((tab for tab in ribbon.Tabs if tab.Title == "BS"), None)
            if bs_tab:
                mmi_panel_original = next((panel for panel in bs_tab.Panels 
                                         if panel.Source and panel.Source.Title == "MMI"), None)
                if mmi_panel_original:
                    # Find the 200 button
                    button_200 = None
                    for item in mmi_panel_original.Source.Items:
                        if hasattr(item, 'Text') and item.Text and "200" in item.Text:
                            button_200 = item
                            break
                    
                    if button_200:
                        # Clone and add the button to our panel
                        mmi_panel.Source.Items.Add(button_200.Clone())
                        
                        # Make sure the panel and button are visible
                        mmi_panel.IsEnabled = True
                        mmi_panel.IsVisible = True
                        new_button = mmi_panel.Source.Items[-1]
                        new_button.IsEnabled = True
                        new_button.IsVisible = True
                        logger.info("Added 200 button to Manage tab")
    except Exception as e:
        logger.error("Error adding MMI button to Manage tab: {}".format(e))

def remove_mmi_button_from_manage_tab():
    try:
        # Get the Revit ribbon
        ribbon = AdWindows.ComponentManager.Ribbon
        
        if not ribbon:
            return
            
        # Find the Manage tab
        manage_tab = next((tab for tab in ribbon.Tabs 
                          if tab.Title == "Manage" and tab.IsVisible), None)
        
        if not manage_tab:
            return
        
        # Find our panel
        mmi_panel = next((panel for panel in manage_tab.Panels 
                         if panel.Source and panel.Source.Title == "MMI Selection"), None)
        
        # If our panel exists, remove it
        if mmi_panel:
            index = manage_tab.Panels.IndexOf(mmi_panel)
            if index >= 0:
                manage_tab.Panels.RemoveAt(index)
                logger.info("Removed MMI panel from Manage tab")
    except Exception as e:
        logger.error("Error removing MMI panel from Manage tab: {}".format(e))

# This function will be called by pyRevit when selection changes
def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    """
    Hook initialization.
    """
    global uiapp
    uiapp = __rvt__

# This is the main hook function that pyRevit will call
def __eventhandler__(sender, args):
    """
    Hook function called when selection changes.
    """
    add_mmi_button_to_manage_tab() 