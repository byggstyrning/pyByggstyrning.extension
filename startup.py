# -*- coding: utf-8 -*-
"""
Startup script for pyByggstyrning extension.
Clones the MMI panel to the Modify tab on startup and when a document is opened.

NOTE: The Switchback API has been moved to pyValidator.extension to avoid conflicts.
"""

import clr
clr.AddReference('AdWindows')
import Autodesk.Windows as AdWindows

from pyrevit import HOST_APP
from System import EventHandler
from Autodesk.Revit.DB import Events
from Autodesk.Revit.UI import Events as UIEvents
from pyrevit.coreutils import logger

script_logger = logger.get_logger('pyByggstyrning_startup')

def find_and_clone_mmi_panel():
    ribbon = AdWindows.ComponentManager.Ribbon
    if not ribbon:
        return
    
    modify_tab = next((tab for tab in ribbon.Tabs 
                      if tab.Title == "Modify" and tab.IsVisible), None)
    
    if not modify_tab:
        return
        
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
        
    if any(panel.Source and panel.Source.Title == source_panel.Source.Title 
           for panel in modify_tab.Panels):
        return
        
    modify_tab.Panels.Add(AdWindows.RibbonPanel())
    new_panel = modify_tab.Panels[modify_tab.Panels.Count - 1]
    new_panel.Source = source_panel.Source.Clone()
    new_panel.IsEnabled = True

# Attempt immediate clone (works on extension reload when pyBS tab already exists)
find_and_clone_mmi_panel()

# Deferred clone via Idling event: pyRevit 6.x runs startup.py before extension
# tabs are created on the ribbon. Use Idling to poll until the pyBS/MMI panel appears.
_idling_attempts = [0]
_MAX_IDLING_ATTEMPTS = 50

def _idling_handler(sender, args):
    _idling_attempts[0] += 1

    ribbon = AdWindows.ComponentManager.Ribbon
    if not ribbon:
        if _idling_attempts[0] >= _MAX_IDLING_ATTEMPTS:
            HOST_APP.uiapp.Idling -= EventHandler[UIEvents.IdlingEventArgs](_idling_handler)
        return

    source_found = False
    for tab in ribbon.Tabs:
        if not tab.IsVisible:
            continue
        for panel in tab.Panels:
            if panel.Source and panel.Source.Title == "MMI":
                source_found = True
                break
        if source_found:
            break

    if not source_found:
        if _idling_attempts[0] >= _MAX_IDLING_ATTEMPTS:
            HOST_APP.uiapp.Idling -= EventHandler[UIEvents.IdlingEventArgs](_idling_handler)
        return

    find_and_clone_mmi_panel()

    try:
        HOST_APP.uiapp.Idling -= EventHandler[UIEvents.IdlingEventArgs](_idling_handler)
    except Exception:
        pass

try:
    HOST_APP.uiapp.Idling += EventHandler[UIEvents.IdlingEventArgs](_idling_handler)
except Exception:
    pass

# Also clone on document open (handles case where extension reloads mid-session)
def doc_opening_handler(sender, args):
    find_and_clone_mmi_panel()

HOST_APP.app.DocumentOpening += \
    EventHandler[Events.DocumentOpeningEventArgs](
        doc_opening_handler
    )

# Register IFC export handler for 3D Zone parameter mapping
try:
    import sys
    import os.path as op
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
