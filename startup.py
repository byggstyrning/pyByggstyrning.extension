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

# region agent log
import json as _json
import time as _time
import os.path as _op
_LOG_PATH = _op.join(_op.dirname(_op.abspath(__file__)), 'debug-cbb2c9.log')
def _dbg(msg, data=None, hyp=""):
    try:
        entry = {"sessionId":"cbb2c9","message":msg,"timestamp":int(_time.time()*1000),"hypothesisId":hyp,"data":data or {},"location":"startup.py"}
        with open(_LOG_PATH, 'a') as f:
            f.write(_json.dumps(entry) + '\n')
    except Exception:
        pass

_panel_check_attempts = [0]
_MAX_PANEL_CHECK = 30

def _panel_check_handler(sender, args):
    _panel_check_attempts[0] += 1
    try:
        ribbon = AdWindows.ComponentManager.Ribbon
        if not ribbon:
            if _panel_check_attempts[0] >= _MAX_PANEL_CHECK:
                HOST_APP.uiapp.Idling -= EventHandler[UIEvents.IdlingEventArgs](_panel_check_handler)
            return

        pybs_tab = None
        for tab in ribbon.Tabs:
            if tab.Id and 'pyBS' in str(tab.Id):
                pybs_tab = tab
                break
            if tab.Title and 'pyBS' in str(tab.Title):
                pybs_tab = tab
                break

        if not pybs_tab:
            _dbg("panel_check: pyBS tab not found yet, attempt {}".format(_panel_check_attempts[0]), hyp="H1,H2")
            if _panel_check_attempts[0] >= _MAX_PANEL_CHECK:
                _dbg("panel_check: giving up after {} attempts".format(_MAX_PANEL_CHECK), hyp="H1,H2")
                HOST_APP.uiapp.Idling -= EventHandler[UIEvents.IdlingEventArgs](_panel_check_handler)
            return

        _dbg("panel_check: pyBS tab found on attempt {}".format(_panel_check_attempts[0]),
             {"tab_id": str(pybs_tab.Id), "tab_title": str(pybs_tab.Title), "visible": pybs_tab.IsVisible}, hyp="H1")

        panels_data = []
        zone_panel = None
        for panel in pybs_tab.Panels:
            title = str(panel.Source.Title) if panel.Source else "NO_SOURCE"
            panels_data.append(title)
            if "3D" in title and "Zone" in title:
                zone_panel = panel
            elif title == "3D Zone":
                zone_panel = panel

        _dbg("panel_check: pyBS panels list", {"panels": panels_data}, hyp="H1,H2")

        if not zone_panel:
            _dbg("panel_check: 3D Zone panel NOT found in pyBS tab", hyp="H2")
        else:
            _dbg("panel_check: 3D Zone panel FOUND", {"title": str(zone_panel.Source.Title) if zone_panel.Source else "N/A"}, hyp="H2")
            items_data = []
            try:
                source = zone_panel.Source
                if source and hasattr(source, 'Items'):
                    for item in source.Items:
                        item_info = {
                            "type": str(type(item).__name__),
                            "id": str(getattr(item, 'Id', 'N/A')),
                            "text": str(getattr(item, 'Text', 'N/A')),
                            "visible": getattr(item, 'IsVisible', 'N/A'),
                            "enabled": getattr(item, 'IsEnabled', 'N/A'),
                        }
                        if hasattr(item, 'Items'):
                            sub_items = []
                            for sub in item.Items:
                                sub_info = {
                                    "type": str(type(sub).__name__),
                                    "id": str(getattr(sub, 'Id', 'N/A')),
                                    "text": str(getattr(sub, 'Text', 'N/A')),
                                    "visible": getattr(sub, 'IsVisible', 'N/A'),
                                    "enabled": getattr(sub, 'IsEnabled', 'N/A'),
                                }
                                sub_items.append(sub_info)
                            item_info["sub_items"] = sub_items
                        items_data.append(item_info)
                _dbg("panel_check: 3D Zone panel items", {"items": items_data, "item_count": len(items_data)}, hyp="H1,H3")
            except Exception as e:
                _dbg("panel_check: error reading 3D Zone items", {"error": str(e)}, hyp="H1,H3")

        try:
            HOST_APP.uiapp.Idling -= EventHandler[UIEvents.IdlingEventArgs](_panel_check_handler)
            _dbg("panel_check: unsubscribed from Idling", hyp="H1")
        except Exception:
            pass

    except Exception as e:
        _dbg("panel_check: ERROR", {"error": str(e)}, hyp="H1,H2,H3")
        if _panel_check_attempts[0] >= _MAX_PANEL_CHECK:
            try:
                HOST_APP.uiapp.Idling -= EventHandler[UIEvents.IdlingEventArgs](_panel_check_handler)
            except Exception:
                pass

try:
    HOST_APP.uiapp.Idling += EventHandler[UIEvents.IdlingEventArgs](_panel_check_handler)
    _dbg("panel_check: Idling handler registered", hyp="H1")
except Exception as e:
    _dbg("panel_check: failed to register Idling handler", {"error": str(e)}, hyp="H1")
# endregion

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
