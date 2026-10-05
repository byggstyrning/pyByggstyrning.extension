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


def _maybe_start_view_hud(document=None):
    """Start the View HUD when Settings auto-start is on."""
    try:
        import sys
        import os.path as op
        extension_dir = op.dirname(op.abspath(__file__))
        lib_path = op.join(extension_dir, 'lib')
        if lib_path not in sys.path:
            sys.path.insert(0, lib_path)
        from revit.view_hud_config import is_auto_start_enabled
        if not is_auto_start_enabled():
            return
        uiapp = HOST_APP.uiapp
        if uiapp is None:
            return
        doc = document
        if doc is None:
            uidoc = uiapp.ActiveUIDocument
            if uidoc is None:
                return
            doc = uidoc.Document
        from revit.context_switchers import (
            is_view_hud_running,
            start_view_hud,
        )
        if is_view_hud_running(doc):
            return
        start_view_hud(uiapp, doc)
    except Exception as ex:
        script_logger.debug("View HUD auto-start skipped: {}".format(ex))


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
    _maybe_start_view_hud()

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


def doc_opened_handler(sender, args):
    find_and_clone_mmi_panel()
    try:
        _maybe_start_view_hud(args.Document)
    except Exception:
        pass

try:
    HOST_APP.app.DocumentOpened += EventHandler[Events.DocumentOpenedEventArgs](
        doc_opened_handler)
except Exception:
    pass

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

# Register the CDE dockable panel (hosts the CDE Schedule cockpit).
# Dockable panes can only be registered during Revit startup; the ribbon
# button (CDE panel > CDE Panel) merely toggles its visibility.
try:
    import sys
    import os.path as op
    extension_dir = op.dirname(op.abspath(__file__))
    lib_path = op.join(extension_dir, 'lib')
    if lib_path not in sys.path:
        sys.path.insert(0, lib_path)

    from pyrevit import forms
    from cde.panel_ui import CDESchedulePanel
    # On pyRevit reload the pane is already registered; registering again
    # would construct an orphan panel instance (with its event subscriptions).
    if forms.is_registered_dockable_panel(CDESchedulePanel):
        script_logger.debug("CDE panel already registered; skipping.")
    else:
        forms.register_dockable_panel(CDESchedulePanel, default_visible=False)
        script_logger.info("CDE panel registered.")
except Exception as e:
    # NOTE: Revit only allows dockable-pane registration during its startup
    # sequence. A mid-session pyRevit reload re-runs this script but the
    # registration is rejected - a full Revit restart is needed the first time.
    import traceback
    script_logger.warning(
        "Could not register CDE panel (a full Revit restart is required "
        "the first time): {}\n{}".format(str(e), traceback.format_exc()))
