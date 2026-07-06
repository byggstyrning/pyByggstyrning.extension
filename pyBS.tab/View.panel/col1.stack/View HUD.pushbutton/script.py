# -*- coding: utf-8 -*-
"""Toggle the in-view context HUD pinned over the top middle of the view.

Normal click:
- ON: show a bar of context switchers centered over the viewport's top
  edge (screen-anchored, so it stays put during pan/zoom):
  - Phase — click: next phase, right-click: previous.
  - Active workset — click opens a dropdown to pick the active workset
    (workshared models only).
  The bar follows Revit's light/dark theme and idles at 50% opacity
  until hovered. Switchers hide themselves where they don't apply.
- OFF: close the bar and stop tracking.

Shift+Click:
- Force a redraw and show a diagnostics report.
"""

__title__ = "View HUD"
__author__ = "Byggstyrning AB"
__doc__ = ("Toggle the in-view context HUD: phase badge (click cycles) and "
           "active workset dropdown at the top of the view. "
           "Shift+Click: refresh + diagnostics.")
__highlight__ = 'new'
__persistentengine__ = True

import sys
import os.path as op

import clr
clr.AddReference('RevitAPI')

from pyrevit import script
from pyrevit import forms
from pyrevit import revit

script_path = __file__
pushbutton_dir = op.dirname(script_path)
stack_dir = op.dirname(pushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

_IMPORT_ERROR = None
try:
    from revit.context_switchers import (
        is_view_hud_running,
        start_view_hud,
        stop_view_hud,
        collect_diagnostics,
    )
    _VIEW_HUD_OK = True
except Exception as ex:
    _VIEW_HUD_OK = False
    _IMPORT_ERROR = str(ex)

logger = script.get_logger()
doc = revit.doc
uiapp = __revit__


def _sync_toggle_icon(active):
    try:
        script.toggle_icon(bool(active))
    except Exception:
        pass


def _hud_on():
    host = start_view_hud(uiapp, doc, logger=logger)
    if host is None:
        forms.show_balloon(
            header="View HUD",
            text="Could not start the view HUD.",
            is_new=True)
        return False
    _sync_toggle_icon(True)
    return True


def _hud_off():
    stop_view_hud(doc)
    _sync_toggle_icon(False)
    return True


if __name__ == '__main__':
    if not _VIEW_HUD_OK:
        logger.error(
            "context_switchers import failed: {}".format(_IMPORT_ERROR))
        forms.alert(
            "View HUD failed to load:\n\n{}".format(_IMPORT_ERROR),
            title="View HUD")
    else:
        # pyRevit signals Shift+Click as "config mode" via EXEC_PARAMS,
        # not as a persisted config option
        try:
            from pyrevit import EXEC_PARAMS
            is_shift = bool(EXEC_PARAMS.config_mode)
        except Exception:
            is_shift = script.get_config().get_option('shiftclick', False)
        running = is_view_hud_running(doc)
        if is_shift:
            if running:
                start_view_hud(uiapp, doc, logger=logger)
            else:
                _hud_on()
            forms.alert(
                collect_diagnostics(uiapp, doc),
                title="View HUD diagnostics")
        elif running:
            _hud_off()
        else:
            _hud_on()
