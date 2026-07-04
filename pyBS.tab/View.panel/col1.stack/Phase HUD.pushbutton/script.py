# -*- coding: utf-8 -*-
"""Toggle a WPF phase badge pinned over the top-left corner of the view.

Normal click:
- ON: show the active view's Phase name as a click-through WPF overlay
  anchored to the viewport's top-left corner in screen pixels, so it stays
  put during pan/zoom.
- OFF: close the overlay and stop tracking.

Shift+Click:
- Force a redraw and show a diagnostics report.
"""

__title__ = "Phase HUD"
__author__ = "Byggstyrning AB"
__doc__ = ("Toggle a WPF overlay phase badge in the top-left corner of any "
           "view with a Phase. Shift+Click: refresh + diagnostics.")
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
    from revit.phase_label_wpf import (
        find_phase_hud_driver,
        start_phase_hud_driver,
        stop_phase_hud_driver,
        collect_diagnostics,
    )
    _PHASE_HUD_OK = True
except Exception as ex:
    _PHASE_HUD_OK = False
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
    driver = start_phase_hud_driver(uiapp, doc, logger=logger)
    if driver is None:
        forms.show_balloon(
            header="Phase HUD",
            text="Could not start phase HUD driver.",
            is_new=True)
        return False
    _sync_toggle_icon(True)
    return True


def _hud_off():
    stop_phase_hud_driver(doc)
    _sync_toggle_icon(False)
    return True


if __name__ == '__main__':
    if not _PHASE_HUD_OK:
        logger.error("phase_label_wpf import failed: {}".format(_IMPORT_ERROR))
        forms.alert(
            "Phase HUD failed to load:\n\n{}".format(_IMPORT_ERROR),
            title="Phase HUD")
    else:
        is_shift = script.get_config().get_option('shiftclick', False)
        driver = find_phase_hud_driver(doc)
        if is_shift:
            if driver is None:
                _hud_on()
                driver = find_phase_hud_driver(doc)
            if driver is not None:
                driver.refresh()
            forms.alert(
                collect_diagnostics(uiapp, doc),
                title="Phase HUD diagnostics")
        elif driver is not None:
            _hud_off()
        else:
            _hud_on()
