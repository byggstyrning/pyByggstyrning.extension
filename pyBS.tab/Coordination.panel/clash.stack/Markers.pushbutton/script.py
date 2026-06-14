# -*- coding: utf-8 -*-
"""Toggle clash view markers on/off.

Normal click:
- ON: show temporary clash markers in open clash 3D views (run Clash Views first).
- OFF: clear all marker graphics.

Shift+Click:
- Refresh markers (re-cluster and redraw) without toggling state.
"""

__title__ = "Markers"
__author__ = "Byggstyrning AB"
__doc__ = "Toggle clash markers on/off. Shift+Click: refresh markers."
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

from revit.compat import get_element_id_value

_MARKERS_IMPORT_ERROR = None

try:
    from revit.clash_markers import (
        is_temporary_graphics_available,
        get_marker_session,
        is_marker_toggle_active,
        set_marker_toggle_active,
        find_marker_driver,
        start_or_get_driver,
        clean_marker_session,
        register_marker_session,
        ensure_marker_bitmaps_ready,
        ensure_temporary_graphics_handler,
        _is_view_sheet,
    )
    _MARKERS_OK = is_temporary_graphics_available()
except Exception as ex:
    _MARKERS_OK = False
    _MARKERS_IMPORT_ERROR = str(ex)
    get_marker_session = None
    is_marker_toggle_active = None
    set_marker_toggle_active = None
    find_marker_driver = None
    start_or_get_driver = None
    clean_marker_session = None
    register_marker_session = None
    ensure_marker_bitmaps_ready = None
    ensure_temporary_graphics_handler = None
    _is_view_sheet = None

logger = script.get_logger()
doc = revit.doc
uiapp = __revit__


def _session_view_points():
    if get_marker_session is None:
        return {}
    points = get_marker_session(doc)
    if points:
        return points
    driver = find_marker_driver(doc) if find_marker_driver else None
    if driver is not None:
        return dict(driver._sessions)
    return {}


def _sync_toggle_icon(active):
    try:
        script.toggle_icon(bool(active))
    except Exception:
        pass


def _markers_on(view_points):
    if ensure_marker_bitmaps_ready is not None:
        paths, bmp_errors = ensure_marker_bitmaps_ready()
        if bmp_errors:
            forms.alert(
                "Marker bitmap (BMP) creation failed:\n{}".format(
                    bmp_errors[0]),
                title="Clash Markers")
            return False
    if ensure_temporary_graphics_handler is not None:
        if not ensure_temporary_graphics_handler(doc, logger=logger):
            logger.warning(
                "TemporaryGraphicsHandler registration failed")
    register_marker_session(doc, view_points, toggle_active=True)
    driver = start_or_get_driver(
        uiapp, doc, view_points, get_element_id_value, logger=None)
    if driver is None:
        forms.show_balloon(
            header="Clash Markers",
            text="Could not start marker driver.",
            is_new=True)
        return False
    driver.set_enabled(True)
    set_marker_toggle_active(doc, True)
    _sync_toggle_icon(True)
    driver.refresh_all()
    try:
        uidoc = uiapp.ActiveUIDocument
        if (uidoc is not None and uidoc.ActiveView is not None
                and _is_view_sheet is not None
                and _is_view_sheet(uidoc.ActiveView)):
            forms.show_balloon(
                header="Clash Markers",
                text="Sheet: grouped badge per viewport. Double-click viewport for detail markers.",
                is_new=True)
    except Exception:
        pass
    return True


def _markers_off():
    if clean_marker_session is not None:
        clean_marker_session(doc)
    else:
        driver = find_marker_driver(doc) if find_marker_driver else None
        if driver is not None:
            driver.set_enabled(False)
        if set_marker_toggle_active is not None:
            set_marker_toggle_active(doc, False)
    _sync_toggle_icon(False)
    return True


def _markers_refresh(view_points):
    if not view_points:
        forms.alert(
            "No clash marker session in this model. Run Clash Views first.",
            title="Clash Markers")
        return False
    if not is_marker_toggle_active(doc):
        return _markers_on(view_points)
    driver = find_marker_driver(doc)
    if driver is None:
        return _markers_on(view_points)
    driver.set_enabled(True)
    driver.refresh_all()
    forms.show_balloon(
        header="Clash Markers",
        text="Markers refreshed ({} views).".format(len(view_points)),
        is_new=True)
    return True


if __name__ == '__main__':
    if not _MARKERS_OK:
        if _MARKERS_IMPORT_ERROR:
            logger.error(
                "clash_markers import failed: {}".format(_MARKERS_IMPORT_ERROR))
            forms.alert(
                "Clash markers failed to load:\n\n{}".format(
                    _MARKERS_IMPORT_ERROR),
                title="Clash Markers")
        else:
            forms.alert(
                "Clash markers require Revit 2022 or newer "
                "(TemporaryGraphicsManager API).",
                title="Clash Markers")
    else:
        view_points = _session_view_points()
        is_shift = script.get_config().get_option('shiftclick', False)
        if is_shift:
            _markers_refresh(view_points)
        else:
            was_active = False
            if is_marker_toggle_active is not None:
                was_active = is_marker_toggle_active(doc)
            driver = find_marker_driver(doc) if find_marker_driver else None
            if was_active and driver is None:
                was_active = False
                _sync_toggle_icon(False)
            if was_active:
                _markers_off()
            else:
                if not view_points:
                    forms.alert(
                        "No clash marker session in this model. "
                        "Run Clash Views first.",
                        title="Clash Markers")
                else:
                    _markers_on(view_points)
