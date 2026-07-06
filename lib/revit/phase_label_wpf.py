# -*- coding: utf-8 -*-
"""Phase HUD: view phase badge built on the reusable revit.view_hud driver.

Shows the active view's Phase name centered over the viewport's top edge,
in any graphical view that has a Phase parameter (3D, plan, section,
elevation, ...). The badge follows Revit's light/dark theme, idles at 50%
opacity, and is clickable: left-click switches the view to the next project
phase, right-click to the previous one (via ExternalEvent + transaction).
Beyond that parameter change, nothing is written to the model and nothing
prints.

All overlay mechanics (positioning, theming, hide-on-window-move, click
plumbing) live in revit.view_hud — build more HUDs like this by supplying
different providers to ViewHudDriver.
"""

import sys

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import BuiltInParameter, Transaction

from revit.compat import get_element_id_value
from revit.phase_label import _label_text_for_view
from revit.view_hud import (
    ViewHudDriver,
    find_hud_driver,
    stop_hud_driver,
    get_uiview,
    window_rect,
    revit_theme_is_dark,
)

_HUD_ID = 'phase'


def _ordered_phase_ids(document):
    """Project phases in chronological order (past -> future)."""
    ids = []
    try:
        for phase in document.Phases:
            ids.append(phase.Id)
    except Exception:
        pass
    return ids


def _shift_view_phase(uiapp, step):
    """Set the active view's Phase to the next/previous project phase.

    Must run inside a Revit API context (transaction). Returns the new
    phase name, or None when the phase cannot be changed.
    """
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    doc = uidoc.Document
    view = uidoc.ActiveView
    if view is None:
        return None
    param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
    if param is None or param.IsReadOnly:
        return None
    phase_ids = _ordered_phase_ids(doc)
    if len(phase_ids) < 2:
        return None
    id_values = [get_element_id_value(p) for p in phase_ids]
    try:
        idx = id_values.index(get_element_id_value(param.AsElementId()))
    except ValueError:
        idx = 0
    new_id = phase_ids[(idx + step) % len(phase_ids)]
    t = Transaction(doc, 'Switch view phase')
    t.Start()
    try:
        param.Set(new_id)
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        return None
    try:
        return doc.GetElement(new_id).Name
    except Exception:
        return None


def _phase_tooltip(text):
    return (u"Phase: {}\nClick: next phase. "
            u"Right-click: previous phase.".format(text))


def find_phase_hud_driver(document):
    return find_hud_driver(document, _HUD_ID)


def stop_phase_hud_driver(document):
    return stop_hud_driver(document, _HUD_ID)


def start_phase_hud_driver(uiapp, document, logger=None):
    driver = find_phase_hud_driver(document)
    if driver is not None:
        driver.refresh()
        return driver
    driver = ViewHudDriver(
        uiapp, document, _HUD_ID,
        text_provider=_label_text_for_view,
        tooltip_provider=_phase_tooltip,
        on_left_click=lambda ua: _shift_view_phase(ua, 1),
        on_right_click=lambda ua: _shift_view_phase(ua, -1),
        logger=logger)
    driver.start()
    return driver


def collect_diagnostics(uiapp, document):
    """Step-by-step status report for troubleshooting a blank HUD."""
    lines = []

    def add(key, fn):
        try:
            lines.append('{}: {}'.format(key, fn()))
        except Exception as ex:
            lines.append('{}: ERROR {}'.format(key, ex))

    add('engine', lambda: sys.version.replace('\n', ' '))

    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        lines.append('active_uidoc: None')
        return '\n'.join(lines)
    add('doc_match', lambda: uidoc.Document.Equals(document))
    view = uidoc.ActiveView
    if view is None:
        lines.append('active_view: None')
        return '\n'.join(lines)
    add('active_view', lambda: u'{} ({})'.format(view.Name, type(view).__name__))
    add('is_template', lambda: view.IsTemplate)
    add('phase_param', lambda: view.get_Parameter(
        BuiltInParameter.VIEW_PHASE) is not None)
    text = _label_text_for_view(document, view)
    lines.append(u'phase_text: {}'.format(text))
    add('phase_count', lambda: len(_ordered_phase_ids(document)))

    uiview = get_uiview(uiapp, view.Id)
    lines.append('uiview_found: {}'.format(uiview is not None))
    if uiview is not None:
        add('window_rect', lambda: str(uiview.GetWindowRectangle()))

    driver = find_phase_hud_driver(document)
    lines.append('driver_running: {}'.format(driver is not None))
    if driver is not None:
        lines.append('window_built: {}'.format(driver._window is not None))
        lines.append('window_visible_flag: {}'.format(driver._visible))
        if driver._window is not None:
            add('window_pos', lambda: '{}, {} ({}x{})'.format(
                driver._window.Left, driver._window.Top,
                driver._window.ActualWidth, driver._window.ActualHeight))
            add('window_is_visible', lambda: driver._window.IsVisible)
            add('window_text', lambda: driver._text_block.Text)
        lines.append('owner_hwnd: {}'.format(driver._owner_hwnd_int))
        add('owner_rect', lambda: str(window_rect(driver._owner_hwnd_int)))
        add('theme_dark', revit_theme_is_dark)
        lines.append('last_error: {}'.format(driver._last_error))
    return '\n'.join(lines)
