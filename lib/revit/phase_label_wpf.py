# -*- coding: utf-8 -*-
"""Phase HUD: view phase switcher built on the revit.view_hud framework.

Shows the active view's Phase name on the HUD bar. The name is a cycle
control (left-click next, right-click previous). The caret opens a
dropdown of all project phases. Changes go through ExternalEvent +
transaction; nothing else is written to the model.

Overlay mechanics live in revit.view_hud. This module supplies the phase
options provider and the cycle / pick actions.
"""

import sys

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import BuiltInParameter, Transaction

from revit.compat import get_element_id_value, make_element_id
from revit.phase_label import _label_text_for_view
from revit.view_hud import (
    DropdownSwitcher,
    get_hud_host,
    find_hud_host,
    find_hud_item,
    remove_hud_item,
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


def _view_phase_param(view):
    if view is None:
        return None
    param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
    if param is None or param.IsReadOnly:
        return None
    return param


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
    param = _view_phase_param(view)
    if param is None:
        return None
    phase_ids = _ordered_phase_ids(doc)
    if len(phase_ids) < 2:
        return None
    id_values = [get_element_id_value(p) for p in phase_ids]
    try:
        idx = id_values.index(get_element_id_value(param.AsElementId()))
    except ValueError:
        idx = 0
    return _commit_view_phase(doc, param, phase_ids[(idx + step) % len(phase_ids)])


def _set_view_phase(uiapp, phase_id_value):
    """Set the active view's Phase to the given project phase id value."""
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    doc = uidoc.Document
    view = uidoc.ActiveView
    param = _view_phase_param(view)
    if param is None:
        return None
    return _commit_view_phase(doc, param, make_element_id(phase_id_value))


def _commit_view_phase(doc, param, new_id):
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


def _phase_options(document, view):
    """(current id, current name, [(id, name), ...]) or None if N/A."""
    text = _label_text_for_view(document, view)
    if not text:
        return None
    param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
    if param is None:
        return None
    current_id = get_element_id_value(param.AsElementId())
    options = []
    try:
        for phase in document.Phases:
            options.append((get_element_id_value(phase.Id), phase.Name))
    except Exception:
        return None
    if not options:
        return None
    return (current_id, text, options)


def _phase_tooltip(text):
    return u"Click: next phase.\nRight-click: previous phase."


def _phase_caret_tooltip(text):
    return u"Pick a phase"


def make_phase_switcher():
    """Phase badge: name cycles, caret opens the phase list."""
    return DropdownSwitcher(
        _HUD_ID,
        options_provider=_phase_options,
        on_select=_set_view_phase,
        tooltip_provider=_phase_tooltip,
        caret_tooltip_provider=_phase_caret_tooltip,
        on_left_click=lambda ua: _shift_view_phase(ua, 1),
        on_right_click=lambda ua: _shift_view_phase(ua, -1))


def find_phase_hud_driver(document):
    return find_hud_item(document, _HUD_ID)


def stop_phase_hud_driver(document):
    return remove_hud_item(document, _HUD_ID)


def start_phase_hud_driver(uiapp, document, logger=None):
    item = find_hud_item(document, _HUD_ID)
    if item is not None:
        item.refresh()
        return item
    host = get_hud_host(uiapp, document, logger=logger)
    return host.add_item(make_phase_switcher())


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

    host = find_hud_host(document)
    lines.append('host_running: {}'.format(host is not None))
    if host is not None:
        add('host_items', lambda: ', '.join(host.item_ids()) or '(none)')
        lines.append('window_built: {}'.format(host._window is not None))
        lines.append('window_visible_flag: {}'.format(host._visible))
        if host._window is not None:
            add('window_pos', lambda: '{}, {} ({}x{})'.format(
                host._window.Left, host._window.Top,
                host._window.ActualWidth, host._window.ActualHeight))
            add('window_is_visible', lambda: host._window.IsVisible)
        lines.append('owner_hwnd: {}'.format(host._owner_hwnd_int))
        add('owner_rect', lambda: str(window_rect(host._owner_hwnd_int)))
        add('theme_dark', revit_theme_is_dark)
        lines.append('last_error: {}'.format(host._last_error))
    item = find_hud_item(document, _HUD_ID)
    lines.append('phase_item_registered: {}'.format(item is not None))
    if item is not None and item.root is not None:
        add('item_visibility', lambda: str(item.root.Visibility))
        add('item_text', lambda: item._label.Text if item._label else None)
    return '\n'.join(lines)
