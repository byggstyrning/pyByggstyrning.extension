# -*- coding: utf-8 -*-
"""Stock in-view context switchers and the default View HUD bundle.

Builds on revit.view_hud (ViewHudHost + TextSwitcher / DropdownSwitcher).
The default bar shows, left to right:

- Phase (clickable: cycles the view's Phase — see revit.phase_label_wpf)
- Active workset (dropdown: single-select the document's active workset;
  hidden in non-workshared models)

start_view_hud / stop_view_hud / is_view_hud_running manage the bundle as
one toggle for the ribbon button. Each switcher hides itself when not
applicable, so the bar only shows what matters in the current model.

An active-design-option switcher was considered but left out: the Revit
API has no setter for the active design option (only
DesignOption.GetActiveDesignOptionId), and the only known workaround is a
fragile, out-of-process UI-automation hack on the status-bar combobox
(Jeremy Tammik's DesignOptionModifier). Add a HudItem here if that ever
becomes worthwhile.
"""

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    FilteredWorksetCollector,
    Transaction,
    WorksetKind,
)

from revit.phase_label_wpf import (
    make_phase_switcher,
    collect_diagnostics as _phase_diagnostics,
)
from revit.view_hud import (
    DropdownSwitcher,
    get_hud_host,
    find_hud_item,
    remove_hud_item,
)

PHASE_ID = 'phase'
WORKSET_ID = 'workset'
_DEFAULT_IDS = (PHASE_ID, WORKSET_ID)
_MAX_CHARS = 40


def _truncate(name):
    if name and len(name) > _MAX_CHARS:
        return name[:_MAX_CHARS - 1] + u'…'
    return name


# ---------------------------------------------------------------- workset

def _workset_text(document, view):
    """Active workset name, or None in non-workshared models."""
    try:
        if not document.IsWorkshared:
            return None
        table = document.GetWorksetTable()
        workset = table.GetWorkset(table.GetActiveWorksetId())
        if workset is None:
            return None
        return _truncate(workset.Name)
    except Exception:
        return None


def _user_worksets(document):
    try:
        return list(
            FilteredWorksetCollector(document)
            .OfKind(WorksetKind.UserWorkset))
    except Exception:
        return []


def _workset_options(document, view):
    """(active id, active name, [(id, name), ...]) or None if N/A."""
    try:
        if not document.IsWorkshared:
            return None
        table = document.GetWorksetTable()
        active_id = table.GetActiveWorksetId().IntegerValue
        active = table.GetWorkset(table.GetActiveWorksetId())
        active_name = _truncate(active.Name) if active is not None else None
        options = []
        for w in _user_worksets(document):
            options.append((w.Id.IntegerValue, _truncate(w.Name)))
        if not options:
            return None
        options.sort(key=lambda kv: kv[1].lower())
        if active_name is None:
            active_id, active_name = options[0]
        return (active_id, active_name, options)
    except Exception:
        return None


def _set_active_workset(uiapp, workset_id_value):
    """Set the document's active workset by workset id value.

    Runs inside a Revit API context. SetActiveWorksetId normally needs no
    transaction (the active workset is session state, not a model change);
    a transactional retry covers hosts that disagree.
    """
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    doc = uidoc.Document
    try:
        if not doc.IsWorkshared:
            return None
    except Exception:
        return None
    target = None
    for w in _user_worksets(doc):
        if w.Id.IntegerValue == workset_id_value:
            target = w
            break
    if target is None:
        return None
    table = doc.GetWorksetTable()
    try:
        table.SetActiveWorksetId(target.Id)
    except Exception:
        t = Transaction(doc, 'Switch active workset')
        t.Start()
        try:
            table.SetActiveWorksetId(target.Id)
            t.Commit()
        except Exception:
            try:
                t.RollBack()
            except Exception:
                pass
            return None
    return True


def _workset_tooltip(text):
    return (u"Active workset: {}\nNew elements are created here.\n"
            u"Click to pick a workset.".format(text))


def make_workset_switcher():
    return DropdownSwitcher(
        WORKSET_ID,
        options_provider=_workset_options,
        on_select=_set_active_workset,
        tooltip_provider=_workset_tooltip)


# ------------------------------------------------------------ the bundle

def is_view_hud_running(document):
    for item_id in _DEFAULT_IDS:
        if find_hud_item(document, item_id) is not None:
            return True
    return False


def start_view_hud(uiapp, document, logger=None):
    """Show the default switcher bar (phase, workset)."""
    host = get_hud_host(uiapp, document, logger=logger)
    if host.find_item(PHASE_ID) is None:
        host.add_item(make_phase_switcher())
    if host.find_item(WORKSET_ID) is None:
        host.add_item(make_workset_switcher())
    host.refresh()
    return host


def stop_view_hud(document):
    removed = False
    for item_id in _DEFAULT_IDS:
        removed = remove_hud_item(document, item_id) or removed
    return removed


def collect_diagnostics(uiapp, document):
    """Phase diagnostics plus the workset provider states."""
    lines = [_phase_diagnostics(uiapp, document)]
    uidoc = uiapp.ActiveUIDocument
    view = uidoc.ActiveView if uidoc is not None else None

    def add(key, fn):
        try:
            lines.append(u'{}: {}'.format(key, fn()))
        except Exception as ex:
            lines.append(u'{}: ERROR {}'.format(key, ex))

    add('is_workshared', lambda: document.IsWorkshared)
    lines.append(u'workset_text: {}'.format(_workset_text(document, view)))
    add('user_worksets', lambda: len(_user_worksets(document)))
    return u'\n'.join(lines)
