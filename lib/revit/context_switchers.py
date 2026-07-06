# -*- coding: utf-8 -*-
"""Stock in-view context switchers and the default View HUD bundle.

Builds on revit.view_hud (ViewHudHost + TextSwitcher). The default bar
shows, left to right:

- Phase (clickable: cycles the view's Phase — see revit.phase_label_wpf)
- Active workset (clickable: cycles the document's active workset;
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

from System.Windows.Input import Cursors

from revit.phase_label_wpf import (
    make_phase_switcher,
    collect_diagnostics as _phase_diagnostics,
)
from revit.view_hud import (
    TextSwitcher,
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


def _shift_active_workset(uiapp, step):
    """Set the document's active workset to the next/previous user workset.

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
    worksets = _user_worksets(doc)
    if len(worksets) < 2:
        return None
    table = doc.GetWorksetTable()
    values = [w.Id.IntegerValue for w in worksets]
    try:
        idx = values.index(table.GetActiveWorksetId().IntegerValue)
    except ValueError:
        idx = 0
    new_id = worksets[(idx + step) % len(worksets)].Id
    try:
        table.SetActiveWorksetId(new_id)
    except Exception:
        t = Transaction(doc, 'Switch active workset')
        t.Start()
        try:
            table.SetActiveWorksetId(new_id)
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
            u"Click: next workset. Right-click: previous.".format(text))


class _WorksetSwitcher(TextSwitcher):
    """Workset badge that drops the click affordance when cycling is moot.

    With a single user workset the badge stays as an indicator, but the
    hand cursor goes away and the tooltip stops promising a switch."""

    def sync(self, document, view):
        state = TextSwitcher.sync(self, document, view)
        if state is None or self.root is None:
            return state
        cyclable = len(_user_worksets(document)) > 1
        try:
            self.root.Cursor = Cursors.Hand if cyclable else None
            if not cyclable:
                self.root.ToolTip = (
                    u"Active workset: {}\n"
                    u"New elements are created here.".format(state))
        except Exception:
            pass
        return (state, cyclable)


def make_workset_switcher():
    return _WorksetSwitcher(
        WORKSET_ID,
        text_provider=_workset_text,
        tooltip_provider=_workset_tooltip,
        on_left_click=lambda ua: _shift_active_workset(ua, 1),
        on_right_click=lambda ua: _shift_active_workset(ua, -1))


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
