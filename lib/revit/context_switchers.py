# -*- coding: utf-8 -*-
"""Stock in-view context switchers and the default View HUD bundle.

Builds on revit.view_hud (ViewHudHost + TextSwitcher / DropdownSwitcher).
Which chips appear, bar placement, and auto-start come from
revit.view_hud_config. Available chips:

- Phase (name click cycles; caret picks a phase — revit.phase_label_wpf)
- Active workset (dropdown; hidden in non-workshared models)
- Phase filter (Show All / Show Previous + New / …)
- Detail level (Coarse / Medium / Fine)
- Visual style (Wireframe, Hidden Line, Shaded, …)
- Scope box (dropdown; hidden when the view has no Scope Box parameter)
- Section box (on/off; hidden on non-3D views)

start_view_hud / stop_view_hud / is_view_hud_running manage the bundle as
one toggle for the ribbon button. A ⚙ badge on the right of the bar opens
pyRevit CommandSwitchWindow chip toggles. Each switcher hides itself when
not applicable, so the bar only shows what matters in the current view.

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
    BuiltInCategory,
    BuiltInParameter,
    DisplayStyle,
    ElementId,
    FilteredElementCollector,
    FilteredWorksetCollector,
    PhaseFilter,
    Transaction,
    View3D,
    ViewDetailLevel,
    WorksetKind,
)

from revit.compat import get_element_id_value, make_element_id
from revit.phase_label_wpf import (
    make_phase_switcher,
    collect_diagnostics as _phase_diagnostics,
)
from collections import OrderedDict
from System.Collections.Generic import List

from revit.view_hud import (
    DropdownSwitcher,
    TextSwitcher,
    get_hud_host,
    find_hud_host,
)
from revit.view_hud_config import (
    CHIP_DETAIL_LEVEL,
    CHIP_DISPLAY_STYLE,
    CHIP_LABELS,
    CHIP_ORDER,
    CHIP_PHASE,
    CHIP_PHASE_FILTER,
    CHIP_SCOPE_BOX,
    CHIP_SECTION_BOX,
    CHIP_WORKSET,
    enabled_chip_ids,
    hud_style_from_config,
    load_hud_config,
    save_hud_config,
    ANCHORS,
    ANCHOR_LABELS,
    DEFAULT_ANCHOR,
)
from System.Windows.Input import Cursors

PHASE_ID = CHIP_PHASE
WORKSET_ID = CHIP_WORKSET
PHASE_FILTER_ID = CHIP_PHASE_FILTER
DETAIL_LEVEL_ID = CHIP_DETAIL_LEVEL
DISPLAY_STYLE_ID = CHIP_DISPLAY_STYLE
SCOPE_BOX_ID = CHIP_SCOPE_BOX
SECTION_BOX_ID = CHIP_SECTION_BOX
CONFIG_ID = 'hud_config'
_SAVE = 'Save'
_AUTO_START = 'Auto-start'
_MAX_CHARS = 40
_ANCHOR_HIGHLIGHT = {'background': '0xFFFFBB00'}
_SCOPE_NONE_LABEL = 'None'
_SCOPE_PARAM_NAMES = ('Scope Box', 'Områdesram')
_SCOPE_CAT_NAMES = ('OST_VolumeOfInterest', 'OST_ScopeBoxes')
_SECTION_ON = 'Section on'
_SECTION_OFF = 'Section off'

_DISPLAY_STYLES = (
    ('Wireframe', 'Wireframe'),
    ('HLR', 'Hidden Line'),
    ('Shading', 'Shaded'),
    ('ShadingWithEdges', 'Shaded with Edges'),
    ('ConsistentColors', 'Consistent Colors'),
    ('Realistic', 'Realistic'),
    ('RealisticWithEdges', 'Realistic with Edges'),
)

_DETAIL_LEVELS = (
    ('Coarse', 'Coarse'),
    ('Medium', 'Medium'),
    ('Fine', 'Fine'),
)


def _truncate(name):
    if name and len(name) > _MAX_CHARS:
        return name[:_MAX_CHARS - 1] + u'…'
    return name


def _enum_key(value):
    try:
        return value.ToString()
    except Exception:
        return str(value)


def _commit(doc, title, action):
    t = Transaction(doc, title)
    t.Start()
    try:
        action()
        t.Commit()
        return True
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        return None


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
        return _commit(
            doc, 'Switch active workset',
            lambda: table.SetActiveWorksetId(target.Id))
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


# ---------------------------------------------------------- phase filter

def _phase_filter_options(document, view):
    try:
        param = view.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER)
        if param is None:
            return None
        current_id = get_element_id_value(param.AsElementId())
        options = []
        for pf in FilteredElementCollector(document).OfClass(PhaseFilter):
            options.append((get_element_id_value(pf.Id), _truncate(pf.Name)))
        if not options:
            return None
        options.sort(key=lambda kv: kv[1].lower())
        current_name = None
        for key, name in options:
            if key == current_id:
                current_name = name
                break
        if current_name is None:
            current_id, current_name = options[0]
        return (current_id, current_name, options)
    except Exception:
        return None


def _set_phase_filter(uiapp, filter_id_value):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    view = uidoc.ActiveView
    if view is None:
        return None
    param = view.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER)
    if param is None or param.IsReadOnly:
        return None
    new_id = make_element_id(filter_id_value)
    return _commit(
        uidoc.Document, 'Switch phase filter',
        lambda: param.Set(new_id))


def _phase_filter_tooltip(text):
    return u"Phase filter: {}\nClick to pick a filter.".format(text)


def make_phase_filter_switcher():
    return DropdownSwitcher(
        PHASE_FILTER_ID,
        options_provider=_phase_filter_options,
        on_select=_set_phase_filter,
        tooltip_provider=_phase_filter_tooltip)


# ---------------------------------------------------------- detail level

def _detail_level_options(document, view):
    try:
        current = view.DetailLevel
    except Exception:
        return None
    param = view.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL)
    if param is None:
        return None
    options = []
    current_key = None
    current_label = None
    for attr, label in _DETAIL_LEVELS:
        try:
            value = getattr(ViewDetailLevel, attr)
        except Exception:
            continue
        key = _enum_key(value)
        options.append((key, label))
        if value == current or key == _enum_key(current):
            current_key = key
            current_label = label
    if not options:
        return None
    if param.IsReadOnly:
        if current_key is None:
            return None
        return (current_key, current_label, [(current_key, current_label)])
    if current_key is None:
        current_key, current_label = options[0]
    return (current_key, current_label, options)


def _set_detail_level(uiapp, key):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    view = uidoc.ActiveView
    if view is None:
        return None
    try:
        value = getattr(ViewDetailLevel, key)
    except Exception:
        return None
    def _apply():
        view.DetailLevel = value

    return _commit(uidoc.Document, 'Switch detail level', _apply)


def _detail_level_tooltip(text):
    return u"Detail level: {}\nClick to pick Coarse, Medium, or Fine.".format(
        text)


def make_detail_level_switcher():
    return DropdownSwitcher(
        DETAIL_LEVEL_ID,
        options_provider=_detail_level_options,
        on_select=_set_detail_level,
        tooltip_provider=_detail_level_tooltip)


# --------------------------------------------------------- visual style

def _display_style_options(document, view):
    try:
        current = view.DisplayStyle
    except Exception:
        return None
    options = []
    current_key = None
    current_label = None
    for attr, label in _DISPLAY_STYLES:
        try:
            value = getattr(DisplayStyle, attr)
        except Exception:
            continue
        key = attr
        options.append((key, label))
        if value == current or _enum_key(value) == _enum_key(current):
            current_key = key
            current_label = label
    if not options:
        return None
    if current_key is None:
        current_key, current_label = options[0]
    return (current_key, current_label, options)


def _set_display_style(uiapp, key):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    view = uidoc.ActiveView
    if view is None:
        return None
    try:
        value = getattr(DisplayStyle, key)
    except Exception:
        return None
    def _apply():
        view.DisplayStyle = value

    return _commit(uidoc.Document, 'Switch visual style', _apply)


def _display_style_tooltip(text):
    return u"Visual style: {}\nClick to pick a display style.".format(text)


def make_display_style_switcher():
    return DropdownSwitcher(
        DISPLAY_STYLE_ID,
        options_provider=_display_style_options,
        on_select=_set_display_style,
        tooltip_provider=_display_style_tooltip)


# ------------------------------------------------------------ scope box

def _invalid_id_value():
    return get_element_id_value(ElementId.InvalidElementId)


def _scope_box_param(view):
    if view is None:
        return None
    try:
        param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST)
        if param is not None:
            return param
    except Exception:
        pass
    for name in _SCOPE_PARAM_NAMES:
        try:
            param = view.LookupParameter(name)
            if param is not None:
                return param
        except Exception:
            continue
    return None


def _scope_box_element_id(view):
    """Assigned scope box ElementId, or None when unset / N/A."""
    param = _scope_box_param(view)
    if param is None:
        return None
    try:
        eid = param.AsElementId()
    except Exception:
        return None
    if eid is None:
        return None
    try:
        if eid == ElementId.InvalidElementId:
            return None
    except Exception:
        pass
    if get_element_id_value(eid) < 0:
        return None
    return eid


def _collect_scope_boxes(document):
    seen = {}
    for cat_name in _SCOPE_CAT_NAMES:
        try:
            cat = getattr(BuiltInCategory, cat_name)
        except Exception:
            continue
        try:
            for box in (
                    FilteredElementCollector(document)
                    .OfCategory(cat)
                    .WhereElementIsNotElementType()):
                seen[get_element_id_value(box.Id)] = box
        except Exception:
            continue
    boxes = list(seen.values())
    boxes.sort(key=lambda e: (e.Name or u'').lower())
    return boxes


def _scope_box_options(document, view):
    param = _scope_box_param(view)
    if param is None:
        return None
    none_key = _invalid_id_value()
    current_id = none_key
    current_name = _SCOPE_NONE_LABEL
    assigned = _scope_box_element_id(view)
    if assigned is not None:
        current_id = get_element_id_value(assigned)
        try:
            elem = document.GetElement(assigned)
            if elem is not None and elem.Name:
                current_name = _truncate(elem.Name)
        except Exception:
            current_name = u'Scope box'
    options = [(none_key, _SCOPE_NONE_LABEL)]
    for box in _collect_scope_boxes(document):
        options.append((get_element_id_value(box.Id), _truncate(box.Name)))
    if param.IsReadOnly:
        return (current_id, current_name, [(current_id, current_name)])
    return (current_id, current_name, options)


def _set_scope_box(uiapp, box_id_value):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    view = uidoc.ActiveView
    param = _scope_box_param(view)
    if param is None or param.IsReadOnly:
        return None
    try:
        box_id_value = int(box_id_value)
    except Exception:
        return None
    none_key = _invalid_id_value()
    if box_id_value == none_key or box_id_value < 0:
        new_id = ElementId.InvalidElementId
    else:
        new_id = make_element_id(box_id_value)
    return _commit(
        uidoc.Document, 'Switch scope box',
        lambda: param.Set(new_id))


def _scope_box_tooltip(text):
    return u"Scope box: {}\nClick to pick a scope box.".format(text)


def make_scope_box_switcher():
    return DropdownSwitcher(
        SCOPE_BOX_ID,
        options_provider=_scope_box_options,
        on_select=_set_scope_box,
        tooltip_provider=_scope_box_tooltip)


# ---------------------------------------------------------- section box

def _as_view3d(view):
    if view is None:
        return None
    try:
        if isinstance(view, View3D):
            return view
    except Exception:
        pass
    return None


def _section_box_text(document, view):
    v3d = _as_view3d(view)
    if v3d is None:
        return None
    try:
        if v3d.IsSectionBoxActive:
            return _SECTION_ON
        return _SECTION_OFF
    except Exception:
        return None


def _highlight_assigned_scope_box(uiapp):
    """Select the view's scope box so locked section-box clicks point at it."""
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    eid = _scope_box_element_id(uidoc.ActiveView)
    if eid is None:
        return None
    ids = List[ElementId]()
    ids.Add(eid)
    try:
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        return None
    return True


def _toggle_section_box(uiapp):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    view = uidoc.ActiveView
    if _scope_box_element_id(view) is not None:
        return None
    v3d = _as_view3d(view)
    if v3d is None:
        return None
    try:
        new_value = not bool(v3d.IsSectionBoxActive)
    except Exception:
        return None

    def _apply():
        v3d.IsSectionBoxActive = new_value

    return _commit(uidoc.Document, 'Toggle section box', _apply)


def _section_box_tooltip(text):
    if text == _SECTION_ON:
        return u"Section box is on.\nClick to turn off."
    return u"Section box is off.\nClick to turn on."


class _SectionBoxSwitcher(TextSwitcher):
    """Section-box on/off. Locked when the view has a scope box assigned."""

    def __init__(self):
        TextSwitcher.__init__(
            self,
            SECTION_BOX_ID,
            text_provider=_section_box_text,
            tooltip_provider=_section_box_tooltip,
            on_left_click=_toggle_section_box)
        self._locked = False

    def sync(self, document, view):
        text = TextSwitcher.sync(self, document, view)
        locked = text is not None and _scope_box_element_id(view) is not None
        self._locked = locked
        if text is None:
            return None
        display = (text + u' 🔒') if locked else text
        try:
            if self._text_block is not None:
                self._text_block.Text = display
            self._last_text = text
            if locked:
                self.root.Cursor = Cursors.Arrow
                self.root.ToolTip = (
                    u"Section box is locked while a scope box is assigned.\n"
                    u"Click to highlight the scope box.")
            else:
                self.root.Cursor = Cursors.Hand
                if self._tooltip_provider is not None:
                    self.root.ToolTip = self._tooltip_provider(text)
        except Exception:
            pass
        return (text, locked)

    def _on_mouse_left(self, sender, args):
        if self._locked:
            try:
                args.Handled = True
            except Exception:
                pass
            host = self.host
            item = host.find_item(SCOPE_BOX_ID) if host is not None else None
            if item is not None:
                item.highlight_briefly()
            # Direct UI-thread select — ExternalEvent would host.refresh()
            # and wipe the gold flash.
            try:
                uiapp = host._uiapp if host is not None else None
                if uiapp is not None:
                    _highlight_assigned_scope_box(uiapp)
            except Exception:
                pass
            return
        TextSwitcher._on_mouse_left(self, sender, args)


def make_section_box_switcher():
    return _SectionBoxSwitcher()


_FACTORIES = {
    CHIP_PHASE: make_phase_switcher,
    CHIP_WORKSET: make_workset_switcher,
    CHIP_PHASE_FILTER: make_phase_filter_switcher,
    CHIP_DETAIL_LEVEL: make_detail_level_switcher,
    CHIP_DISPLAY_STYLE: make_display_style_switcher,
    CHIP_SCOPE_BOX: make_scope_box_switcher,
    CHIP_SECTION_BOX: make_section_box_switcher,
}


# -------------------------------------------------------------- HUD config

def _anchor_from_switches(new_switches, current_anchor):
    """Pick a placement from the three position toggles."""
    on_keys = [
        key for key in ANCHORS
        if new_switches.get(ANCHOR_LABELS[key])]
    if len(on_keys) == 1:
        return on_keys[0]
    if not on_keys:
        return current_anchor
    for key in on_keys:
        if key != current_anchor:
            return key
    return current_anchor


def _prompt_and_save_hud_config():
    """pyRevit CommandSwitchWindow: placement, chip toggles, Save."""
    from pyrevit import forms

    cfg = load_hud_config()
    chips = cfg.get('chips') or {}
    current_anchor = cfg.get('anchor', DEFAULT_ANCHOR)
    if current_anchor not in ANCHOR_LABELS:
        current_anchor = DEFAULT_ANCHOR
    switches = OrderedDict()
    for key in ANCHORS:
        switches[ANCHOR_LABELS[key]] = (key == current_anchor)
    for chip_id in CHIP_ORDER:
        switches[CHIP_LABELS[chip_id]] = bool(chips.get(chip_id))
    switches[_AUTO_START] = bool(cfg.get('auto_start'))
    cfgs = {ANCHOR_LABELS[current_anchor]: _ANCHOR_HIGHLIGHT}
    result = forms.CommandSwitchWindow.show(
        [_SAVE],
        switches=switches,
        message='HUD settings',
        config=cfgs,
        recognize_access_key=False,
        height=720)
    if not result:
        return False
    option, new_switches = result
    if option != _SAVE:
        return False
    updated = {}
    for chip_id in CHIP_ORDER:
        updated[chip_id] = bool(new_switches.get(CHIP_LABELS[chip_id], False))
    cfg['chips'] = updated
    cfg['auto_start'] = bool(new_switches.get(_AUTO_START, False))
    cfg['anchor'] = _anchor_from_switches(new_switches, current_anchor)
    save_hud_config(cfg)
    return True


class _ConfigSwitcher(TextSwitcher):
    """Always-visible gear on the right; opens CommandSwitchWindow."""

    def _on_mouse_left(self, sender, args):
        if self.host is not None:
            self.host.close_item_popups()
        if not _prompt_and_save_hud_config():
            return

        def _rebuild(uiapp):
            uidoc = uiapp.ActiveUIDocument
            if uidoc is None:
                return
            apply_hud_chip_settings(uiapp, uidoc.Document)

        self.run_in_api_context(_rebuild)


def make_config_switcher():
    return _ConfigSwitcher(
        CONFIG_ID,
        text_provider=lambda doc, view: u'⚙',
        tooltip_provider=lambda text: u'HUD settings',
        on_left_click=lambda ua: None)


def _sync_host_chips(host):
    """Add/remove user chips; keep config badge last. No host teardown."""
    wanted = enabled_chip_ids()
    for item_id in list(host.item_ids()):
        if item_id != CONFIG_ID and item_id not in wanted:
            host.remove_item(item_id, _teardown_if_empty=False, refresh=False)
    for idx, chip_id in enumerate(wanted):
        if host.find_item(chip_id) is None:
            factory = _FACTORIES.get(chip_id)
            if factory is not None:
                host.add_item(factory(), index=idx, refresh=False)
    host.remove_item(CONFIG_ID, _teardown_if_empty=False, refresh=False)
    host.add_item(make_config_switcher(), refresh=False)
    host.refresh()
    return host


def apply_hud_chip_settings(uiapp, document, logger=None):
    """Rebuild chips and placement on a running HUD from saved settings."""
    host = find_hud_host(document)
    cfg = load_hud_config()
    if host is None:
        return start_view_hud(uiapp, document, logger=logger)
    if host._style is not None:
        host._style.anchor = cfg.get('anchor', DEFAULT_ANCHOR)
    return _sync_host_chips(host)


# ------------------------------------------------------------ the bundle

def is_view_hud_running(document):
    host = find_hud_host(document)
    return (host is not None
            and bool(host.item_ids())
            and not host.is_paused())


def start_view_hud(uiapp, document, logger=None):
    """Show the HUD bar using saved chip / placement settings."""
    cfg = load_hud_config()
    existing = find_hud_host(document)
    if existing is not None and existing.is_paused() and existing.item_ids():
        existing.resume()
        return existing
    if existing is not None:
        existing.stop()
    style = hud_style_from_config(cfg)
    host = get_hud_host(uiapp, document, style=style, logger=logger)
    for chip_id in enabled_chip_ids(cfg):
        factory = _FACTORIES.get(chip_id)
        if factory is not None:
            host.add_item(factory(), refresh=False)
    host.add_item(make_config_switcher(), refresh=False)
    host.refresh()
    return host


def stop_view_hud(document):
    host = find_hud_host(document)
    if host is None:
        return False
    host.pause()
    return True


def collect_diagnostics(uiapp, document):
    """Phase diagnostics plus the other chip provider states."""
    lines = [_phase_diagnostics(uiapp, document)]
    uidoc = uiapp.ActiveUIDocument
    view = uidoc.ActiveView if uidoc is not None else None
    cfg = load_hud_config()

    def add(key, fn):
        try:
            lines.append(u'{}: {}'.format(key, fn()))
        except Exception as ex:
            lines.append(u'{}: ERROR {}'.format(key, ex))

    add('hud_anchor', lambda: cfg.get('anchor'))
    add('hud_idle_opacity', lambda: cfg.get('idle_opacity'))
    add('hud_auto_start', lambda: cfg.get('auto_start'))
    add('hud_chips', lambda: ', '.join(enabled_chip_ids(cfg)) or '(none)')
    add('is_workshared', lambda: document.IsWorkshared)
    lines.append(u'workset_text: {}'.format(_workset_text(document, view)))
    add('user_worksets', lambda: len(_user_worksets(document)))
    if view is not None:
        add('phase_filter', lambda: _phase_filter_options(document, view))
        add('detail_level', lambda: _detail_level_options(document, view))
        add('display_style', lambda: _display_style_options(document, view))
        add('scope_box', lambda: _scope_box_options(document, view))
        add('section_box', lambda: _section_box_text(document, view))
    return u'\n'.join(lines)
