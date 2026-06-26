# -*- coding: utf-8 -*-
"""CDE door/element schedule - a modeless Revit cockpit over the CDE graph.

Joins CDE elements (by IFC GlobalId) to Revit elements, lists them in a
groupable/filterable table, lets the user add CDE/Revit parameter columns,
color elements in the active view by any value (temporary view graphics),
sync selection back into Revit when rows are picked, and vertically set values back to the CDE.

DB-first: vertical value setting writes to the CDE; Revit parameter sync is a
separate, opt-in action (later phase).
"""
__title__ = "CDE\nSchedule"
__author__ = "Byggstyrning AB"
__doc__ = "Open the CDE door/element schedule for the mapped project."

import os.path as op
import sys

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Action
from System.Threading import Thread, ThreadStart
from System.Windows import Thickness, Visibility, HorizontalAlignment, VerticalAlignment
from System.Windows.Threading import DispatcherPriority
from System.Windows.Data import (
    Binding, IValueConverter, CollectionViewSource, PropertyGroupDescription)
from System.Windows.Controls import (
    DataGridTextColumn, DataGridCheckBoxColumn, CheckBox, TextBox, Button,
    DataGridLength, DataGridLengthUnitType, StackPanel, TextBlock,
    Orientation, DataGridEditAction, DataGridEditingUnit)
from System.Windows.Media import SolidColorBrush, Color as WpfColor, FontFamily

from pyrevit import forms, revit, script, HOST_APP

# --- lib bootstrap --------------------------------------------------------
_pushbutton_dir = op.dirname(__file__)
_extension_dir = op.dirname(op.dirname(op.dirname(_pushbutton_dir)))
_lib_path = op.join(_extension_dir, "lib")
if _lib_path not in sys.path:
    sys.path.insert(0, _lib_path)

from styles import load_styles_to_window
from cde import storage, coloring, config as cde_config
import cde.service as cde_service
import cde.matching as cde_matching
import cde.viewmodels as cde_viewmodels
reload(cde_service)
reload(cde_matching)
reload(cde_viewmodels)
matching = cde_matching
from cde.auth import CDEAuthClient
from cde.service import CDEService, MockCDEService, param_def, MUTATION_SUCCESS_STATES
from cde.api import CDEPreconditionError, CDEConflictError
from cde.viewmodels import ElementRow, ScheduleViewModel
from cde.dfp_catalog import DFP_PSET_NAME
from cde.revit_events import ExternalEventRunner, ActiveViewWatcher

_DFP_MARKERS_OK = False
_DFP_MARKERS_IMPORT_ERROR = None
try:
    from revit.dfp_markers import (
        is_temporary_graphics_available as _dfp_tgm_available,
        register_dfp_session,
        is_dfp_toggle_active,
        set_dfp_toggle_active,
        start_or_get_dfp_driver,
        clean_dfp_session,
        soft_clean_dfp_session,
        refresh_dfp_session,
        ensure_temporary_graphics_handler as _ensure_dfp_tgm_handler,
        register_dfp_marker_click_callback,
        register_dfp_schedule_window,
        update_dfp_door_after_toggle,
        find_dfp_driver as _find_dfp_driver,
        hide_dfp_graphics_for_apply,
        set_dfp_apply_block,
    )
    from cde.dfp_markers import build_dfp_view_points, default_dfp_param_keys
    from cde.dfp_icons import has_authored_dfp_value as _has_authored_dfp_value
    from revit.compat import get_element_id_value as _get_element_id_value
    _DFP_MARKERS_OK = _dfp_tgm_available()
except Exception as _dfp_ex:
    _DFP_MARKERS_IMPORT_ERROR = str(_dfp_ex)
    register_dfp_session = None
    is_dfp_toggle_active = None
    set_dfp_toggle_active = None
    start_or_get_dfp_driver = None
    clean_dfp_session = None
    soft_clean_dfp_session = None
    refresh_dfp_session = None
    _ensure_dfp_tgm_handler = None
    register_dfp_marker_click_callback = None
    register_dfp_schedule_window = None
    update_dfp_door_after_toggle = None
    _find_dfp_driver = None
    hide_dfp_graphics_for_apply = None
    set_dfp_apply_block = None
    build_dfp_view_points = None
    default_dfp_param_keys = None
    _has_authored_dfp_value = None
    _get_element_id_value = None

logger = script.get_logger()

# Categories the schedule can drive (label -> IfcClass).
CATEGORIES = [("Doors", "IfcDoor"), ("Windows", "IfcWindow"), ("Walls", "IfcWall")]
# Group-by options (label -> ElementRow property path, None = ungrouped).
GROUP_OPTIONS = [("No grouping", None), ("Level", "Level"),
                 ("From room", "FromRoom"), ("From room no.", "FromRoomNumber"),
                 ("To room", "ToRoom"), ("To room no.", "ToRoomNumber")]
# DEBUG: temporarily disable the Revit-parameter scan to isolate the native crash.
_ENABLE_REVIT_PARAM_SCAN = False
# Fixed grid columns that support header grouping (path -> source tint).
FIXED_COLUMN_META = [
    ("GlobalId", "GlobalId", "revit"),
    ("Mark", "Mark", "revit"),
    ("Level", "Level", "revit"),
    ("FromRoomNumber", "From no.", "revit"),
    ("FromRoom", "From room", "revit"),
    ("ToRoomNumber", "To no.", "revit"),
    ("ToRoom", "To room", "revit"),
    ("MatchedInRevit", "In Revit", "revit"),
]
# Segoe MDL2 Assets glyphs for source / actions.
ICON_CDE = u"\uE753"
ICON_REVIT = u"\uE821"
ICON_GROUP = u"\uE8FD"
# IronPython: do not call DataGridLength via a stored type ref on self — use this.
_VALUE_COLUMN_WIDTH = DataGridLength(110.0, DataGridLengthUnitType.Pixel)

# Keep a module reference so the modeless window is not garbage-collected.
__window__ = None


class _ComboItem(object):
    def __init__(self, name, value):
        self.name = name
        self.value = value

    # DisplayMemberPath does not resolve attributes on plain IronPython objects;
    # WPF falls back to ToString(), so expose the label that way instead.
    def ToString(self):
        return self.name if self.name is not None else u""

    __str__ = ToString
    __repr__ = ToString


class _LegendItem(object):
    def __init__(self, label, brush):
        self.Label = label
        self.Brush = brush


class _PropertyRow(object):
    def __init__(self, key, value):
        self.Key = key
        self.Value = u"" if value is None else unicode(value)


def _source_icon(source):
    return ICON_CDE if source == "cde" else ICON_REVIT


class _ParamDisplayItem(object):
    """ComboBox row for parameter picker (SearchableComboBox display_name binding)."""
    _ICON_CDE = ICON_CDE
    _ICON_REVIT = ICON_REVIT

    def __init__(self, param_def, cde_brush, revit_brush):
        self.param_def = param_def
        self.display_name = param_def.label
        self.source = getattr(param_def, "source", "cde") or "cde"
        self.source_icon = self._ICON_CDE if self.source == "cde" else self._ICON_REVIT
        self.source_brush = cde_brush if self.source == "cde" else revit_brush

    def ToString(self):
        return self.display_name

    __str__ = ToString
    __repr__ = ToString


class _ColumnPickerItem(object):
    """Checkbox row inside the columns SearchableComboBox dropdown."""
    _ICON_CDE = ICON_CDE
    _ICON_REVIT = ICON_REVIT

    def __init__(self, param_def, is_checked, cde_brush, revit_brush, owner):
        self.param_def = param_def
        self.key = param_def.key
        self.display_name = param_def.label
        self.source = getattr(param_def, "source", "cde") or "cde"
        self.source_icon = self._ICON_CDE if self.source == "cde" else self._ICON_REVIT
        self.source_brush = cde_brush if self.source == "cde" else revit_brush
        self.is_checked = bool(is_checked)
        self._owner = owner

    def ToString(self):
        return self.display_name

    __str__ = ToString
    __repr__ = ToString


class CellValueConverter(IValueConverter):
    """Resolves a dynamic CDE/Revit cell value from a row + column key."""

    def __init__(self, owner=None):
        self._owner = owner

    def _row_for(self, value):
        if value is not None and hasattr(value, "get_cell"):
            return value
        if self._owner is not None:
            return self._owner._row_for_grid_key(value)
        return None

    def Convert(self, value, target_type, parameter, culture):
        try:
            row = self._row_for(value)
            if row is None:
                return ""
            raw = row.get_cell(parameter)
            if isinstance(raw, bool):
                return raw
            if raw in (0, 1, "0", "1", "True", "False", "true", "false"):
                lowered = unicode(raw).lower()
                if lowered in ("true", "1"):
                    return True
                if lowered in ("false", "0"):
                    return False
            return raw
        except Exception:
            return ""

    def ConvertBack(self, value, target_type, parameter, culture):
        return None


_BOOL_STRINGS = frozenset(["true", "false", "1", "0", "yes", "no"])


class GroupKeyConverter(IValueConverter):
    """Converts a string row-key to the ElementRow.GroupValue for CollectionView grouping."""

    def __init__(self, owner=None):
        self._owner = owner

    def Convert(self, value, target_type, parameter, culture):
        try:
            if self._owner is None:
                return u""
            row = self._owner._row_for_grid_key(value)
            if row is None:
                return u""
            return row.GroupValue
        except Exception:
            return u""

    def ConvertBack(self, value, target_type, parameter, culture):
        return None


def _int_types_for_isinstance():
    """IronPython-safe isinstance target for integral values."""
    try:
        return (int, long)
    except NameError:
        return (int,)


def _lower_text(value):
    try:
        if isinstance(value, str):
            return value.decode("utf-8", "replace").lower()
        return unicode(value).lower()
    except NameError:
        return str(value).lower()
    except UnicodeDecodeError:
        return str(value).lower()


def _infer_cde_value_type(sample_values):
    """Guess value_type from CDE element payloads (auto-discovered columns)."""
    samples = []
    for val in sample_values or []:
        if val is None or val == "":
            continue
        samples.append(val)
        if len(samples) >= 50:
            break
    if not samples:
        return "string"
    if all(isinstance(v, bool) for v in samples):
        return "bool"
    int_types = _int_types_for_isinstance()
    if all(isinstance(v, int_types) and v in (0, 1) for v in samples):
        return "bool"
    if all(_lower_text(v) in _BOOL_STRINGS for v in samples):
        return "bool"
    return "string"


class ScheduleWindow(forms.WPFWindow):

    def __init__(self):
        forms.WPFWindow.__init__(self, op.join(_pushbutton_dir, "ScheduleWindow.xaml"))
        load_styles_to_window(self)

        self.uiapp = HOST_APP.uiapp
        self.doc = revit.doc
        self.auth = CDEAuthClient()
        self.offline = False
        self.service = CDEService(self.auth)

        self.vm = ScheduleViewModel()
        self._row_by_grid_key = {}
        self._grid_items = None
        self._cell_converter = CellValueConverter(self)
        self._group_converter = GroupKeyConverter(self)
        self._refresh_grid_items_from_vm()
        self.doorGrid.ItemsSource = self._grid_items
        self._param_defs = []
        self._def_by_key = {}
        self._revit_param_defs = []
        self._revit_infos = {}
        self._all_param_items = []
        self._all_column_items = []
        # key -> DataGridColumn for the dynamic CDE property columns.
        self._value_column_map = {}
        self._column_meta = {}
        self._column_key_by_column = {}
        # IfcClass the auto-columns were initialized for (None until first load).
        self._columns_ifc_class = None
        self._columns_dropdown_open = False
        self._param_search_wired = False
        self._columns_search_wired = False
        self._suppress_group_changed = False
        self._refresh_token = 0
        self._fetch_running = False
        self._apply_in_flight = False
        self._pending_etag_persist = None
        self._pending_commit_ctx = None
        self._dfp_graphics_refresh_gid = None
        self._dfp_markers_were_active = False
        self._dfp_markers_need_manual_reset = False
        self._apply_revit_quarantine = False
        self._suppress_inline_staging = False
        self._pending_restore = None
        self._detail_fetch_token = 0
        self._node_by_gid = {}
        self._dfp_markers_active = False
        self._runner = ExternalEventRunner()
        self._watcher = ActiveViewWatcher(self.uiapp, self._on_active_view_changed)
        # Capture ALL module-level names used in any method; pyRevit may dispose the
        # script scope after __init__ returns (modeless window stays alive, scope dies).
        self._logger = logger
        self._matching = matching
        self._coloring = coloring
        self._Action = Action
        self._Thread = Thread
        self._ThreadStart = ThreadStart
        self._CollectionViewSource = CollectionViewSource
        self._PropertyGroupDescription = PropertyGroupDescription
        self._DataGridTextColumn = DataGridTextColumn
        self._Binding = Binding
        self._CheckBox = CheckBox
        self._Thickness = Thickness
        self._SolidColorBrush = SolidColorBrush
        self._WpfColor = WpfColor
        self._ComboItem = _ComboItem
        self._LegendItem = _LegendItem
        self._PropertyRow = _PropertyRow
        self._ParamDisplayItem = _ParamDisplayItem
        self._ColumnPickerItem = _ColumnPickerItem
        self._param_def = param_def
        self._ElementRow = ElementRow
        self._dfp_pset_name = DFP_PSET_NAME
        self._dfp_markers_ok = _DFP_MARKERS_OK
        self._dfp_markers_import_error = _DFP_MARKERS_IMPORT_ERROR
        self._build_dfp_view_points = build_dfp_view_points
        self._default_dfp_param_keys = default_dfp_param_keys
        self._register_dfp_session = register_dfp_session
        self._set_dfp_toggle_active = set_dfp_toggle_active
        self._start_or_get_dfp_driver = start_or_get_dfp_driver
        self._clean_dfp_session = clean_dfp_session
        self._soft_clean_dfp_session = soft_clean_dfp_session
        self._ensure_dfp_tgm_handler = _ensure_dfp_tgm_handler
        self._register_dfp_marker_click_callback = register_dfp_marker_click_callback
        self._register_dfp_schedule_window = register_dfp_schedule_window
        self._update_dfp_door_after_toggle = update_dfp_door_after_toggle
        self._find_dfp_driver = _find_dfp_driver
        self._hide_dfp_graphics_for_apply = hide_dfp_graphics_for_apply
        self._set_dfp_apply_block = set_dfp_apply_block
        self._has_authored_dfp_value = _has_authored_dfp_value
        self._get_element_id_value = _get_element_id_value
        self._value_column_width = _VALUE_COLUMN_WIDTH
        # Closures — module globals are unavailable after pyRevit disposes script scope.
        self._enable_revit_param_scan = _ENABLE_REVIT_PARAM_SCAN
        _int_types = _int_types_for_isinstance()
        _bool_strings = _BOOL_STRINGS
        _lower = _lower_text
        def _bound_infer_cde_value_type(sample_values):
            samples = []
            for val in sample_values or []:
                if val is None or val == "":
                    continue
                samples.append(val)
                if len(samples) >= 50:
                    break
            if not samples:
                return "string"
            if all(isinstance(v, bool) for v in samples):
                return "bool"
            if all(isinstance(v, _int_types) and v in (0, 1) for v in samples):
                return "bool"
            if all(_lower(v) in _bool_strings for v in samples):
                return "bool"
            return "string"
        self._infer_cde_value_type = _bound_infer_cde_value_type
        self._fixed_column_meta = FIXED_COLUMN_META
        self._DispatcherPriority = DispatcherPriority
        self._DataGridCheckBoxColumn = DataGridCheckBoxColumn
        self._StackPanel = StackPanel
        self._TextBlock = TextBlock
        self._Button = Button
        self._TextBox = TextBox
        self._Orientation = Orientation
        self._Visibility = Visibility
        self._HorizontalAlignment = HorizontalAlignment
        self._VerticalAlignment = VerticalAlignment
        self._FontFamily = FontFamily
        self._DataGridEditAction = DataGridEditAction
        self._DataGridEditingUnit = DataGridEditingUnit
        self._ICON_GROUP = ICON_GROUP
        self._ICON_CDE = ICON_CDE
        self._ICON_REVIT = ICON_REVIT
        self._mutation_success_states = (
            set(MUTATION_SUCCESS_STATES) | {"committed"})
        self._cde_brush = self._SolidColorBrush(self._WpfColor.FromRgb(0, 120, 212))
        self._revit_brush = self._SolidColorBrush(self._WpfColor.FromRgb(232, 108, 0))
        self._icon_font = self._FontFamily("Segoe MDL2 Assets")

        self._init_combos()
        self._style_fixed_column_headers()
        self.vm.on_pending_changed(self._update_pending_ui)
        self._install_dfp_marker_click_handler()
        self.mapping = storage.load_mapping(self.doc)
        if self.mapping:
            stored = (self.mapping.get("base_url") or "").rstrip("/")
            cfg = cde_config.get_base_url()
            effective = stored or cfg
            auth_url = (self.auth.base_url or "").rstrip("/")
            logger.debug(
                "CDE Schedule URLs — stored: '{}', config: '{}', "
                "effective: '{}', auth client: '{}'".format(
                    stored or "(empty)", cfg, effective, auth_url))
        else:
            self._set_status("Model is not mapped. Run 'CDE Login' first "
                             "(or enable offline demo via Login).")
        self._update_pending_ui()

    def _update_pending_ui(self):
        count = self.vm.pending_count()
        try:
            self.pendingCountText.Text = "Pending: {}".format(count)
            enabled = (
                count > 0
                and not self._apply_in_flight
                and not self._dfp_markers_active)
            self.applyPendingButton.IsEnabled = enabled
            self.discardPendingButton.IsEnabled = enabled
        except Exception as ex:
            self._logger.debug("CDE: pending UI update failed: {}".format(ex))

    def _set_apply_in_flight(self, in_flight):
        self._apply_in_flight = bool(in_flight)
        try:
            can_apply = (
                not in_flight
                and self.vm.pending_count() > 0
                and not self._dfp_markers_active)
            self.applyPendingButton.IsEnabled = can_apply
            self.discardPendingButton.IsEnabled = not in_flight and self.vm.pending_count() > 0
            self.setValueButton.IsEnabled = not in_flight
        except Exception:
            pass

    def _project_revision_ids(self):
        if not self.mapping and not self.offline:
            return None, None
        project_id = self.mapping["project_id"] if self.mapping else "demo-1"
        revision_id = self.mapping["revision_id"] if self.mapping else "rev-2"
        return project_id, revision_id

    def _file_version_id(self):
        """CDE file version id behind the mapped revision (for mutation scoping)."""
        if not self.mapping:
            return None
        return self.mapping.get("file_version_id") or None

    def _file_id(self):
        """CDE file id (== backend model_id) for computing the revision ETag."""
        if not self.mapping:
            return None
        return self.mapping.get("file_id") or None

    def _cache_revision_etag(self, etag):
        """Memory-only etag update (safe from background threads)."""
        if etag and self.mapping is not None:
            self.mapping["revision_etag"] = etag

    def _persist_revision_etag(self, etag):
        """Write etag to Extensible Storage on the Revit API thread."""
        if not etag:
            return
        self._cache_revision_etag(etag)
        if self._apply_in_flight or self._apply_revit_quarantine:
            self._pending_etag_persist = etag
            return

        def task(uiapp):
            try:
                doc = uiapp.ActiveUIDocument.Document
                storage.update_revision_etag(doc, etag)
            except Exception as ex:
                self._logger.debug("CDE: etag persist failed: {}".format(ex))

        self._runner.run(task)

    def _flush_deferred_etag_persist(self):
        """Persist etag deferred while Apply was in flight."""
        etag = self._pending_etag_persist
        if not etag or self._apply_in_flight:
            return
        self._pending_etag_persist = None

        def task(uiapp):
            try:
                doc = uiapp.ActiveUIDocument.Document
                storage.update_revision_etag(doc, etag)
            except Exception as ex:
                self._logger.debug("CDE: etag persist failed: {}".format(ex))

        self._runner.run(task)

    def _get_revision_etag(self, project_id, revision_id):
        """Return cached or fetched etag (HTTP only — no Revit API)."""
        cached = ""
        if self.mapping:
            cached = self.mapping.get("revision_etag") or ""
        if cached:
            return cached
        etag = self.service.fetch_revision_etag(project_id, revision_id)
        if etag:
            self._cache_revision_etag(etag)
        return etag

    def _save_revision_etag(self, etag):
        """Cache etag and persist to the active document on the Revit thread."""
        self._persist_revision_etag(etag)

    # --- setup ----------------------------------------------------------

    def _init_combos(self):
        self.categoryCombo.ItemsSource = [_ComboItem(l, c) for l, c in CATEGORIES]
        self.categoryCombo.SelectedIndex = 0
        self.groupCombo.ItemsSource = [_ComboItem(l, p) for l, p in GROUP_OPTIONS]
        self.groupCombo.SelectedIndex = 0

    def _make_grid_key(self, index, row):
        return "{}".format(index)

    def _row_for_grid_key(self, key):
        if key is None:
            return None
        if hasattr(key, "get_cell"):
            return key
        return self._row_by_grid_key.get(unicode(key))

    def _refresh_grid_items_from_vm(self):
        from System.Collections.ObjectModel import ObservableCollection
        from System import Object
        items = ObservableCollection[Object]()
        self._row_by_grid_key = {}
        for idx, row in enumerate(self.vm.rows):
            key = self._make_grid_key(idx, row)
            self._row_by_grid_key[key] = row
            items.Add(key)
        self._grid_items = items

    def _set_status(self, message):
        self.statusText.Text = message

    def on_door_grid_loaded(self, sender, args):
        """Remeasure rows once the grid is in the visual tree."""
        try:
            self.doorGrid.InvalidateMeasure()
            self.doorGrid.InvalidateArrange()
        except Exception as ex:
            self._logger.debug("CDE: door grid loaded layout failed: {}".format(ex))

    @property
    def ifc_class(self):
        item = self.categoryCombo.SelectedItem
        return item.value if item else "IfcDoor"

    # --- async helper ---------------------------------------------------

    def _run_async(self, work, on_done=None):
        def runner():
            result = None
            error = None
            try:
                result = work()
            except Exception as ex:
                error = ex
            if on_done is not None:
                try:
                    self.Dispatcher.BeginInvoke(
                        self._DispatcherPriority.Normal,
                        self._Action(lambda: on_done(result, error)))
                except Exception as inv_ex:
                    self._logger.error("CDE: UI callback failed: {}".format(inv_ex))
        thread = self._Thread(self._ThreadStart(runner))
        thread.IsBackground = True
        thread.Start()

    # --- refresh orchestration -----------------------------------------

    def on_refresh_click(self, sender, args):
        self.refresh()

    def refresh(self, preserve_pending=False):
        if not preserve_pending and self.vm.pending_count() > 0:
            # Non-modal gate: a modal dialog here pumps Revit's message loop
            # while ArrowEditor is active and crashes (0xe0434352). Require an
            # explicit non-modal Discard/Apply instead.
            count = self.vm.pending_count()
            self._set_status(
                "{} pending edit(s) — click Discard or Apply before "
                "refreshing.".format(count))
            return
        if preserve_pending:
            self._pending_restore = self.vm.collect_pending_changes()
        else:
            self._pending_restore = None
        self._refresh_token += 1
        token = self._refresh_token
        self.vm.ifc_class = self.ifc_class
        self.vm.active_view_only = bool(self.activeViewCheck.IsChecked)
        self._set_status("Reading Revit elements...")
        # Step 1: collect Revit data on the Revit API thread.
        self._runner.run(lambda uiapp: self._collect_revit(uiapp, token))

    def _collect_revit(self, uiapp, token):
        if token != self._refresh_token:
            return
        try:
            doc = uiapp.ActiveUIDocument.Document
            self.doc = doc
            view = doc.ActiveView if self.vm.active_view_only else None
            param_reader = None
            if self._enable_revit_param_scan:
                param_reader = self._matching.read_revit_parameters
            self._revit_infos = self._matching.collect_revit_infos(
                doc, self.vm.ifc_class, view,
                include_params=self._enable_revit_param_scan,
                param_reader=param_reader)
            if self._enable_revit_param_scan:
                self._revit_param_defs = self._matching.collect_revit_parameter_defs(
                    doc, self._revit_infos, self._param_def)
            else:
                self._revit_param_defs = []
        except Exception as ex:
            self._logger.error("CDE: collect Revit failed: {}".format(ex))
            self._revit_infos = {}
            self._revit_param_defs = []
        # BeginInvoke avoids Dispatcher.Invoke deadlock from ExternalEvent handler.
        self.Dispatcher.BeginInvoke(
            self._DispatcherPriority.Normal,
            self._Action(lambda: self._fetch_cde(token)))

    def _baseline_param_defs(self):
        """Fixed Revit column defs when CDE/scan yield no dynamic parameters."""
        result = []
        for path, label, source in self._fixed_column_meta:
            if path == "GlobalId":
                continue
            value_type = "bool" if path == "MatchedInRevit" else "string"
            result.append(self._param_def(
                path, label, value_type=value_type, group="Revit", source=source))
        return result

    def _fetch_cde(self, token):
        if token != self._refresh_token:
            return
        if self._fetch_running:
            return
        if not self.mapping and not self.offline:
            self._build_rows(([], []), None, token)
            return
        if not self.offline:
            if hasattr(self.auth, "reload_from_disk"):
                self.auth.reload_from_disk()
            if not self.auth.is_authenticated():
                self._build_rows(
                    ([], []),
                    "CDE session expired or not signed in. "
                    "Run CDE Login, then click Refresh (or close and reopen Schedule).",
                    token)
                return
        project_id = self.mapping["project_id"] if self.mapping else "demo-1"
        revision_id = self.mapping["revision_id"] if self.mapping else "rev-2"
        ifc_class = self.vm.ifc_class
        self._fetch_running = True
        self._set_status(
            "Fetching CDE elements...  [rev: {}]".format(revision_id or "(none)"))

        def work():
            elements = self.service.list_elements(project_id, revision_id, ifc_class)
            defs = self.service.get_parameter_defs(project_id, ifc_class)
            fetch_error = getattr(self.service, "last_fetch_error", None)
            if fetch_error and not elements:
                raise Exception(fetch_error)
            return elements, defs

        def on_fetch_done(payload, error):
            self._fetch_running = False
            self._build_rows(payload, error, token)
            if token != self._refresh_token:
                self._fetch_cde(self._refresh_token)

        self._run_async(work, on_fetch_done)

    def _build_rows(self, payload, error, token=None):
        if token is not None and token != self._refresh_token:
            return
        try:
            self._build_rows_inner(payload, error)
        except Exception as ex:
            import traceback
            self._logger.error("CDE: build_rows failed: {}\n{}".format(
                ex, traceback.format_exc()))
            self._set_status("Internal error building rows: {}".format(ex))

    def _build_rows_inner(self, payload, error):
        if error is not None:
            self._set_status("CDE fetch failed: {}".format(error))
            payload = ([], [])
        elements, defs = payload

        api_defs = list(defs or [])
        auto_defs = []
        if elements:
            seen = set()
            samples_by_key = {}
            dfp_prefix = self._dfp_pset_name + "."
            for e in elements:
                for key, val in e.values.items():
                    if not key.startswith(dfp_prefix):
                        continue
                    samples_by_key.setdefault(key, []).append(val)
                    if key not in seen:
                        seen.add(key)
                        parts = key.split(".", 1)
                        label = parts[1] if len(parts) == 2 else key
                        value_type = self._infer_cde_value_type(samples_by_key[key])
                        auto_defs.append(self._param_def(
                            key, label, value_type=value_type, source="cde"))
            auto_defs.sort(key=lambda d: d.key)
            label_counts = {}
            for d in auto_defs:
                label_counts[d.label] = label_counts.get(d.label, 0) + 1
            deduped = []
            for d in auto_defs:
                if label_counts[d.label] > 1:
                    parts = d.key.split(".", 1)
                    pset = parts[0] if len(parts) == 2 else d.key
                    deduped.append(self._param_def(
                        d.key, u"{} ({})".format(d.label, pset),
                        value_type=d.value_type, source="cde"))
                else:
                    deduped.append(d)
            auto_defs = deduped

        seen_keys = set(d.key for d in api_defs)
        defs = list(api_defs)
        for d in auto_defs:
            if d.key not in seen_keys:
                defs.append(d)
                seen_keys.add(d.key)
        for d in self._revit_param_defs or []:
            if d.key not in seen_keys:
                defs.append(d)
                seen_keys.add(d.key)
        for d in self._baseline_param_defs():
            if d.key not in seen_keys:
                defs.append(d)
                seen_keys.add(d.key)
        defs.sort(key=lambda d: (
            0 if getattr(d, "source", "cde") == "cde"
            and getattr(d, "group", "") == "Function" else 1,
            getattr(d, "source", "cde"),
            (getattr(d, "label", "") or "").lower()))

        cde_by_guid = {e.global_id: e for e in (elements or [])}
        self._node_by_gid = getattr(self.service, "last_node_by_gid", None) or {}
        active_only = self.vm.active_view_only

        rows = []
        guids = set(cde_by_guid.keys()) | set(self._revit_infos.keys())
        for guid in guids:
            info = self._revit_infos.get(guid)
            if active_only and info is None:
                continue
            cde = cde_by_guid.get(guid)
            values = dict(cde.values) if cde else {}
            if info:
                values.update(info.get("revit_params") or {})
            rows.append(self._ElementRow(guid, info, values, self.vm))

        rows.sort(key=lambda r: (r.Level, r.Mark, r.GlobalId))
        param_defs = defs or []
        element_list = elements or []

        try:
            self._param_defs = param_defs
            self._def_by_key = dict((d.key, d) for d in param_defs)
            self._refresh_param_combo()
            try:
                self.doorGrid.ItemsSource = None
            except Exception:
                pass
            self._sync_columns(element_list)
            self.vm.group_key = None
            self._clear_grouping()
            self.vm.set_rows(rows)
            self._refresh_grid_items_from_vm()
            if self._pending_restore:
                self.vm.apply_pending_snapshot(self._pending_restore)
                self._pending_restore = None
                self._update_pending_ui()
                self._refresh_grid_items_from_vm()

            matched = sum(1 for r in rows if r.MatchedInRevit)
            cde_count = len(element_list)
            prop_count = len(self._param_defs)
            self._set_status(self._format_refresh_status(rows, matched, cde_count, prop_count))

            self.doorGrid.ItemsSource = self._grid_items
            self._apply_grouping()
            if self._dfp_markers_active:
                self.Dispatcher.BeginInvoke(
                    self._DispatcherPriority.Background,
                    self._Action(self._refresh_dfp_markers))
        except Exception as ex:
            self._logger.error("CDE: apply grid failed: {}".format(ex))
            self._set_status("Internal error applying grid: {}".format(ex))

    def _format_refresh_status(self, rows, matched, cde_count, prop_count):
        if cde_count == 0:
            return ("{} Revit element(s); 0 from CDE for {}. Check sign-in / mapping / "
                    "revision.".format(len(rows), self.vm.ifc_class))
        message = "{} element(s), {} matched in Revit, {} CDE property(ies).".format(
            len(rows), matched, prop_count)
        dfp_defs = sum(
            1 for d in self._param_defs
            if getattr(d, "group", "") == "Function")
        dfp_cells = sum(
            1 for r in rows
            for k in r.cells
            if k.startswith("Pset_DFP."))
        if dfp_defs and dfp_cells == 0:
            message += (
                "  Door functions (Pset_DFP): 0 values in CDE for this revision — "
                "assign in Hub door cards or run DFP recipes, then refresh.")
        elif dfp_cells:
            doors_with_dfp = sum(
                1 for r in rows
                if any(k.startswith("Pset_DFP.") for k in r.cells))
            message += "  Door functions on {} door(s).".format(doors_with_dfp)
        trunc = getattr(self.service, "last_truncation", None)
        if trunc:
            reason = trunc.get("reason")
            if reason == "pagination_cap":
                label = trunc.get("ifc_class") or self.vm.ifc_class or "elements"
                message += (
                    "  WARNING: CDE fetch for {} may be incomplete "
                    "({} retrieved; pagination stopped early).".format(
                        label, trunc.get("retrieved")))
            elif trunc.get("total") is not None:
                message += (
                    "  WARNING: CDE returned only {} of {} elements "
                    "(backend cap) — list may be incomplete.".format(
                        trunc["retrieved"], trunc["total"]))
        try:
            from cde.request_log import (
                get_log_path, get_get_index_path, get_capture_dir)
            self._logger.debug(
                "CDE: HTTP capture index={} GETs={} bodies={}".format(
                    get_log_path(), get_get_index_path(), get_capture_dir()))
        except Exception:
            pass
        return message

    def _refresh_param_combo(self):
        self._all_param_items = [
            self._ParamDisplayItem(d, self._cde_brush, self._revit_brush)
            for d in self._param_defs]
        from System.Collections.ObjectModel import ObservableCollection
        from System import Object
        collection = ObservableCollection[Object]()
        for item in self._all_param_items:
            collection.Add(item)
        self.paramCombo.ItemsSource = collection
        if self._all_param_items:
            self.paramCombo.SelectedIndex = 0
        else:
            self.paramCombo.SelectedIndex = -1

    # --- filtering / grouping ------------------------------------------

    def on_filter_changed(self, sender, args):
        self.vm.filter_text = self.filterBox.Text or ""
        self._clear_grouping()
        self.vm.apply_filter()
        self._refresh_grid_items_from_vm()
        self.doorGrid.ItemsSource = self._grid_items
        self._apply_grouping()
        self._maybe_refresh_dfp_markers()

    def on_group_changed(self, sender, args):
        if getattr(self, "_suppress_group_changed", False):
            return
        item = self.groupCombo.SelectedItem
        self.vm.group_key = item.value if item else None
        self._apply_grouping()

    def _clear_grouping(self):
        try:
            view = self._CollectionViewSource.GetDefaultView(self._grid_items)
            if view is not None:
                view.GroupDescriptions.Clear()
        except Exception as ex:
            self._logger.debug("CDE: clear grouping failed: {}".format(ex))

    def _apply_grouping(self):
        try:
            view = self._CollectionViewSource.GetDefaultView(self._grid_items)
            if view is None:
                return
            view.GroupDescriptions.Clear()
            path = self.vm.group_key
            if not path:
                item = self.groupCombo.SelectedItem
                path = item.value if item else None
            if path:
                pgd = self._PropertyGroupDescription(".", self._group_converter)
                view.GroupDescriptions.Add(pgd)
        except Exception as ex:
            self._logger.debug("CDE: grouping failed: {}".format(ex))

    def on_category_changed(self, sender, args):
        if self.IsLoaded:
            self.refresh()

    def on_active_view_toggled(self, sender, args):
        # ViewActivated subscription must happen on the Revit API thread;
        # route through ExternalEventRunner instead of calling directly.
        enabled = bool(self.activeViewCheck.IsChecked)
        def task(uiapp):
            if enabled:
                self._watcher.start(uiapp)
            else:
                self._watcher.stop()
        self._runner.run(task)
        self.refresh()

    def _on_active_view_changed(self, view):
        if bool(self.activeViewCheck.IsChecked):
            self.refresh()
        elif self._dfp_markers_active:
            self._refresh_dfp_markers()

    # --- dynamic columns (searchable combo + header grouping) ------------

    def _source_header_brush(self, source):
        color = self._WpfColor.FromArgb(28, 0, 120, 212)
        if source == "revit":
            color = self._WpfColor.FromArgb(28, 232, 108, 0)
        return self._SolidColorBrush(color)

    def _source_border_brush(self, source):
        return self._revit_brush if source == "revit" else self._cde_brush

    def _build_column_header(self, label, source, group_key):
        """Header with source tint and group icon (visible on hover)."""
        panel = self._StackPanel()
        panel.Orientation = self._Orientation.Horizontal
        panel.Margin = self._Thickness(4, 0, 4, 0)
        panel.Background = self._source_header_brush(source)
        panel.Tag = group_key

        icon = self._TextBlock()
        icon.Text = self._ICON_CDE if source == "cde" else self._ICON_REVIT
        icon.FontFamily = self._icon_font
        icon.Foreground = self._source_border_brush(source)
        icon.FontSize = 12
        icon.Margin = self._Thickness(0, 0, 4, 0)
        icon.VerticalAlignment = self._VerticalAlignment.Center
        panel.Children.Add(icon)

        title = self._TextBlock()
        title.Text = label
        title.VerticalAlignment = self._VerticalAlignment.Center
        title.Margin = self._Thickness(0, 0, 6, 0)
        try:
            title.Foreground = self.FindResource("TextBrush")
        except Exception:
            pass
        panel.Children.Add(title)

        group_btn = self._Button()
        group_btn.Content = self._ICON_GROUP
        group_btn.FontFamily = self._icon_font
        group_btn.ToolTip = "Group by this column"
        group_btn.Padding = self._Thickness(2, 0, 2, 0)
        group_btn.MinWidth = 22
        group_btn.MinHeight = 22
        group_btn.Visibility = self._Visibility.Collapsed
        group_btn.Tag = group_key
        group_btn.Click += self.on_header_group_click
        try:
            group_btn.Style = self.FindResource("SecondaryButtonStyle")
        except Exception:
            pass
        panel.Children.Add(group_btn)

        def show_group_btn(sender, args):
            group_btn.Visibility = self._Visibility.Visible

        def hide_group_btn(sender, args):
            group_btn.Visibility = self._Visibility.Collapsed

        panel.MouseEnter += show_group_btn
        panel.MouseLeave += hide_group_btn
        return panel

    def _set_column_key(self, column, key):
        """Map grid column -> logical key (IronPython may lack DataGridColumn.Tag)."""
        if column is not None and key:
            self._column_key_by_column[column] = key
            try:
                column.Tag = key
            except Exception:
                pass

    def _column_key(self, column):
        if column is None:
            return None
        key = self._column_key_by_column.get(column)
        if key:
            return key
        try:
            return column.Tag
        except Exception:
            return None

    def _style_fixed_column_headers(self):
        try:
            for idx, (path, label, source) in enumerate(self._fixed_column_meta):
                if idx >= self.doorGrid.Columns.Count:
                    break
                col = self.doorGrid.Columns[idx]
                binding = self._Binding(".")
                binding.Converter = self._cell_converter
                binding.ConverterParameter = path
                col.Binding = binding
                col.Header = self._build_column_header(label, source, path)
                self._set_column_key(col, path)
        except Exception as ex:
            self._logger.debug("CDE: fixed header styling failed: {}".format(ex))

    def on_header_group_click(self, sender, args):
        key = sender.Tag
        if not key:
            return
        self.vm.group_key = key
        self._sync_group_combo(key)
        self._apply_grouping()
        self._set_status("Grouped by '{}'.".format(key))

    def _sync_group_combo(self, key):
        try:
            for idx, item in enumerate(self.groupCombo.ItemsSource or []):
                if item.value == key:
                    self._suppress_group_changed = True
                    try:
                        self.groupCombo.SelectedIndex = idx
                    finally:
                        self._suppress_group_changed = False
                    return
            # Header grouped by a column outside toolbar presets — avoid stale combo label.
            self._suppress_group_changed = True
            try:
                self.groupCombo.SelectedIndex = 0
            finally:
                self._suppress_group_changed = False
        except Exception:
            pass

    def _end_grid_edit(self, commit=False):
        """Leave cell/row edit mode before programmatic grid refresh (avoids native crash)."""
        try:
            unit_cell = self._DataGridEditingUnit.Cell
            unit_row = self._DataGridEditingUnit.Row
            if commit:
                self.doorGrid.CommitEdit(unit_cell, True)
                self.doorGrid.CommitEdit(unit_row, True)
            else:
                self.doorGrid.CancelEdit(unit_cell)
                self.doorGrid.CancelEdit(unit_row)
        except Exception as ex:
            self._logger.debug("CDE: end grid edit failed: {}".format(ex))

    def _refresh_grid_cells(self):
        """Rebind dynamic cells after programmatic multi-row updates."""
        try:
            view = self._CollectionViewSource.GetDefaultView(self._grid_items)
            if view is not None:
                view.Refresh()
        except Exception as ex:
            self._logger.debug("CDE: grid view refresh failed: {}".format(ex))
        try:
            self.doorGrid.Items.Refresh()
        except Exception:
            pass
        try:
            self.doorGrid.InvalidateVisual()
        except Exception:
            pass

    def _normalize_global_id(self, gid):
        try:
            return unicode(gid or "").strip().upper()
        except NameError:
            return str(gid or "").strip().upper()

    def _ensure_dfp_column_for_key(self, key):
        """Show a Pset_DFP column in the grid when staging from view markers."""
        if not key or key in self._value_column_map:
            return
        pdef = self._def_by_key.get(key)
        if pdef is None and self._is_dfp_param_key(key):
            parts = key.split(".", 1)
            label = parts[1] if len(parts) == 2 else key
            pdef = self._param_def(
                key, label, value_type="bool", source="cde", group="Function")
            self._def_by_key[key] = pdef
        if pdef is not None:
            self._add_value_column(key, pdef.label, pdef)

    def _select_grid_row(self, row):
        """Focus schedule row after marker edit (grid binds string keys)."""
        try:
            for idx, candidate in enumerate(self.vm.rows):
                if candidate is row:
                    key = self._make_grid_key(idx, row)
                    self.doorGrid.SelectedItem = key
                    self.doorGrid.ScrollIntoView(key)
                    break
        except Exception as ex:
            self._logger.debug("CDE: select grid row failed: {}".format(ex))

    def _merge_row_detail_values(self, row):
        """Row cells (incl. pending) plus non-pending graph payload for inspector."""
        cached = dict(row.cells)
        pending = row.pending_for()
        node = self._node_by_gid.get(row.GlobalId)
        if node is not None:
            graph_vals = getattr(self.service, "graph_node_values", None)
            if graph_vals is not None:
                try:
                    for key, val in (graph_vals(node) or {}).items():
                        if key not in pending:
                            cached[key] = val
                except Exception as ex:
                    self._logger.debug(
                        "CDE: detail graph merge failed: {}".format(ex))
        return cached

    def _refresh_row_detail_pane(self, row):
        """Update property inspector when the selected row was edited."""
        if row is None:
            return
        selected = self._resolve_selected_row()
        if selected is None or self._normalize_global_id(selected.GlobalId) != self._normalize_global_id(row.GlobalId):
            return
        self._apply_detail_pane(row, self._merge_row_detail_values(row))

    def _refresh_staged_row_ui(self, row, key=None):
        """Refresh grid + detail after staging from markers or bulk edit."""
        if key and self._is_dfp_param_key(key):
            self._ensure_dfp_column_for_key(key)
        self._refresh_grid_cells()
        self._update_pending_ui()
        self._refresh_row_detail_pane(row)

    def _add_value_column(self, key, label=None, param_def_row=None):
        """Add a dynamic property column (idempotent on key)."""
        if key in self._value_column_map:
            return
        pdef = param_def_row or self._def_by_key.get(key)
        if pdef is not None:
            label = label or pdef.label
            source = getattr(pdef, "source", "cde") or "cde"
            value_type = getattr(pdef, "value_type", "string")
        else:
            source = "cde"
            value_type = "string"
            label = label or key

        if value_type == "bool":
            column = self._DataGridCheckBoxColumn()
        else:
            column = self._DataGridTextColumn()

        binding = self._Binding(".")
        binding.Converter = self._cell_converter
        binding.ConverterParameter = key
        column.Binding = binding
        column.Header = self._build_column_header(label, source, key)
        column.Width = self._value_column_width
        column.IsReadOnly = False
        column.CanUserSort = True
        column.SortMemberPath = "SortValue"
        self._set_column_key(column, key)
        self.doorGrid.Columns.Add(column)
        self._value_column_map[key] = column
        self._column_meta[key] = {"source": source, "value_type": value_type, "label": label}
        if key not in self.vm.value_columns:
            self.vm.value_columns.append(key)

    def _remove_value_column(self, key):
        """Remove a previously added dynamic column."""
        column = self._value_column_map.pop(key, None)
        self._column_meta.pop(key, None)
        if column is not None:
            self._column_key_by_column.pop(column, None)
        if column is not None and self.doorGrid.Columns.Contains(column):
            self.doorGrid.Columns.Remove(column)
        if key in self.vm.value_columns:
            self.vm.value_columns.remove(key)

    def _is_dfp_param_def(self, pdef):
        key = getattr(pdef, "key", "") or ""
        if key.startswith(self._dfp_pset_name + "."):
            return True
        return getattr(pdef, "group", "") == "Function"

    def _pick_default_column_defs(self, elements, param_defs):
        """Default columns: Pset_DFP only (doors), preferring functions with values."""
        dfp_defs = [d for d in (param_defs or []) if self._is_dfp_param_def(d)]
        if not dfp_defs:
            return []

        dfp_prefix = self._dfp_pset_name + "."
        counts = {}
        for element in (elements or []):
            for key, val in element.values.items():
                if not key.startswith(dfp_prefix):
                    continue
                if val is None or val is False or val == "" or val == 0:
                    continue
                counts[key] = counts.get(key, 0) + 1

        if counts:
            by_key = dict((d.key, d) for d in dfp_defs)
            ranked = sorted(counts.keys(), key=lambda k: (-counts[k], k))
            chosen = []
            for key in ranked:
                pdef = by_key.get(key)
                if pdef is not None:
                    chosen.append(pdef)
            if chosen:
                return chosen

        return dfp_defs

    def _sync_columns(self, elements):
        """Reconcile dynamic columns with the freshly loaded parameter set."""
        current_class = self.vm.ifc_class
        category_changed = current_class != self._columns_ifc_class
        if category_changed:
            for key in list(self._value_column_map.keys()):
                self._remove_value_column(key)
            if self._param_defs:
                chosen = self._pick_default_column_defs(
                    elements, self._param_defs)
                for d in chosen:
                    self._add_value_column(d.key, d.label, d)
                self._columns_ifc_class = current_class
        else:
            valid_keys = set(d.key for d in self._param_defs)
            for key in list(self._value_column_map.keys()):
                if key not in valid_keys:
                    self._remove_value_column(key)
        # Column picker repopulates on DropDownOpened — skip on every refresh.

    def _populate_columns_combo(self):
        """Build searchable column picker items from current defs."""
        shown = set(self.vm.value_columns)
        self._all_column_items = [
            self._ColumnPickerItem(
                d, d.key in shown, self._cde_brush, self._revit_brush, self)
            for d in self._param_defs]
        from System.Collections.ObjectModel import ObservableCollection
        from System import Object
        collection = ObservableCollection[Object]()
        for item in self._all_column_items:
            collection.Add(item)
        self.columnsCombo.ItemsSource = collection
        self.columnsCombo.SelectedIndex = -1

    def on_column_picker_checked(self, sender, args):
        item = sender.DataContext
        if item is None:
            return
        key = item.key
        if sender.IsChecked:
            self._add_value_column(key, item.display_name, item.param_def)
        else:
            self._remove_value_column(key)
        item.is_checked = bool(sender.IsChecked)

    def on_columns_dropdown_opened(self, sender, args):
        self._columns_dropdown_open = True
        self._populate_columns_combo()
        self.Dispatcher.BeginInvoke(
            self._DispatcherPriority.Loaded,
            self._Action(lambda: self._wire_columns_dropdown()))
        self.Dispatcher.BeginInvoke(
            self._DispatcherPriority.Loaded,
            self._Action(lambda: self._wire_search_textbox(
                self.columnsCombo, self._filter_column_items, "_columns_search_wired")))

    def _wire_columns_dropdown(self):
        """Keep columns dropdown open while toggling checkboxes."""
        try:
            if not self.columnsCombo.Template:
                return
            popup = self.columnsCombo.Template.FindName("Popup", self.columnsCombo)
            if popup is not None:
                popup.StaysOpen = True
        except Exception as ex:
            self._logger.debug("CDE: columns popup stays-open failed: {}".format(ex))

    def on_columns_dropdown_closed(self, sender, args):
        self._columns_dropdown_open = False
        try:
            if self.columnsCombo.Template:
                popup = self.columnsCombo.Template.FindName("Popup", self.columnsCombo)
                if popup is not None:
                    popup.StaysOpen = False
        except Exception:
            pass
        self.columnsCombo.SelectedIndex = -1

    def on_param_dropdown_opened(self, sender, args):
        self.Dispatcher.BeginInvoke(
            self._DispatcherPriority.Loaded,
            self._Action(lambda: self._wire_search_textbox(
                self.paramCombo, self._filter_param_items, "_param_search_wired")))

    def _wire_search_textbox(self, combo, filter_handler, wired_attr):
        try:
            if not combo.Template:
                return
            search_textbox = combo.Template.FindName("SearchTextBox", combo)
            if search_textbox is None:
                popup = combo.Template.FindName("Popup", combo)
                if popup and popup.Child:
                    from System.Windows.Media import VisualTreeHelper

                    def find_child_by_name(parent, name):
                        if parent is None:
                            return None
                        if hasattr(parent, "Name") and parent.Name == name:
                            return parent
                        for i in range(VisualTreeHelper.GetChildrenCount(parent)):
                            child = VisualTreeHelper.GetChild(parent, i)
                            found = find_child_by_name(child, name)
                            if found is not None:
                                return found
                        return None

                    search_textbox = find_child_by_name(popup.Child, "SearchTextBox")
            if search_textbox is None:
                return
            search_textbox.Text = ""
            if not getattr(self, wired_attr, False):
                search_textbox.TextChanged += filter_handler
                setattr(self, wired_attr, True)
            search_textbox.Focus()
        except Exception as ex:
            self._logger.debug("CDE: search textbox wiring failed: {}".format(ex))

    def _filter_param_items(self, sender, args):
        try:
            needle = (sender.Text or "").lower().strip()
            selected = self.paramCombo.SelectedItem
            selected_name = None
            if selected is not None:
                selected_name = getattr(selected, "display_name", None)
            from System.Collections.ObjectModel import ObservableCollection
            from System import Object
            filtered = ObservableCollection[Object]()
            for item in self._all_param_items:
                if not needle or needle in item.display_name.lower():
                    filtered.Add(item)
            self.paramCombo.ItemsSource = filtered
            if selected_name:
                for i in range(filtered.Count):
                    if filtered[i].display_name == selected_name:
                        self.paramCombo.SelectedIndex = i
                        return
            if filtered.Count == 0:
                self.paramCombo.SelectedIndex = -1
        except Exception as ex:
            self._logger.debug("CDE: param filter failed: {}".format(ex))

    def _filter_column_items(self, sender, args):
        try:
            needle = (sender.Text or "").lower().strip()
            from System.Collections.ObjectModel import ObservableCollection
            from System import Object
            filtered = ObservableCollection[Object]()
            for item in self._all_column_items:
                if not needle or needle in item.display_name.lower():
                    filtered.Add(item)
            self.columnsCombo.ItemsSource = filtered
        except Exception as ex:
            self._logger.debug("CDE: column filter failed: {}".format(ex))

    def on_grid_sorting(self, sender, args):
        try:
            col = args.Column
            key = self._column_key(col)
            if key:
                self.vm.sort_key = key
        except Exception as ex:
            self._logger.debug("CDE: sorting hook failed: {}".format(ex))

    def _row_from_grid_row_container(self, row_container):
        """Resolve ElementRow from a DataGridRow (or grid key string)."""
        if row_container is None:
            return None
        item = getattr(row_container, "Item", None)
        if item is not None and hasattr(item, "record_pending"):
            return item
        return self._row_for_grid_key(item)

    def _row_from_cell_edit(self, args):
        """Resolve ElementRow from DataGridCellEditEndingEventArgs."""
        row = self._row_from_grid_row_container(args.Row)
        if row is not None:
            return row
        editing = args.EditingElement
        if editing is not None:
            ctx = getattr(editing, "DataContext", None)
            if ctx is not None and hasattr(ctx, "record_pending"):
                return ctx
            row = self._row_for_grid_key(ctx)
            if row is not None:
                return row
        return None

    def _stage_cde_cell_edit(self, key, row, new_value, show_status=True):
        """Stage one inline CDE cell edit and refresh pending UI."""
        if getattr(self, "_suppress_inline_staging", False):
            return
        meta = self._column_meta.get(key, {})
        if meta.get("source") != "cde":
            return
        targets = self._selected_rows()
        if row not in targets:
            targets = [row]
        for target in targets:
            target.record_pending(key, new_value)
        self._refresh_grid_cells()
        self._update_pending_ui()
        if show_status:
            if len(targets) > 1:
                self._set_status(
                    "Staged '{}' on {} row(s). Click Apply to commit.".format(
                        key, len(targets)))
            else:
                self._set_status(
                    "Staged '{}'. Click Apply to commit.".format(key))

    def _is_dfp_param_key(self, key):
        return bool(key) and key.startswith(self._dfp_pset_name + ".")

    def _row_for_global_id(self, gid):
        want = self._normalize_global_id(gid)
        if not want:
            return None
        for row in self.vm.all_rows():
            if self._normalize_global_id(row.GlobalId) == want:
                return row
        return None

    def _dfp_cell_active(self, row, key):
        if self._has_authored_dfp_value is None:
            return False
        return self._has_authored_dfp_value(row.get_cell(key))

    def _install_dfp_marker_click_handler(self):
        if self._register_dfp_marker_click_callback is None:
            return
        try:
            if self._register_dfp_schedule_window is not None:
                self._register_dfp_schedule_window(self)

            def on_cell_click(payload, command_data):
                import sys
                win = getattr(sys, "_pyBS_dfp_schedule_window", None)
                if win is None:
                    return
                p = dict(payload or {})
                c = command_data

                def ui():
                    try:
                        win._apply_dfp_marker_click(p, c)
                    except Exception as ex:
                        win._logger.warning(
                            "CDE: DFP marker click failed: {}".format(ex))

                try:
                    win.Dispatcher.Invoke(
                        win._DispatcherPriority.Normal,
                        win._Action(ui))
                except Exception:
                    win.Dispatcher.BeginInvoke(
                        win._DispatcherPriority.Normal,
                        win._Action(ui))

            self._register_dfp_marker_click_callback(on_cell_click)
        except Exception as ex:
            self._logger.warning(
                "CDE: DFP click handler install failed: {}".format(ex))

    def _stage_dfp_marker_edit(self, key, row, new_value):
        """Stage DFP toggle from view markers (bypasses column_meta gate)."""
        if getattr(self, "_suppress_inline_staging", False):
            return
        if not self._is_dfp_param_key(key):
            return
        row.record_pending(key, new_value)
        self._refresh_staged_row_ui(row, key=key)
        code = key.split(".", 1)[-1] if "." in key else key
        self._set_status(
            "Staged {} = {}. Click Apply to commit.".format(code, new_value))

    def _on_dfp_marker_clicked(self, payload, command_data):
        """Legacy entry — handler uses closure from _install_dfp_marker_click_handler."""
        self._apply_dfp_marker_click(payload, command_data)

    def _apply_dfp_marker_click(self, payload, command_data):
        gid = payload.get("global_id")
        key = payload.get("param_key")
        if not gid or not key:
            self._logger.debug(
                "CDE: DFP click ignored — gid={!r} key={!r}".format(gid, key))
            return
        row = self._row_for_global_id(gid)
        if row is None:
            self._logger.warning(
                "CDE: DFP click — no row for GlobalId {!r}".format(gid))
            return
        new_val = not self._dfp_cell_active(row, key)
        self._stage_dfp_marker_edit(key, row, new_val)
        self._select_grid_row(row)
        if self._apply_in_flight:
            return
        # Refresh graphics on next idling poll (same thread as hover) — avoids
        # queuing SetElementIds/TGM ExternalEvents that crash ArrowEditor on Apply.
        self._dfp_graphics_refresh_gid = gid

    def _arm_dfp_apply_quarantine(self):
        """Block DFP idling during apply (no TGM API — crashes ArrowEditor)."""
        if self._set_dfp_apply_block is not None:
            self._set_dfp_apply_block(True)
        self._apply_revit_quarantine = True

    def _finish_apply_revit_side(self):
        """Release apply locks. Never touch TGM here — user resets DFP via button."""
        self._apply_revit_quarantine = False
        were_active = self._dfp_markers_were_active
        self._dfp_markers_were_active = False
        if were_active:
            self._dfp_markers_need_manual_reset = True
            self._dfp_markers_active = False
            self._update_dfp_markers_button()
            if self._set_dfp_apply_block is not None:
                self._set_dfp_apply_block(True)
        else:
            self._dfp_markers_need_manual_reset = False
            if self._set_dfp_apply_block is not None:
                self._set_dfp_apply_block(False)

    def _release_dfp_frozen_state(self):
        if self._dfp_markers_need_manual_reset:
            self._dfp_markers_need_manual_reset = False
            if self._set_dfp_apply_block is not None:
                self._set_dfp_apply_block(False)

    def _dfp_off_apply_hint(self):
        if not self._dfp_markers_ok:
            return ""
        if self._dfp_markers_need_manual_reset:
            return (
                " DFP graphics frozen — click DFP markers to clear and re-enable.")
        return " DFP markers are off — use the button to turn them back on."

    def _maybe_refresh_dfp_markers(self):
        if self._apply_in_flight or self._apply_revit_quarantine:
            return
        if self._dfp_markers_need_manual_reset:
            return
        if self._dfp_markers_active:
            self._refresh_dfp_markers()

    def on_preparing_cell_for_edit(self, sender, args):
        """Wire checkbox toggles so pending updates before the cell loses focus."""
        try:
            col = args.Column
            key = self._column_key(col)
            if not key or key not in self._value_column_map:
                return
            meta = self._column_meta.get(key, {})
            if meta.get("source") != "cde":
                return
            editing = args.EditingElement
            if not isinstance(editing, self._CheckBox):
                return
            row_container = args.Row
            editing.Tag = key
            editing._cde_row = row_container
            if not getattr(editing, "_cde_inline_wired", False):
                editing.Checked += self._on_inline_checkbox_changed
                editing.Unchecked += self._on_inline_checkbox_changed
                editing._cde_inline_wired = True
        except Exception as ex:
            self._logger.debug("CDE: preparing cell for edit failed: {}".format(ex))

    def _on_inline_checkbox_changed(self, sender, args):
        """Apply checkbox toggle immediately (CellEditEnding fires only on row change)."""
        try:
            key = sender.Tag
            row = self._row_from_grid_row_container(sender._cde_row)
            if not key or row is None:
                return
            self._stage_cde_cell_edit(key, row, bool(sender.IsChecked))
            if self._is_dfp_param_key(key):
                self._maybe_refresh_dfp_markers()
        except Exception as ex:
            self._logger.debug("CDE: inline checkbox change failed: {}".format(ex))

    def on_cell_edit_ending(self, sender, args):
        if args.EditAction == self._DataGridEditAction.Cancel:
            return
        try:
            col = args.Column
            key = self._column_key(col)
            if not key or key not in self._value_column_map:
                return
            meta = self._column_meta.get(key, {})
            if meta.get("source") != "cde":
                return
            row = self._row_from_cell_edit(args)
            if row is None:
                return
            new_value = None
            if isinstance(args.EditingElement, self._CheckBox):
                new_value = bool(args.EditingElement.IsChecked)
            elif isinstance(args.EditingElement, self._TextBox):
                text = args.EditingElement.Text
                if meta.get("value_type") == "bool":
                    new_value = text.strip().lower() in ("1", "true", "yes")
                else:
                    new_value = text
            if new_value is None:
                return
            show_status = not isinstance(args.EditingElement, self._CheckBox)
            self._stage_cde_cell_edit(key, row, new_value, show_status=show_status)
        except Exception as ex:
            self._logger.debug("CDE: cell edit failed: {}".format(ex))

    # --- coloring -------------------------------------------------------

    def _dfp_marker_rows(self):
        """Doors shown in the schedule grid (filter/group), not whole model."""
        try:
            return list(self.vm.rows)
        except Exception:
            return self._current_rows()

    def _current_rows(self):
        return self.vm.all_rows()

    def on_color_click(self, sender, args):
        item = self.paramCombo.SelectedItem
        if item is None:
            self._set_status("Select a parameter to color by.")
            return
        pdef = item.param_def
        key = pdef.key
        rows = self._current_rows()
        self._set_status("Applying colors...")

        def task(uiapp):
            doc = uiapp.ActiveUIDocument.Document
            view = doc.ActiveView
            color_map = self._coloring.apply_coloring(doc, view, rows, key)
            legend = self._coloring.build_legend(color_map)
            self.Dispatcher.Invoke(
                self._Action(lambda: self._render_legend(legend, pdef.label)))

        self._runner.run(task)

    def _render_legend(self, legend, label):
        items = []
        for value, rgb in legend:
            brush = self._SolidColorBrush(
                self._WpfColor.FromRgb(rgb[0], rgb[1], rgb[2]))
            items.append(self._LegendItem(value, brush))
        self.legendList.ItemsSource = items
        self._set_status("Colored by '{}'.".format(label))

    def on_reset_color_click(self, sender, args):
        rows = self._current_rows()

        def task(uiapp):
            doc = uiapp.ActiveUIDocument.Document
            self._coloring.reset_coloring(doc, doc.ActiveView, rows)
            self.Dispatcher.Invoke(self._Action(lambda: self._clear_legend()))

        self._runner.run(task)

    def _clear_legend(self):
        self.legendList.ItemsSource = None
        self._set_status("Colors reset.")

    # --- DFP temporary graphics (Hub icon markers) ----------------------

    def _visible_dfp_param_keys(self):
        visible = list(self._value_column_map.keys())
        if self._default_dfp_param_keys is None:
            return []
        keys = self._default_dfp_param_keys(self._param_defs, visible_keys=visible)
        if keys:
            return keys
        return self._default_dfp_param_keys(self._param_defs)

    def _all_dfp_param_keys(self):
        """Full DFP catalog for hover overlay (all functions, not just visible cols)."""
        if self._default_dfp_param_keys is None:
            return []
        return self._default_dfp_param_keys(self._param_defs)

    def _set_dfp_status(self, message):
        self.Dispatcher.BeginInvoke(
            self._DispatcherPriority.Normal,
            self._Action(lambda: self._set_status(message)))

    def _update_dfp_markers_button(self):
        if not hasattr(self, "dfpMarkersButton"):
            return
        if self._dfp_markers_active:
            self.dfpMarkersButton.Content = "DFP markers ON"
            try:
                self.dfpMarkersButton.Style = self.FindResource("DefaultButtonStyle")
            except Exception:
                pass
        else:
            self.dfpMarkersButton.Content = "DFP markers"
            try:
                self.dfpMarkersButton.Style = self.FindResource("SecondaryButtonStyle")
            except Exception:
                pass

    def on_dfp_markers_click(self, sender, args):
        self._release_dfp_frozen_state()
        if not self._dfp_markers_ok or self._build_dfp_view_points is None:
            msg = (
                "DFP markers require Revit 2022 or newer "
                "(TemporaryGraphicsManager API).")
            if self._dfp_markers_import_error:
                msg += "\n\n{}".format(self._dfp_markers_import_error)
            forms.alert(msg, title="DFP Markers")
            return
        if self._dfp_markers_active:
            self._dfp_markers_active = False
            self._update_dfp_markers_button()
            self._update_pending_ui()

            def off_task(uiapp):
                try:
                    doc = uiapp.ActiveUIDocument.Document
                    cleaner = (
                        self._soft_clean_dfp_session
                        or self._clean_dfp_session)
                    if cleaner is not None:
                        cleaner(doc)
                except Exception as ex:
                    self._logger.warning("CDE: DFP markers off failed: {}".format(ex))
            self._runner.run(off_task)
            self._set_status("DFP markers off.")
            return

        self._dfp_markers_active = True
        self._update_dfp_markers_button()
        self._update_pending_ui()
        self._set_status("Placing DFP markers...")
        self._refresh_dfp_markers()

    def _refresh_dfp_markers(self):
        if not self._dfp_markers_active or not self._dfp_markers_ok:
            return
        if self._apply_in_flight:
            return
        rows = self._dfp_marker_rows()
        param_keys = self._all_dfp_param_keys()
        if not param_keys:
            param_keys = self._visible_dfp_param_keys()
        if not param_keys:
            self._set_status(
                "No DFP columns visible. Show Pset_DFP columns or refresh data.")
            return

        def task(uiapp):
            try:
                doc = uiapp.ActiveUIDocument.Document
                if self._ensure_dfp_tgm_handler is not None:
                    self._ensure_dfp_tgm_handler(doc, logger=self._logger)
                view = doc.ActiveView
                view_map = self._build_dfp_view_points(
                    doc, rows, param_keys, view=view)
                if not view_map:
                    self._set_dfp_status(
                        "No authored DFP values on matched doors in this view.")
                    return
                self._register_dfp_session(doc, view_map, toggle_active=True)
                self._set_dfp_toggle_active(doc, True)
                driver = self._start_or_get_dfp_driver(
                    uiapp, doc, view_map, self._get_element_id_value,
                    logger=self._logger)
                if driver is not None:
                    driver.set_enabled(True)
                    driver.refresh_all()
                control_count = sum(len(v) for v in view_map.values())
                door_gids = set()
                for pts in view_map.values():
                    for pt in pts:
                        gid = pt.get("global_id")
                        if gid:
                            door_gids.add(gid)
                self._set_dfp_status(
                    "DFP — {} door(s), {} function(s). "
                    "Hover summary to edit; click cell to stage.".format(
                        len(door_gids), len(param_keys)))
            except Exception as ex:
                self._logger.error("CDE: DFP markers failed: {}".format(ex))
                self._dfp_markers_active = False
                self.Dispatcher.BeginInvoke(
                    self._DispatcherPriority.Normal,
                    self._Action(self._update_dfp_markers_button))
                self._set_dfp_status("DFP markers failed: {}".format(ex))

        self._runner.run(task)

    # --- selection ------------------------------------------------------

    def _selected_rows(self):
        rows = []
        for item in self.doorGrid.SelectedItems:
            row = self._row_for_grid_key(item)
            if row is not None:
                rows.append(row)
        return rows

    def _sync_revit_selection(self):
        """Mirror grid selection into the active Revit document."""
        element_ids = [
            r.element_id for r in self._selected_rows() if r.element_id is not None]

        def task(uiapp):
            from System.Collections.Generic import List
            import clr as _clr
            _clr.AddReference("RevitAPI")
            from Autodesk.Revit.DB import ElementId
            id_list = List[ElementId]()
            for eid in element_ids:
                id_list.Add(eid)
            uiapp.ActiveUIDocument.Selection.SetElementIds(id_list)

        self._runner.run(task)

    # --- vertical value set / staged apply ------------------------------

    def _coerce_bulk_value(self, pdef, raw_text):
        if pdef.value_type == "bool":
            return raw_text.strip().lower() in ("1", "true", "yes")
        return raw_text

    def on_set_value_click(self, sender, args):
        item = self.paramCombo.SelectedItem
        if item is None:
            self._set_status("Select a parameter to set.")
            return
        pdef = item.param_def
        rows = self._selected_rows()
        if not rows:
            self._set_status("Select one or more rows first.")
            return
        if pdef.source == "revit":
            self._set_status("Revit parameter write-back is not wired yet.")
            return
        key = pdef.key
        value = self._coerce_bulk_value(pdef, self.valueBox.Text)
        for row in rows:
            row.record_pending(key, value)
        self._refresh_grid_cells()
        self._update_pending_ui()
        if self._is_dfp_param_key(key):
            self._maybe_refresh_dfp_markers()
        self._set_status(
            "Staged '{}' on {} row(s). Click Apply to commit.".format(
                pdef.label, len(rows)))

    def on_discard_pending_click(self, sender, args):
        if self._apply_in_flight:
            return
        self._end_grid_edit(commit=False)
        count = self.vm.pending_count()
        if count == 0:
            return
        self._suppress_inline_staging = True
        try:
            self.vm.discard_all_pending()
            self._refresh_grid_cells()
            self._update_pending_ui()
            self._maybe_refresh_dfp_markers()
            self._set_status("Discarded {} pending edit(s).".format(count))
        finally:
            self._suppress_inline_staging = False

    def _format_dry_run_message(self, plan):
        matched = plan.get("matchedElements") or []
        conflicts = plan.get("conflicts") or []
        return (
            "Dry run preview:\n"
            "- {} element(s) matched\n"
            "- {} conflict(s)\n\n"
            "Commit these changes to the CDE?".format(len(matched), len(conflicts)))

    def on_apply_pending_click(self, sender, args):
        if self._apply_in_flight:
            return
        self._end_grid_edit(commit=False)
        cleared = self._runner.clear_pending()
        self._dfp_graphics_refresh_gid = None
        changes = self.vm.collect_pending_changes()
        if not changes:
            self._set_status("No pending edits to apply.")
            return
        project_id, revision_id = self._project_revision_ids()
        if not project_id or not revision_id:
            self._set_status("Map the model before writing values.")
            return

        if self._dfp_markers_active:
            self._set_status(
                "Turn off DFP markers before Apply "
                "(keeps Revit stable after view edits).")
            return

        self._dfp_markers_were_active = False

        self._set_apply_in_flight(True)
        self._begin_apply_dry_run(changes, project_id, revision_id)

    def _begin_apply_dry_run(self, changes, project_id, revision_id):
        self._set_status("Running dry run...")

        def work():
            etag = self._get_revision_etag(project_id, revision_id)
            return self.service.apply_element_mutations(
                project_id, revision_id, changes, dry_run=True, etag=etag,
                file_version_id=self._file_version_id(),
                file_id=self._file_id())

        def done(outcome, error):
            if error is not None:
                self._set_apply_in_flight(False)
                self._finish_apply_revit_side()
                self._handle_apply_error(error)
                return
            if outcome and outcome.etag:
                self._save_revision_etag(outcome.etag)
            plan = (outcome.plan if outcome else {}) or {}
            msg = self._format_dry_run_message(plan)
            conflicts = plan.get("conflicts") or []
            if conflicts:
                self._set_status(
                    "Dry run: {} matched, {} conflict(s). Confirm to proceed.".format(
                        len(plan.get("matchedElements") or []), len(conflicts)))
            else:
                self._set_status(
                    "Dry run OK ({} element(s)). Confirm to commit.".format(
                        len(plan.get("matchedElements") or changes)))
            self._pending_commit_ctx = (project_id, revision_id)
            self._show_commit_confirm(True)
            self._set_status(
                "{}  Click 'Confirm commit' to write, or Cancel.".format(
                    self._format_dry_run_message(plan).split("\n\n")[0]
                    .replace("\n", "  ")))

        self._run_async(work, done)

    def _show_commit_confirm(self, show):
        """Toggle the non-modal commit confirm bar (avoids modal reentrancy crash)."""
        try:
            vis = (self._Visibility.Visible if show
                   else self._Visibility.Collapsed)
            self.confirmCommitButton.Visibility = vis
            self.cancelCommitButton.Visibility = vis
        except Exception as ex:
            self._logger.debug("CDE: confirm bar toggle failed: {}".format(ex))

    def on_cancel_commit_click(self, sender, args):
        self._show_commit_confirm(False)
        self._pending_commit_ctx = None
        self._set_apply_in_flight(False)
        self._finish_apply_revit_side()
        self._set_status("Apply cancelled." + self._dfp_off_apply_hint())

    def on_confirm_commit_click(self, sender, args):
        self._show_commit_confirm(False)
        ctx = self._pending_commit_ctx
        self._pending_commit_ctx = None
        if not ctx:
            return
        project_id, revision_id = ctx
        fresh_changes = self.vm.collect_pending_changes()
        if not fresh_changes:
            self._set_apply_in_flight(False)
            self._finish_apply_revit_side()
            self._set_status(
                "No pending edits to commit." + self._dfp_off_apply_hint())
            return
        self._commit_pending_changes(fresh_changes, project_id, revision_id)

    def _handle_apply_error(self, error):
        hint = self._dfp_off_apply_hint()
        if isinstance(error, CDEPreconditionError):
            project_id, revision_id = self._project_revision_ids()
            if project_id and revision_id:
                etag = self.service.fetch_revision_etag(project_id, revision_id)
                if etag:
                    self._cache_revision_etag(etag)
            self._set_status(
                "Model changed (stale etag). Click Refresh, then retry Apply."
                + hint)
            return
        if isinstance(error, CDEConflictError):
            self._set_status("Write conflict (409): {}{}".format(error, hint))
            return
        self._set_status("Apply failed: {}{}".format(error, hint))

    def _commit_pending_changes(self, changes, project_id, revision_id):
        self._set_apply_in_flight(True)
        self._set_status("Committing to CDE...")

        def work():
            etag = self._get_revision_etag(project_id, revision_id)
            try:
                outcome = self.service.apply_element_mutations(
                    project_id, revision_id, changes, dry_run=False, etag=etag,
                    file_version_id=self._file_version_id(),
                    file_id=self._file_id())
            except CDEPreconditionError:
                fresh_etag = self.service.fetch_revision_etag(
                    project_id, revision_id)
                if fresh_etag:
                    self._cache_revision_etag(fresh_etag)
                raise
            return outcome, changes

        def done(result, error):
            self._set_apply_in_flight(False)
            if error is not None:
                self._finish_apply_revit_side()
                self._handle_apply_error(error)
                return
            outcome, committed_changes = result
            status_data = outcome.status_data or {}
            if outcome.etag:
                self._save_revision_etag(outcome.etag)
            if outcome.mutation_id:
                status_data = self.service.poll_mutation(outcome.mutation_id)
            state = (status_data.get("status") or status_data.get("state") or "")
            state_lower = state.lower() if state else ""
            if state_lower in ("failed", "rejected", "error", "cancelled"):
                self._finish_apply_revit_side()
                self._set_status(
                    "Mutation {} — pending edits kept.{}".format(
                        state, self._dfp_off_apply_hint()))
                return
            if outcome.mutation_id:
                if not state_lower:
                    self._finish_apply_revit_side()
                    self._set_status(
                        "Mutation status unknown — pending edits kept."
                        + self._dfp_off_apply_hint())
                    return
                if state_lower not in self._mutation_success_states:
                    self._finish_apply_revit_side()
                    self._set_status(
                        "Mutation state '{}' — pending edits kept.{}".format(
                            state, self._dfp_off_apply_hint()))
                    return
            gids = list(committed_changes.keys())
            fresh_values = {}
            try:
                fresh_values = self.service.get_element_values(
                    project_id, revision_id, gids)
            except Exception as ex:
                self._logger.debug("CDE: post-commit read-back failed: {}".format(ex))
            confirmed_count = 0
            for row in self.vm.all_rows():
                gid = row.GlobalId
                if gid not in committed_changes:
                    continue
                updated = fresh_values.get(gid) or {}
                confirmed_keys = []
                for key in committed_changes[gid].keys():
                    if key in updated:
                        row.set_cell(key, updated[key])
                        confirmed_keys.append(key)
                if confirmed_keys:
                    row.commit_pending(confirmed_keys)
                    confirmed_count += len(confirmed_keys)
            self._refresh_grid_cells()
            self._update_pending_ui()
            self._finish_apply_revit_side()
            if outcome.etag:
                self._cache_revision_etag(outcome.etag)
                self._pending_etag_persist = outcome.etag
            if confirmed_count:
                self._set_status(
                    "Committed {} cell edit(s) on {} element(s).".format(
                        confirmed_count, len(committed_changes)))
            elif outcome.mutation_id:
                self._set_status(
                    "Mutation accepted; read-back incomplete — pending edits kept.")
            else:
                self._set_status("Commit completed.")

        self._run_async(work, done)

    # --- element detail inspector --------------------------------------

    def _resolve_selected_row(self):
        """Return the selected ElementRow, or None."""
        selected = self.doorGrid.SelectedItem
        if selected is not None and hasattr(selected, "GlobalId"):
            return selected
        row = self._row_for_grid_key(selected)
        if row is not None:
            return row
        rows = self._selected_rows()
        if rows:
            return rows[0]
        return None

    def _property_rows_for_detail(self, row, cde_values):
        """Build inspector rows: Revit fields plus full CDE graph payload."""
        rows = [
            self._PropertyRow("Revit.Mark", row.Mark),
            self._PropertyRow("Revit.Level", row.Level),
            self._PropertyRow("Revit.FromRoomNumber", row.FromRoomNumber),
            self._PropertyRow("Revit.FromRoom", row.FromRoom),
            self._PropertyRow("Revit.ToRoomNumber", row.ToRoomNumber),
            self._PropertyRow("Revit.ToRoom", row.ToRoom),
            self._PropertyRow("Revit.MatchedInRevit", row.MatchedInRevit),
        ]

        def _detail_sort_key(key):
            if key.startswith("RuleTrace."):
                return (5, key)
            if key.startswith("Relationship."):
                return (6, key)
            if key.startswith("authored."):
                return (2, key)
            if key.startswith("derived."):
                return (3, key)
            if key.startswith("effective."):
                return (4, key)
            return (1, key)

        for key, val in sorted((cde_values or {}).items(), key=lambda kv: _detail_sort_key(kv[0])):
            rows.append(self._PropertyRow(key, val))
        return rows

    def _apply_detail_pane(self, row, cde_values, note=None):
        if row is None:
            self.detailHeader.Text = "Select a row to inspect its CDE properties."
            self.propGrid.ItemsSource = None
            return
        header = u"Element: {}   Mark: {}".format(row.GlobalId, row.Mark)
        if note:
            header += u"   — {}".format(note)
        self.detailHeader.Text = header
        prop_rows = self._property_rows_for_detail(row, cde_values)
        self.propGrid.ItemsSource = prop_rows if prop_rows else None

    def _fetch_detail_for_row(self, row, token):
        if not self.mapping and not self.offline:
            self._apply_detail_pane(row, dict(row.cells),
                                    "Not mapped — showing cached row cells only.")
            return
        project_id = self.mapping["project_id"]
        revision_id = self.mapping["revision_id"]
        gid = row.GlobalId

        def work():
            element = None
            error = None
            try:
                fetch = getattr(self.service, "fetch_element_detail", None)
                if fetch is not None:
                    element = fetch(
                        project_id, revision_id, gid, self._node_by_gid)
                else:
                    element = self.service.get_element(
                        project_id, revision_id, gid)
            except Exception as ex:
                error = ex
            return element, error

        def done(result, error):
            if token != self._detail_fetch_token:
                return
            current = self._resolve_selected_row()
            if current is None or current.GlobalId != gid:
                return
            element, fetch_error = result if result is not None else (None, error)
            if fetch_error is not None:
                merged = dict(row.cells)
                self._apply_detail_pane(
                    row, merged,
                    "CDE fetch failed — showing {} cached cell(s).".format(len(merged)))
                return
            if element is None:
                merged = dict(row.cells)
                self._apply_detail_pane(
                    row, merged,
                    "No CDE element for this GlobalId in the current revision.")
                return
            merged = dict(row.cells)
            merged.update(element.values or {})
            self._apply_detail_pane(row, merged)

        self._run_async(work, done)

    def on_grid_selection_changed(self, sender, args):
        row = self._resolve_selected_row()
        # Revit selection sync intentionally disabled: SetElementIds() triggers
        # the native Revit SelectionChanged event which crashes Coordination
        # Model Interface addin in Revit 2026.
        self._detail_fetch_token += 1
        if row is None:
            self._apply_detail_pane(None, None)
            return
        cached = self._merge_row_detail_values(row)
        self._apply_detail_pane(row, cached)

    # --- lifecycle ------------------------------------------------------

    def window_closing(self, sender, args):
        if self._apply_in_flight:
            args.Cancel = True
            return
        pending = self.vm.pending_count()
        if pending > 0:
            if not forms.alert(
                    "You have {} pending edit(s). Close without applying?".format(
                        pending),
                    yes=True, no=True):
                args.Cancel = True
                return

        def stop_watcher(uiapp):
            try:
                self._flush_deferred_etag_persist()
            except Exception:
                pass
            try:
                self._watcher.stop()
            except Exception:
                pass
            if self._dfp_markers_active and self._clean_dfp_session is not None:
                try:
                    doc = uiapp.ActiveUIDocument.Document
                    self._clean_dfp_session(doc)
                except Exception:
                    pass

        self._runner.run(stop_watcher)


def _detect_offline(window):
    """Use the offline mock service if no token and no mapping exist."""
    if not window.mapping and not window.auth.is_authenticated():
        window.offline = True
        window.service = MockCDEService(window.auth)
        window._set_status("Offline demo mode (no mapping/sign-in found). "
                           "Showing sample data.")


if __name__ == "__main__":
    __window__ = ScheduleWindow()
    __window__.Closing += __window__.window_closing
    _detect_offline(__window__)
    __window__.Show()
    __window__.refresh()
