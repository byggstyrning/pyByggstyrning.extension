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
    Orientation, DataGridEditAction)
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
from cde.revit_events import ExternalEventRunner, ActiveViewWatcher

logger = script.get_logger()

# Categories the schedule can drive (label -> IfcClass).
CATEGORIES = [("Doors", "IfcDoor"), ("Windows", "IfcWindow"), ("Walls", "IfcWall")]
# Group-by options (label -> ElementRow property path, None = ungrouped).
GROUP_OPTIONS = [("No grouping", None), ("Level", "Level"),
                 ("From room", "FromRoom"), ("To room", "ToRoom")]
# How many CDE properties to auto-show as columns the first time data loads.
DEFAULT_COLUMN_COUNT = 10
# DEBUG: temporarily disable the Revit-parameter scan to isolate the native crash.
_ENABLE_REVIT_PARAM_SCAN = False
# Fixed grid columns that support header grouping (path -> source tint).
FIXED_COLUMN_META = [
    ("GlobalId", "GlobalId", "revit"),
    ("Mark", "Mark", "revit"),
    ("Level", "Level", "revit"),
    ("FromRoom", "From room", "revit"),
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

    def Convert(self, value, target_type, parameter, culture):
        try:
            raw = value.get_cell(parameter)
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
        self.doorGrid.ItemsSource = CollectionViewSource.GetDefaultView(self.vm.rows)
        self._cell_converter = CellValueConverter()
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
        self._pending_restore = None
        self._detail_fetch_token = 0
        self._node_by_gid = {}
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
        self._default_columns = DEFAULT_COLUMN_COUNT
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
        self._ICON_GROUP = ICON_GROUP
        self._ICON_CDE = ICON_CDE
        self._ICON_REVIT = ICON_REVIT
        self._mutation_success_states = MUTATION_SUCCESS_STATES
        self._cde_brush = self._SolidColorBrush(self._WpfColor.FromRgb(0, 120, 212))
        self._revit_brush = self._SolidColorBrush(self._WpfColor.FromRgb(232, 108, 0))
        self._icon_font = self._FontFamily("Segoe MDL2 Assets")

        self._init_combos()
        self._style_fixed_column_headers()
        self.vm.on_pending_changed(self._update_pending_ui)
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
            enabled = count > 0 and not self._apply_in_flight
            self.applyPendingButton.IsEnabled = enabled
            self.discardPendingButton.IsEnabled = enabled
        except Exception as ex:
            self._logger.debug("CDE: pending UI update failed: {}".format(ex))

    def _set_apply_in_flight(self, in_flight):
        self._apply_in_flight = bool(in_flight)
        try:
            self.applyPendingButton.IsEnabled = not in_flight and self.vm.pending_count() > 0
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

    def _cache_revision_etag(self, etag):
        """Memory-only etag update (safe from background threads)."""
        if etag and self.mapping is not None:
            self.mapping["revision_etag"] = etag

    def _persist_revision_etag(self, etag):
        """Write etag to Extensible Storage on the Revit API thread."""
        if not etag:
            return
        self._cache_revision_etag(etag)

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
                        self._DispatcherPriority.ApplicationIdle,
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
            count = self.vm.pending_count()
            if not forms.alert(
                    "You have {} pending edit(s). Refresh will discard them.\n\n"
                    "Continue?".format(count),
                    yes=True, no=True):
                return
            self.vm.discard_all_pending()
        elif preserve_pending:
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
        project_id = self.mapping["project_id"] if self.mapping else "demo-1"
        revision_id = self.mapping["revision_id"] if self.mapping else "rev-2"
        ifc_class = self.vm.ifc_class
        self._fetch_running = True
        self._set_status(
            "Fetching CDE elements...  [rev: {}]".format(revision_id or "(none)"))

        def work():
            elements = self.service.list_elements(project_id, revision_id, ifc_class)
            defs = self.service.get_parameter_defs(project_id, ifc_class)
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
            for e in elements:
                for key, val in e.values.items():
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
            if self._pending_restore:
                self.vm.apply_pending_snapshot(self._pending_restore)
                self._pending_restore = None
                self._update_pending_ui()

            matched = sum(1 for r in rows if r.MatchedInRevit)
            cde_count = len(element_list)
            prop_count = len(self._param_defs)
            self._set_status(self._format_refresh_status(rows, matched, cde_count, prop_count))

            grid_view = self._CollectionViewSource.GetDefaultView(self.vm.rows)
            self.doorGrid.ItemsSource = grid_view
            self._apply_grouping()
            self._logger.info(
                "CDE: built {} rows ({} CDE, {} matched in Revit), {} prop(s) for {}".format(
                    len(rows), cde_count, matched, prop_count, self.vm.ifc_class))
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
            message += ("  WARNING: CDE returned only {} of {} elements "
                        "(backend cap) - list is INCOMPLETE.".format(
                            trunc["retrieved"], trunc["total"]))
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
        self.doorGrid.ItemsSource = self._CollectionViewSource.GetDefaultView(self.vm.rows)
        self._apply_grouping()

    def on_group_changed(self, sender, args):
        if getattr(self, "_suppress_group_changed", False):
            return
        item = self.groupCombo.SelectedItem
        self.vm.group_key = item.value if item else None
        self._apply_grouping()

    def _clear_grouping(self):
        try:
            view = self._CollectionViewSource.GetDefaultView(self.vm.rows)
            if view is not None:
                view.GroupDescriptions.Clear()
        except Exception as ex:
            self._logger.debug("CDE: clear grouping failed: {}".format(ex))

    def _apply_grouping(self):
        try:
            view = self._CollectionViewSource.GetDefaultView(self.vm.rows)
            if view is None:
                return
            view.GroupDescriptions.Clear()
            path = self.vm.group_key
            if not path:
                item = self.groupCombo.SelectedItem
                path = item.value if item else None
            if path:
                view.GroupDescriptions.Add(self._PropertyGroupDescription("GroupValue"))
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

    def _refresh_grid_cells(self):
        """Rebind dynamic cells after programmatic multi-row updates."""
        try:
            view = self._CollectionViewSource.GetDefaultView(self.vm.rows)
            if view is not None:
                view.Refresh()
        except Exception as ex:
            self._logger.debug("CDE: grid view refresh failed: {}".format(ex))
        try:
            self.doorGrid.InvalidateVisual()
        except Exception:
            pass

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

    def _pick_default_column_defs(self, elements, param_defs, limit):
        """Pick CDE columns that actually have values in the loaded graph."""
        counts = {}
        for element in (elements or []):
            for key, val in element.values.items():
                if val is None or val == "":
                    continue
                counts[key] = counts.get(key, 0) + 1
        if counts:
            by_key = dict((d.key, d) for d in (param_defs or []))
            ranked = sorted(counts.keys(), key=lambda k: (-counts[k], k))
            chosen = []
            for key in ranked:
                pdef = by_key.get(key)
                if pdef is None:
                    continue
                if getattr(pdef, "source", "cde") != "cde":
                    continue
                chosen.append(pdef)
                if len(chosen) >= limit:
                    break
            if chosen:
                return chosen
        function_defs = [
            d for d in (param_defs or [])
            if getattr(d, "group", "") == "Function"]
        if function_defs:
            return function_defs[:limit]
        cde_defs = [
            d for d in (param_defs or [])
            if getattr(d, "source", "cde") == "cde"]
        if cde_defs:
            return cde_defs[:limit]
        return list(param_defs or [])[:limit]

    def _sync_columns(self, elements):
        """Reconcile dynamic columns with the freshly loaded parameter set."""
        current_class = self.vm.ifc_class
        category_changed = current_class != self._columns_ifc_class
        if category_changed:
            for key in list(self._value_column_map.keys()):
                self._remove_value_column(key)
            if self._param_defs:
                chosen = self._pick_default_column_defs(
                    elements, self._param_defs, self._default_columns)
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

    def _row_from_cell_edit(self, args):
        """Resolve ElementRow from DataGridCellEditEndingEventArgs."""
        row_container = args.Row
        if row_container is not None:
            item = getattr(row_container, "Item", None)
            if item is not None and hasattr(item, "record_pending"):
                return item
        editing = args.EditingElement
        if editing is not None:
            ctx = getattr(editing, "DataContext", None)
            if ctx is not None and hasattr(ctx, "record_pending"):
                return ctx
        return None

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
            targets = self._selected_rows()
            if row not in targets:
                targets = [row]
            for target in targets:
                target.record_pending(key, new_value)
            self._refresh_grid_cells()
            self._update_pending_ui()
            if len(targets) > 1:
                self._set_status(
                    "Staged '{}' on {} row(s). Click Apply to commit.".format(
                        key, len(targets)))
            else:
                self._set_status(
                    "Staged '{}'. Click Apply to commit.".format(key))
        except Exception as ex:
            self._logger.debug("CDE: cell edit failed: {}".format(ex))

    # --- coloring -------------------------------------------------------

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

    # --- selection ------------------------------------------------------

    def _selected_rows(self):
        return [r for r in self.doorGrid.SelectedItems]

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
        self._set_status(
            "Staged '{}' on {} row(s). Click Apply to commit.".format(
                pdef.label, len(rows)))

    def on_discard_pending_click(self, sender, args):
        if self._apply_in_flight:
            return
        count = self.vm.pending_count()
        if count == 0:
            return
        if not forms.alert(
                "Discard {} pending edit(s)?".format(count),
                yes=True, no=True):
            return
        self.vm.discard_all_pending()
        self._refresh_grid_cells()
        self._update_pending_ui()
        self._set_status("Discarded {} pending edit(s).".format(count))

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
        changes = self.vm.collect_pending_changes()
        if not changes:
            self._set_status("No pending edits to apply.")
            return
        project_id, revision_id = self._project_revision_ids()
        if not project_id or not revision_id:
            self._set_status("Map the model before writing values.")
            return
        self._set_apply_in_flight(True)
        self._set_status("Running dry run...")

        def work():
            etag = self._get_revision_etag(project_id, revision_id)
            return self.service.apply_element_mutations(
                project_id, revision_id, changes, dry_run=True, etag=etag)

        def done(outcome, error):
            if error is not None:
                self._set_apply_in_flight(False)
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
            if not forms.alert(msg, yes=True, no=True):
                self._set_apply_in_flight(False)
                self._set_status("Apply cancelled.")
                return
            fresh_changes = self.vm.collect_pending_changes()
            if not fresh_changes:
                self._set_apply_in_flight(False)
                self._set_status("No pending edits to commit.")
                return
            self._commit_pending_changes(fresh_changes, project_id, revision_id)

        self._run_async(work, done)

    def _handle_apply_error(self, error):
        if isinstance(error, CDEPreconditionError):
            project_id, revision_id = self._project_revision_ids()
            if project_id and revision_id:
                etag = self.service.fetch_revision_etag(project_id, revision_id)
                if etag:
                    self._save_revision_etag(etag)
            self._set_status(
                "Model changed (stale etag). Refreshing — review and retry Apply.")
            self.refresh(preserve_pending=True)
            return
        if isinstance(error, CDEConflictError):
            self._set_status("Write conflict (409): {}".format(error))
            return
        self._set_status("Apply failed: {}".format(error))

    def _commit_pending_changes(self, changes, project_id, revision_id):
        self._set_apply_in_flight(True)
        self._set_status("Committing to CDE...")

        def work():
            etag = self._get_revision_etag(project_id, revision_id)
            try:
                outcome = self.service.apply_element_mutations(
                    project_id, revision_id, changes, dry_run=False, etag=etag)
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
                self._set_status("Mutation {} — pending edits kept.".format(state))
                return
            if outcome.mutation_id:
                if not state_lower:
                    self._set_status(
                        "Mutation status unknown — pending edits kept.")
                    return
                if state_lower not in self._mutation_success_states:
                    self._set_status(
                        "Mutation state '{}' — pending edits kept.".format(state))
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
        rows = self._selected_rows()
        if rows:
            return rows[0]
        return None

    def _property_rows_for_detail(self, row, cde_values):
        """Build inspector rows: Revit fields plus CDE effective values."""
        rows = [
            self._PropertyRow("Revit.Mark", row.Mark),
            self._PropertyRow("Revit.Level", row.Level),
            self._PropertyRow("Revit.FromRoom", row.FromRoom),
            self._PropertyRow("Revit.ToRoom", row.ToRoom),
            self._PropertyRow("Revit.MatchedInRevit", row.MatchedInRevit),
        ]
        for key, val in sorted((cde_values or {}).items()):
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
        self._sync_revit_selection()
        self._detail_fetch_token += 1
        token = self._detail_fetch_token
        if row is None:
            self._apply_detail_pane(None, None)
            return
        cached = dict(row.cells)
        self._apply_detail_pane(row, cached, "Loading CDE properties...")
        self._fetch_detail_for_row(row, token)

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
                self._watcher.stop()
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
