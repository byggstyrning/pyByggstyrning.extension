# -*- coding: utf-8 -*-
"""CDE door/element schedule - a modeless Revit cockpit over the CDE graph.

Joins CDE elements (by IFC GlobalId) to Revit elements, lists them in a
groupable/filterable table, lets the user add CDE/Revit parameter columns,
color elements in the active view by any value (temporary view graphics),
select doors back into Revit, and vertically set values back to the CDE.

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
from System.Windows.Data import (
    Binding, IValueConverter, CollectionViewSource, PropertyGroupDescription)
from System.Windows.Controls import DataGridTextColumn
from System.Windows.Media import SolidColorBrush, Color as WpfColor

from pyrevit import forms, revit, script, HOST_APP

# --- lib bootstrap --------------------------------------------------------
_pushbutton_dir = op.dirname(__file__)
_extension_dir = op.dirname(op.dirname(op.dirname(_pushbutton_dir)))
_lib_path = op.join(_extension_dir, "lib")
if _lib_path not in sys.path:
    sys.path.insert(0, _lib_path)

from styles import load_styles_to_window
from cde import storage, matching, coloring
from cde.auth import CDEAuthClient
from cde.service import CDEService, MockCDEService
from cde.viewmodels import ElementRow, ScheduleViewModel
from cde.revit_events import ExternalEventRunner, ActiveViewWatcher

logger = script.get_logger()

# Categories the schedule can drive (label -> IfcClass).
CATEGORIES = [("Doors", "IfcDoor"), ("Windows", "IfcWindow"), ("Walls", "IfcWall")]
# Group-by options (label -> ElementRow property path, None = ungrouped).
GROUP_OPTIONS = [("No grouping", None), ("Level", "Level"),
                 ("From room", "FromRoom"), ("To room", "ToRoom")]

# Keep a module reference so the modeless window is not garbage-collected.
__window__ = None


class _ComboItem(object):
    def __init__(self, name, value):
        self.name = name
        self.value = value


class _LegendItem(object):
    def __init__(self, label, rgb):
        self.Label = label
        self.Brush = SolidColorBrush(WpfColor.FromRgb(rgb[0], rgb[1], rgb[2]))


class CellValueConverter(IValueConverter):
    """Resolves a dynamic CDE/Revit cell value from a row + column key."""

    def Convert(self, value, target_type, parameter, culture):
        try:
            return value.get_cell(parameter)
        except Exception:
            return ""

    def ConvertBack(self, value, target_type, parameter, culture):
        return None


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
        self.doorGrid.ItemsSource = self.vm.rows
        self._cell_converter = CellValueConverter()
        self._param_defs = []
        self._revit_infos = {}
        self._runner = ExternalEventRunner()
        self._watcher = ActiveViewWatcher(self.uiapp, self._on_active_view_changed)
        # Capture module refs for ExternalEvent callbacks (pyRevit may dispose script scope).
        self._logger = logger

        self._init_combos()
        self.mapping = storage.load_mapping(self.doc)
        if not self.mapping:
            self._set_status("Model is not mapped. Run 'CDE Login' first "
                             "(or enable offline demo via Login).")

    # --- setup ----------------------------------------------------------

    def _init_combos(self):
        self.categoryCombo.ItemsSource = [_ComboItem(l, c) for l, c in CATEGORIES]
        self.categoryCombo.SelectedIndex = 0
        self.groupCombo.ItemsSource = [_ComboItem(l, p) for l, p in GROUP_OPTIONS]
        self.groupCombo.SelectedIndex = 0

    def _set_status(self, message):
        self.statusText.Text = message

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
                self.Dispatcher.Invoke(Action(lambda: on_done(result, error)))
        thread = Thread(ThreadStart(runner))
        thread.IsBackground = True
        thread.Start()

    # --- refresh orchestration -----------------------------------------

    def on_refresh_click(self, sender, args):
        self.refresh()

    def refresh(self):
        self.vm.ifc_class = self.ifc_class
        self.vm.active_view_only = bool(self.activeViewCheck.IsChecked)
        self._set_status("Reading Revit elements...")
        # Step 1: collect Revit data on the Revit API thread.
        self._runner.run(self._collect_revit)

    def _collect_revit(self, uiapp):
        try:
            doc = uiapp.ActiveUIDocument.Document
            view = doc.ActiveView if self.vm.active_view_only else None
            index = matching.build_guid_index(doc, self.vm.ifc_class, view)
            infos = {}
            for guid, eid in index.items():
                element = doc.GetElement(eid)
                if element is not None:
                    infos[guid] = matching.get_revit_info(doc, element, self.vm.ifc_class)
            self._revit_infos = infos
        except Exception as ex:
            self._logger.error("CDE: collect Revit failed: {}".format(ex))
            self._revit_infos = {}
        # Step 2: hop back to the UI thread to fetch CDE data.
        self.Dispatcher.Invoke(Action(self._fetch_cde))

    def _fetch_cde(self):
        if not self.mapping and not self.offline:
            # Revit-only view: still useful for matching/coloring later.
            self._build_rows(([], []), None)
            return
        project_id = self.mapping["project_id"] if self.mapping else "demo-1"
        revision_id = self.mapping["revision_id"] if self.mapping else "rev-2"
        ifc_class = self.vm.ifc_class
        self._set_status("Fetching CDE elements...")

        def work():
            elements = self.service.list_elements(project_id, revision_id, ifc_class)
            defs = self.service.get_parameter_defs(project_id, ifc_class)
            return elements, defs

        self._run_async(work, self._build_rows)

    def _build_rows(self, payload, error):
        if error is not None:
            self._set_status("CDE fetch failed: {}".format(error))
            payload = ([], [])
        elements, defs = payload
        self._param_defs = defs or []
        self._refresh_param_combo()

        cde_by_guid = {e.global_id: e for e in (elements or [])}
        active_only = self.vm.active_view_only

        rows = []
        guids = set(cde_by_guid.keys()) | set(self._revit_infos.keys())
        for guid in guids:
            info = self._revit_infos.get(guid)
            if active_only and info is None:
                continue  # not visible in the active view
            cde = cde_by_guid.get(guid)
            values = dict(cde.values) if cde else {}
            rows.append(ElementRow(guid, info, values))

        rows.sort(key=lambda r: (r.Level, r.Mark, r.GlobalId))
        self.vm.set_rows(rows)
        self._apply_grouping()
        matched = sum(1 for r in rows if r.MatchedInRevit)
        self._set_status("{} element(s), {} matched in Revit.".format(len(rows), matched))

    def _refresh_param_combo(self):
        items = [_ComboItem(d.label, d) for d in self._param_defs]
        self.paramCombo.ItemsSource = items
        if items:
            self.paramCombo.SelectedIndex = 0

    # --- filtering / grouping ------------------------------------------

    def on_filter_changed(self, sender, args):
        self.vm.filter_text = self.filterBox.Text or ""
        self.vm.apply_filter()
        self._apply_grouping()

    def on_group_changed(self, sender, args):
        self._apply_grouping()

    def _apply_grouping(self):
        try:
            view = CollectionViewSource.GetDefaultView(self.vm.rows)
            if view is None:
                return
            view.GroupDescriptions.Clear()
            item = self.groupCombo.SelectedItem
            path = item.value if item else None
            if path:
                view.GroupDescriptions.Add(PropertyGroupDescription(path))
        except Exception as ex:
            self._logger.debug("CDE: grouping failed: {}".format(ex))

    def on_category_changed(self, sender, args):
        if self.IsLoaded:
            self.refresh()

    def on_active_view_toggled(self, sender, args):
        if bool(self.activeViewCheck.IsChecked):
            self._watcher.start()
        else:
            self._watcher.stop()
        self.refresh()

    def _on_active_view_changed(self, view):
        if bool(self.activeViewCheck.IsChecked):
            self.refresh()

    # --- dynamic columns ------------------------------------------------

    def on_add_column_click(self, sender, args):
        item = self.paramCombo.SelectedItem
        if item is None:
            self._set_status("Select a parameter to add as a column.")
            return
        key = item.value.key
        if key in self.vm.value_columns:
            return
        self.vm.value_columns.append(key)
        column = DataGridTextColumn()
        column.Header = item.value.label
        binding = Binding()
        binding.Converter = self._cell_converter
        binding.ConverterParameter = key
        column.Binding = binding
        column.Width = 110
        self.doorGrid.Columns.Add(column)
        self._set_status("Added column '{}'.".format(item.value.label))

    # --- coloring -------------------------------------------------------

    def _current_rows(self):
        return self.vm.all_rows()

    def on_color_click(self, sender, args):
        item = self.paramCombo.SelectedItem
        if item is None:
            self._set_status("Select a parameter to color by.")
            return
        key = item.value.key
        rows = self._current_rows()
        self._set_status("Applying colors...")

        def task(uiapp):
            doc = uiapp.ActiveUIDocument.Document
            view = doc.ActiveView
            color_map = coloring.apply_coloring(doc, view, rows, key)
            legend = coloring.build_legend(color_map)
            self.Dispatcher.Invoke(Action(lambda: self._render_legend(legend, item.value.label)))

        self._runner.run(task)

    def _render_legend(self, legend, label):
        self.legendList.ItemsSource = [_LegendItem(value, rgb) for value, rgb in legend]
        self._set_status("Colored by '{}'.".format(label))

    def on_reset_color_click(self, sender, args):
        rows = self._current_rows()

        def task(uiapp):
            doc = uiapp.ActiveUIDocument.Document
            coloring.reset_coloring(doc, doc.ActiveView, rows)
            self.Dispatcher.Invoke(Action(lambda: self._clear_legend()))

        self._runner.run(task)

    def _clear_legend(self):
        self.legendList.ItemsSource = None
        self._set_status("Colors reset.")

    # --- selection ------------------------------------------------------

    def _selected_rows(self):
        return [r for r in self.doorGrid.SelectedItems]

    def on_select_click(self, sender, args):
        element_ids = [r.element_id for r in self._selected_rows() if r.element_id is not None]
        if not element_ids:
            self._set_status("No Revit-matched rows selected.")
            return

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
        self._set_status("Selected {} element(s) in Revit.".format(len(element_ids)))

    # --- vertical value set --------------------------------------------

    def on_set_value_click(self, sender, args):
        item = self.paramCombo.SelectedItem
        if item is None:
            self._set_status("Select a parameter to set.")
            return
        rows = self._selected_rows()
        global_ids = [r.GlobalId for r in rows if r.GlobalId]
        if not global_ids:
            self._set_status("Select one or more rows first.")
            return
        key = item.value.key
        value = self.valueBox.Text
        if not self.mapping and not self.offline:
            self._set_status("Map the model before writing values.")
            return
        project_id = self.mapping["project_id"] if self.mapping else "demo-1"
        revision_id = self.mapping["revision_id"] if self.mapping else "rev-2"
        self._set_status("Writing value to CDE...")

        def work():
            return self.service.set_element_values(
                project_id, revision_id, global_ids, {key: value})

        def done(ok, error):
            if error is not None:
                self._set_status("Write failed: {}".format(error))
                return
            # Optimistic local update.
            wanted = set(global_ids)
            for row in rows:
                if row.GlobalId in wanted:
                    row.set_cell(key, value)
            self._set_status("Set '{}' = '{}' on {} element(s).".format(
                item.value.label, value, len(global_ids)))

        self._run_async(work, done)

    # --- lifecycle ------------------------------------------------------

    def window_closing(self, sender, args):
        try:
            self._watcher.stop()
        except Exception:
            pass


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
