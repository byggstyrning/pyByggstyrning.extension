# -*- coding: utf-8 -*-
"""WPF-bindable models for the CDE schedule table.

``ElementRow`` is a ``forms.Reactive`` row carrying fixed Revit-derived columns
plus a dynamic set of CDE/Revit parameter cells (bound in XAML via an indexer,
e.g. ``Path=[needs_lock]``). ``ScheduleViewModel`` owns the master row list and
applies the quick-filter into the bound ``ObservableCollection``.
"""
from System.Collections.ObjectModel import ObservableCollection

from pyrevit import forms


class ElementRow(forms.Reactive):
    """One door/element row joining a CDE element to a Revit element."""

    def __init__(self, global_id, revit_info=None, cde_values=None, vm=None):
        super(ElementRow, self).__init__()
        self._vm = vm
        self._global_id = global_id or ""
        revit_info = revit_info or {}
        self.element_id = revit_info.get("element_id")
        self.element_id_value = revit_info.get("element_id_value")
        self._mark = revit_info.get("mark", "") or ""
        self._level = revit_info.get("level", "") or ""
        self._from_room = revit_info.get("from_room", "") or ""
        self._to_room = revit_info.get("to_room", "") or ""
        self._matched = self.element_id is not None
        # Dynamic parameter cells (CDE values merged with any pulled Revit params).
        self._cells = dict(cde_values or {})
        # Committed baseline for staged revert; updated after successful Apply.
        self._original_cells = dict(cde_values or {})
        # Staged edits keyed by parameter column key.
        self._pending_cells = {}

    # --- fixed columns --------------------------------------------------

    @property
    def GlobalId(self):
        return self._global_id

    @property
    def Mark(self):
        return self._mark

    @property
    def Level(self):
        return self._level

    @property
    def FromRoom(self):
        return self._from_room

    @property
    def ToRoom(self):
        return self._to_room

    @property
    def MatchedInRevit(self):
        return self._matched

    @property
    def IsDirty(self):
        """True when the row has staged pending edits."""
        return bool(self._pending_cells)

    _FIXED_BUCKET = {
        "GlobalId": "GlobalId",
        "Mark": "Mark",
        "Level": "Level",
        "FromRoom": "FromRoom",
        "ToRoom": "ToRoom",
        "MatchedInRevit": "MatchedInRevit",
    }

    @property
    def GroupValue(self):
        """Value used by PropertyGroupDescription for the active group column."""
        key = self._vm.group_key if self._vm else None
        if not key:
            return u""
        fixed = self._FIXED_BUCKET.get(key)
        if fixed:
            value = getattr(self, fixed, "")
            if isinstance(value, bool):
                return value
            value = value or ""
        else:
            value = self.get_cell(key, "")
        return value if value != "" else u"(empty)"

    @property
    def SortValue(self):
        """Value used when sorting dynamic columns (see ``ScheduleViewModel.sort_key``)."""
        key = self._vm.sort_key if self._vm else None
        if not key:
            return u""
        fixed = self._FIXED_BUCKET.get(key)
        if fixed:
            value = getattr(self, fixed, "")
            if isinstance(value, bool):
                return value
            return value or ""
        return self.get_cell(key, "")

    # --- dynamic cells (WPF indexer binding) ---------------------------

    def __getitem__(self, key):
        value = self._cells.get(key)
        return "" if value is None else value

    def __setitem__(self, key, value):
        self._cells[key] = value
        # WPF observes indexer changes through the special "Item[]" property.
        self.OnPropertyChanged("Item[]")

    def get_cell(self, key, default=""):
        fixed = self._FIXED_BUCKET.get(key)
        if fixed:
            value = getattr(self, fixed, default)
            return default if value is None else value
        value = self._cells.get(key)
        return default if value is None else value

    def set_cell(self, key, value):
        self[key] = value
        self.OnPropertyChanged("SortValue")
        self.OnPropertyChanged("GroupValue")

    def is_cell_pending(self, key):
        return key in self._pending_cells

    def record_pending(self, key, new_value):
        """Stage a local edit without persisting to the CDE."""
        baseline = self._original_cells.get(key, self._cells.get(key))
        if new_value == baseline:
            if key in self._pending_cells:
                del self._pending_cells[key]
                self.set_cell(key, baseline if baseline is not None else "")
                self.OnPropertyChanged("IsDirty")
                if self._vm is not None:
                    self._vm.notify_pending_changed()
            return
        if key not in self._pending_cells:
            self._original_cells.setdefault(key, self._cells.get(key))
        self._pending_cells[key] = new_value
        self.set_cell(key, new_value)
        self.OnPropertyChanged("IsDirty")
        if self._vm is not None:
            self._vm.notify_pending_changed()

    def pending_for(self):
        return dict(self._pending_cells)

    def revert_pending(self):
        """Discard staged edits and restore committed cell values."""
        for key in list(self._pending_cells.keys()):
            orig = self._original_cells.get(key, "")
            self.set_cell(key, orig)
        self._pending_cells.clear()
        self.OnPropertyChanged("IsDirty")
        if self._vm is not None:
            self._vm.notify_pending_changed()

    def commit_pending(self, keys=None):
        """Mark staged edits as committed (after successful CDE write)."""
        if keys is None:
            keys = list(self._pending_cells.keys())
        for key in keys:
            if key in self._pending_cells:
                self._original_cells[key] = self._pending_cells[key]
                del self._pending_cells[key]
        self.OnPropertyChanged("IsDirty")
        if self._vm is not None:
            self._vm.notify_pending_changed()

    def reset_baseline(self, cde_values):
        """Replace committed baseline after a full refresh from the CDE."""
        self._cells = dict(cde_values or {})
        self._original_cells = dict(cde_values or {})
        self._pending_cells.clear()
        self.OnPropertyChanged("Item[]")
        self.OnPropertyChanged("IsDirty")
        self.OnPropertyChanged("SortValue")
        self.OnPropertyChanged("GroupValue")

    @property
    def cells(self):
        return self._cells

    def matches_filter(self, text):
        """Case-insensitive substring match across all visible content."""
        if not text:
            return True
        needle = text.lower()
        haystack = [self._global_id, self._mark, self._level,
                    self._from_room, self._to_room]
        haystack.extend(unicode(v) for v in self._cells.values())
        return any(needle in (h or "").lower() for h in haystack)


class ScheduleViewModel(object):
    """Holds all rows and projects the filtered subset into the bound grid."""

    def __init__(self):
        self.rows = ObservableCollection[ElementRow]()
        self._all_rows = []
        self.ifc_class = "IfcDoor"
        self.active_view_only = False
        self.filter_text = ""
        # Parameter keys currently shown as dynamic columns.
        self.value_columns = []
        # Active group / sort column (fixed property or dynamic cell key).
        self.group_key = None
        self.sort_key = None
        self._pending_change_handlers = []

    def on_pending_changed(self, handler):
        self._pending_change_handlers.append(handler)

    def notify_pending_changed(self):
        for handler in self._pending_change_handlers:
            try:
                handler()
            except Exception:
                pass

    def set_rows(self, rows):
        self._all_rows = list(rows)
        self.apply_filter()

    def apply_filter(self):
        filtered = ObservableCollection[ElementRow]()
        needle = self.filter_text
        for row in self._all_rows:
            if row.matches_filter(needle):
                filtered.Add(row)
        self.rows = filtered

    def all_rows(self):
        return list(self._all_rows)

    def selected_global_ids(self, rows):
        return [r.GlobalId for r in rows if r.GlobalId]

    def selected_element_ids(self, rows):
        return [r.element_id for r in rows if r.element_id is not None]

    def pending_count(self):
        total = 0
        for row in self._all_rows:
            total += len(row.pending_for())
        return total

    def collect_pending_changes(self):
        """Return {global_id: {param_key: staged_value}} for Apply."""
        changes = {}
        for row in self._all_rows:
            pending = row.pending_for()
            if pending and row.GlobalId:
                changes[row.GlobalId] = dict(pending)
        return changes

    def discard_all_pending(self):
        for row in self._all_rows:
            if row.IsDirty:
                row.revert_pending()

    def clear_pending_for_global_ids(self, global_ids):
        wanted = set(global_ids or [])
        for row in self._all_rows:
            if row.GlobalId in wanted:
                row.commit_pending()

    def rows_with_pending(self):
        return [r for r in self._all_rows if r.IsDirty]

    def apply_pending_snapshot(self, snapshot):
        """Re-stage pending edits onto rows after a grid rebuild (by GlobalId)."""
        if not snapshot:
            return
        by_gid = dict((r.GlobalId, r) for r in self._all_rows if r.GlobalId)
        for gid, pending in snapshot.items():
            row = by_gid.get(gid)
            if row is None:
                continue
            for key, value in pending.items():
                row.record_pending(key, value)
