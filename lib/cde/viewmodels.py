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

    def __init__(self, global_id, revit_info=None, cde_values=None):
        super(ElementRow, self).__init__()
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

    # --- dynamic cells (WPF indexer binding) ---------------------------

    def __getitem__(self, key):
        value = self._cells.get(key)
        return "" if value is None else value

    def __setitem__(self, key, value):
        self._cells[key] = value
        # WPF observes indexer changes through the special "Item[]" property.
        self.OnPropertyChanged("Item[]")

    def get_cell(self, key, default=""):
        value = self._cells.get(key)
        return default if value is None else value

    def set_cell(self, key, value):
        self[key] = value

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

    def set_rows(self, rows):
        self._all_rows = list(rows)
        self.apply_filter()

    def apply_filter(self):
        self.rows.Clear()
        for row in self._all_rows:
            if row.matches_filter(self.filter_text):
                self.rows.Add(row)

    def all_rows(self):
        return list(self._all_rows)

    def selected_global_ids(self, rows):
        return [r.GlobalId for r in rows if r.GlobalId]

    def selected_element_ids(self, rows):
        return [r.element_id for r in rows if r.element_id is not None]
