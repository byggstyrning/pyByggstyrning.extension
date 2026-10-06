# -*- coding: utf-8 -*-
__title__ = "Create References"
__author__ = "Jonatan Jacobsson"
__doc__ = """Places a 3D View Reference family at the location and extent of the selected views.
Views that already have a reference get it updated instead of getting a second one.
"""

# Import .NET libraries
import clr
clr.AddReference("System")
clr.AddReference("WindowsBase")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from System import Action, Double
from System.Windows import CornerRadius, FontWeights, HorizontalAlignment, Size, Thickness, VerticalAlignment, Visibility
from System.Windows.Documents import Run
from System.ComponentModel import ListSortDirection, SortDescription
from System.Windows.Controls import (
    Border, Button, CheckBox, ComboBox, ComboBoxItem, DataGridRow, Grid, Panel, TextBlock, TextBox,
)
from System.Windows.Input import Cursors, Key
from System.Windows.Threading import DispatcherPriority
from System.Windows.Media import Brushes, FontFamily
from System.Windows.Data import CollectionViewSource
from System.Windows.Media import VisualTreeHelper

# Import Revit API
from Autodesk.Revit.DB import ElementId

# Import pyRevit libraries
import json
import os
import sys
import os.path as op
from pyrevit import revit
from pyrevit import forms, script

# Add the extension directory to the path
script_path = __file__
pushbutton_dir = op.dirname(script_path)
stack_dir = op.dirname(pushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from revit import view_references
from revit.compat import get_element_id_value

# Get the current script directory (for XAML file path)
script_dir = pushbutton_dir
logger = script.get_logger()

# Get Revit document
doc = __revit__.ActiveUIDocument.Document

# Filter dropdown choices
SHEET_ALL, SHEET_ON, SHEET_OFF = "On sheet or not", "On sheet", "Not on sheet"
SHEET_FILTERS = [SHEET_ALL, SHEET_ON, SHEET_OFF]
REFERENCE_ALL, REFERENCE_PLACED, REFERENCE_MISSING = "Placed or not", "Placed", "Not placed"
REFERENCE_FILTERS = [REFERENCE_ALL, REFERENCE_PLACED, REFERENCE_MISSING]
NO_SHEET_PARAMETER = "(no sheet parameter)"
CATEGORY_ALL = "All categories"


class PartTag(object):
    """Which formula and which part a control belongs to."""

    def __init__(self, key, index):
        self.key = key
        self.index = index


class ViewItemData(forms.Reactive):
    """Class for view data binding with WPF UI."""

    def __init__(self, view, sheet, has_reference):
        """Initialize with a Revit view and the sheet it is placed on, if any."""
        super(ViewItemData, self).__init__()
        self.view = view
        self.sheet = sheet
        self.kind = view_references.get_view_kind(view)
        self.has_reference = has_reference
        self._is_selected = False
        self._sheet_parameter_value = ""
        self.view_name = view.Name
        self.view_category = view_references.get_view_kind_label(view)
        self.view_scale = "1:{}".format(view.Scale) if view.Scale else "Unknown"
        self.area = view_references.get_view_area(view)  # m2, None without a crop box
        self.sheet_reference = view_references.get_sheet_label(sheet) if sheet else "Not on sheet"
        self.reference_status = "Placed" if has_reference else ""

    def matches(self, words):
        """True if every search word occurs somewhere in the row's texts."""
        text = u" ".join([self.view_name, self.view_category, self.view_scale,
                          self.sheet_reference, self.reference_status,
                          self._sheet_parameter_value]).lower()
        return all(word in text for word in words)

    @property
    def SheetParameterValue(self):
        return self._sheet_parameter_value

    @SheetParameterValue.setter
    def SheetParameterValue(self, value):
        self._sheet_parameter_value = value
        self.OnPropertyChanged("SheetParameterValue")

    @property
    def ViewName(self):
        return self.view_name

    @property
    def ViewCategory(self):
        return self.view_category

    @property
    def ViewScale(self):
        return self.view_scale

    @property
    def ViewArea(self):
        return "" if self.area is None else "{:.2f}".format(self.area)

    @property
    def ViewAreaValue(self):
        """What the Area column sorts on."""
        return -1.0 if self.area is None else self.area

    @property
    def ViewScaleValue(self):
        """What the View Scale column sorts on: 1:50 before 1:100."""
        return self.view.Scale or 0

    @property
    def SheetReference(self):
        return self.sheet_reference

    @property
    def ReferenceStatus(self):
        return self.reference_status

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        self.OnPropertyChanged("IsSelected")


class Generate3DViewReferencesWindow(forms.WPFWindow):
    """WPF window for selecting and generating 3D view references."""

    def __init__(self):
        """Initialize the window."""
        xaml_file = os.path.join(script_dir, "Generate3DViewReferencesWindow.xaml")
        forms.WPFWindow.__init__(self, xaml_file)

        # Load styles AFTER window initialization (window-scoped, does not affect Revit UI)
        from styles import load_styles_to_window
        load_styles_to_window(self)

        # Store created elements for isolation
        self.created_elements = []

        # All rows, and the rows that pass the filters (what the grid shows)
        self.all_items = None
        self.views_data = ObservableCollection[ViewItemData]()

        self.family_symbol = view_references.find_family_symbol(doc)
        if not self.family_symbol:
            self._alert_family_missing()

        # Set by "Go to view/sheet": the view to open once this dialog has closed
        self.go_to_element = None
        self.context_item = None

        self._setup_filters()
        saved_state = self._read_state()
        self._restore_filters(saved_state)
        self._setup_formulas()
        self._load_views()

        # Bind views to DataGrid
        self.viewsDataGrid.ItemsSource = self.views_data
        self._restore_rows(saved_state)
        self._refresh_formula_examples()
        self.Closed += self._save_state
        logger.debug("UI setup complete. Found {} views.".format(self.views_data.Count))

    def _alert_family_missing(self):
        forms.alert(
            "This tool requires the '{}' family.\n\n"
            "Load it with the Load Family button and try again.".format(
                view_references.FAMILY_NAME),
            title="Required Family Not Found"
        )

    def set_busy(self, is_busy, message="Loading..."):
        """Show or hide the busy overlay indicator."""
        try:
            if is_busy:
                self.busyOverlay.Visibility = Visibility.Visible
                self.busyTextBlock.Text = message
            else:
                self.busyOverlay.Visibility = Visibility.Collapsed
        except Exception as e:
            logger.debug("Error setting busy indicator: {}".format(str(e)))

    def _setup_filters(self):
        """Fill the filter dropdowns. The sheet parameter choice is remembered."""
        self._formula_labels = {}
        self._formula_kept = {}
        self._formula_part_tags = {}
        self._formula_search_boxes = {}
        self._formula_search_wired = set()
        labels = [CATEGORY_ALL] + [label for _kind, label in view_references.VIEW_KINDS]
        self.categoryFilterComboBox.ItemsSource = List[str](labels)
        self.categoryFilterComboBox.SelectedIndex = 0
        self._formula_labels[id(self.categoryFilterComboBox)] = list(labels)
        self.categoryFilterComboBox.DropDownOpened += self.FormulaDropDown_Opened
        self.categoryFilterComboBox.DropDownClosed += self.FormulaDropDown_Closed
        self.sheetFilterComboBox.ItemsSource = List[str](SHEET_FILTERS)
        self.sheetFilterComboBox.SelectedIndex = 0
        self.referenceFilterComboBox.ItemsSource = List[str](REFERENCE_FILTERS)
        self.referenceFilterComboBox.SelectedIndex = 0

        self.manualDepthTextBox.Text = script.get_config().get_option("manual_depth_mm", "")
        self.area_filter_error = False
        self.maxAreaTextBox.Text = script.get_config().get_option("max_area_m2", "")
        self.Closed += self._save_max_area

        names = [NO_SHEET_PARAMETER] + view_references.get_sheet_parameter_names(doc)
        self.sheetParameterComboBox.ItemsSource = List[str](names)
        remembered = script.get_config().get_option("sheet_parameter", NO_SHEET_PARAMETER)
        self.sheetParameterComboBox.SelectedItem = (
            remembered if remembered in names else NO_SHEET_PARAMETER)

    def _manual_depth(self):
        """The typed depth in feet, None if the box is empty. ValueError if not a number."""
        text = (self.manualDepthTextBox.Text or "").strip().replace(",", ".")
        if not text:
            return None
        millimetres = float(text)
        if millimetres <= 0:
            raise ValueError(text)
        return millimetres / 304.8

    def _max_area(self):
        """The typed area limit in m2, None if empty. ValueError if not a number."""
        text = (self.maxAreaTextBox.Text or "").strip().replace(",", ".")
        if not text:
            return None
        area = float(text)
        if area <= 0:
            raise ValueError(text)
        return area

    def _save_max_area(self, sender, args):
        script.get_config().set_option("max_area_m2", (self.maxAreaTextBox.Text or "").strip())
        script.save_config()

    def _setup_formulas(self):
        """Sheet Number and View Name, each built left to right."""
        self._formula_loading = True
        self.editors = {
            "view": self._make_editor(
                "view",
                view_references.load_view_name_formula(doc),
                view_references.view_name_choice_list(doc),
                self.viewFormulaPanel,
                self.viewFormulaExample,
                view_references.SOURCE_VIEW_NAME,
                (view_references.SOURCE_VIEW, view_references.SOURCE_PROJECT,
                 view_references.SOURCE_SHEET),
                "Add a view, sheet or project parameter in front of the view name"),
            "sheet": self._make_editor(
                "sheet",
                view_references.load_sheet_number_formula(doc),
                view_references.sheet_number_choice_list(doc),
                self.sheetFormulaPanel,
                self.sheetFormulaExample,
                view_references.SOURCE_SHEET_NUMBER,
                (view_references.SOURCE_PROJECT, view_references.SOURCE_SHEET),
                "Add a sheet or project parameter in front of the sheet number"),
        }
        self._formula_loading = False
        self._rebuild_formula_panel("view")
        self._rebuild_formula_panel("sheet")

    def _make_editor(self, key, formula, choices, panel, example, default_source,
                     prefer, add_tooltip):
        return {
            "key": key,
            "parts": formula["parts"],
            "choices": choices,
            "panel": panel,
            "example": example,
            "default_source": default_source,
            "prefer": prefer,
            "add_tooltip": add_tooltip,
        }

    def _same_part(self, left, right):
        if left.get("source") != right.get("source"):
            return False
        if left.get("source") in (
                view_references.SOURCE_SHEET_NUMBER, view_references.SOURCE_VIEW_NAME):
            return True
        return left.get("name") == right.get("name")

    def _label_for_part(self, editor, part):
        for label, candidate in editor["choices"]:
            if self._same_part(candidate, part):
                return label
        return view_references.sheet_number_part_label(part)

    def _part_for_label(self, editor, label):
        for candidate_label, part in editor["choices"]:
            if candidate_label == label:
                item = {"source": part.get("source")}
                if part.get("name"):
                    item["name"] = part.get("name")
                return item
        return None

    def _formula_dict(self, key):
        return {"parts": [dict(part) for part in self.editors[key]["parts"]]}

    def _save_formulas(self):
        if getattr(self, "_formula_loading", False):
            return
        try:
            view_references.save_view_name_formula(doc, self._formula_dict("view"))
            view_references.save_sheet_number_formula(doc, self._formula_dict("sheet"))
        except Exception as ex:
            logger.debug("Could not save name formulas: {}".format(ex))

    def _part_to_add(self, editor):
        """The first unused preferred parameter, or another copy of the default part."""
        used = set((part.get("source"), part.get("name")) for part in editor["parts"])
        for source in editor["prefer"]:
            for _label, part in editor["choices"]:
                if part.get("source") != source:
                    continue
                key = (part.get("source"), part.get("name"))
                if key in used:
                    continue
                item = {"source": part.get("source")}
                if part.get("name"):
                    item["name"] = part.get("name")
                return item
        return {"source": editor["default_source"]}

    def _show_remove(self, sender, args):
        button = sender.Tag
        if button is not None:
            button.Visibility = Visibility.Visible

    def _hide_remove(self, sender, args):
        button = sender.Tag
        if button is not None:
            button.Visibility = Visibility.Collapsed

    def _remove_hover_on(self, sender, args):
        sender.SetResourceReference(Border.BackgroundProperty, "ErrorBrush")
        if sender.Child is not None:
            sender.Child.Foreground = Brushes.White

    def _remove_hover_off(self, sender, args):
        sender.Background = Brushes.Transparent
        if sender.Child is not None:
            sender.Child.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush")

    def _rebuild_formula_panel(self, key):
        editor = self.editors[key]
        panel = editor["panel"]
        self._formula_loading = True
        try:
            panel.Children.Clear()
            labels = List[str]([label for label, part in editor["choices"]])
            parts = editor["parts"]
            last = len(parts) - 1
            control_height = 33
            for index, part in enumerate(parts):
                host = Grid()
                host.Width = 240
                host.Height = control_height
                host.Margin = Thickness(0, 0, 4, 0)
                host.VerticalAlignment = VerticalAlignment.Center
                combo = ComboBox()
                self._formula_labels[id(combo)] = [label for label, _part in editor["choices"]]
                combo.Margin = Thickness(0)
                combo.VerticalAlignment = VerticalAlignment.Center
                combo.MaxDropDownHeight = 320
                combo.SetResourceReference(ComboBox.StyleProperty, "SearchableComboBoxStyle")
                combo.ItemsSource = labels
                combo.SelectedItem = self._label_for_part(editor, part)
                combo.DropDownOpened += self.FormulaDropDown_Opened
                combo.DropDownClosed += self.FormulaDropDown_Closed
                combo.SelectionChanged += self.FormulaPart_Changed
                self._formula_part_tags[id(combo)] = PartTag(key, index)
                host.Children.Add(combo)
                if len(parts) > 1:
                    mark = TextBlock()
                    mark.Text = u"\u00d7"
                    mark.FontSize = 11
                    mark.HorizontalAlignment = HorizontalAlignment.Center
                    mark.VerticalAlignment = VerticalAlignment.Center
                    mark.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush")
                    remove = Border()
                    remove.Child = mark
                    remove.Width = 16
                    remove.Height = 16
                    remove.Margin = Thickness(0, 1, 1, 0)
                    remove.CornerRadius = CornerRadius(2)
                    remove.Background = Brushes.Transparent
                    remove.HorizontalAlignment = HorizontalAlignment.Right
                    remove.VerticalAlignment = VerticalAlignment.Top
                    remove.Visibility = Visibility.Collapsed
                    remove.Cursor = Cursors.Hand
                    remove.Tag = PartTag(key, index)
                    remove.ToolTip = "Remove this part"
                    remove.MouseEnter += self._remove_hover_on
                    remove.MouseLeave += self._remove_hover_off
                    remove.MouseLeftButtonUp += self.RemoveFormulaPart_Click
                    Panel.SetZIndex(remove, 2)
                    host.Children.Add(remove)
                    host.Tag = remove
                    host.MouseEnter += self._show_remove
                    host.MouseLeave += self._hide_remove
                panel.Children.Add(host)
                if index < last:
                    box = TextBox()
                    box.Height = control_height
                    box.MinHeight = control_height
                    box.MinWidth = 28
                    box.Margin = Thickness(0, 0, 4, 0)
                    box.VerticalAlignment = VerticalAlignment.Center
                    box.VerticalContentAlignment = VerticalAlignment.Center
                    box.Tag = PartTag(key, index)
                    box.Text = part.get("separator") or ""
                    box.ToolTip = "Text between this parameter and the next. Leave empty to join them."
                    box.SetResourceReference(TextBox.StyleProperty, "PlaceholderTextBoxStyle")
                    self._fit_separator(box)
                    box.TextChanged += self.FormulaSeparator_Changed
                    box.LostFocus += self.FormulaSeparator_LostFocus
                    panel.Children.Add(box)
            add = Button()
            add.Content = "+"
            add.Width = control_height
            add.Height = control_height
            add.MinHeight = control_height
            add.Margin = Thickness(0)
            add.VerticalAlignment = VerticalAlignment.Center
            add.ToolTip = editor["add_tooltip"]
            add.Tag = key
            add.SetResourceReference(Button.StyleProperty, "StandardButtonStyle")
            add.Click += self.AddFormulaPart_Click
            panel.Children.Add(add)
        finally:
            self._formula_loading = False

    def _fit_separator(self, box):
        """Grow or shrink the separator box so the whole text stays visible."""
        text = box.Text or u""
        size = box.FontSize
        if Double.IsNaN(size) or size <= 0:
            size = 12.0
        measured = 0.0
        try:
            probe = TextBlock()
            probe.Text = text if text else u" "
            probe.FontFamily = box.FontFamily or FontFamily("Segoe UI")
            probe.FontSize = size
            probe.FontStyle = box.FontStyle
            probe.FontWeight = box.FontWeight
            probe.Measure(Size(Double.PositiveInfinity, Double.PositiveInfinity))
            measured = probe.DesiredSize.Width
        except Exception:
            measured = 0.0
        # PlaceholderTextBoxStyle pads 8 on each side; keep the caret inside too.
        # The character count is a floor so a failed measure cannot clip the text.
        width = max(36.0, measured + 32.0, 28.0 + len(text) * size)
        box.MinWidth = width
        box.Width = width

    def _part_tag(self, sender):
        tag = self._formula_part_tags.get(id(sender))
        if isinstance(tag, PartTag):
            return tag
        tag = sender.Tag
        if isinstance(tag, PartTag):
            return tag
        return None

    def FormulaDropDown_Opened(self, sender, args):
        """Open the type-to-search field used by the searchable combo."""
        self._formula_kept[id(sender)] = sender.SelectedItem
        sender.Dispatcher.BeginInvoke(
            DispatcherPriority.Loaded,
            Action(lambda combo=sender: self._focus_formula_search(combo)))

    def FormulaDropDown_Closed(self, sender, args):
        """Put the full list back. A search that hid the current part does not clear it."""
        labels = self._formula_labels.get(id(sender))
        if not labels:
            return
        selected = sender.SelectedItem or self._formula_kept.get(id(sender))
        self._formula_loading = True
        try:
            sender.ItemsSource = List[str](labels)
            if selected:
                sender.SelectedItem = selected
        finally:
            self._formula_loading = False

    def _focus_formula_search(self, combo):
        box = self._find_template_child(combo, "SearchTextBox")
        if box is None:
            return
        if id(box) not in self._formula_search_wired:
            self._formula_search_boxes[id(box)] = combo
            box.TextChanged += self.FormulaSearch_Changed
            box.PreviewKeyDown += self.FormulaSearch_PreviewKeyDown
            self._formula_search_wired.add(id(box))
        else:
            self._formula_search_boxes[id(box)] = combo
        box.Text = ""
        box.Focus()

    def _find_template_child(self, combo, name):
        if combo.Template is None:
            return None
        found = combo.Template.FindName(name, combo)
        if found is not None:
            return found
        popup = combo.Template.FindName("Popup", combo)
        if popup is None or popup.Child is None:
            return None
        return self._find_named_child(popup.Child, name)

    def _find_named_child(self, parent, name):
        if parent is None:
            return None
        if getattr(parent, "Name", None) == name:
            return parent
        count = VisualTreeHelper.GetChildrenCount(parent)
        for index in range(count):
            found = self._find_named_child(VisualTreeHelper.GetChild(parent, index), name)
            if found is not None:
                return found
        return None

    def FormulaSearch_Changed(self, sender, args):
        combo = self._formula_search_boxes.get(id(sender))
        labels = self._formula_labels.get(id(combo)) if combo is not None else None
        if not labels:
            return
        search = (sender.Text or "").strip().lower()
        selected = combo.SelectedItem
        if search:
            filtered = [label for label in labels if search in label.lower()]
        else:
            filtered = list(labels)
        self._formula_loading = True
        try:
            combo.ItemsSource = List[str](filtered)
            if selected in filtered:
                combo.SelectedItem = selected
        finally:
            self._formula_loading = False
        combo.Dispatcher.BeginInvoke(
            DispatcherPriority.Input,
            Action(lambda box=sender: box.Focus()))

    def FormulaSearch_PreviewKeyDown(self, sender, args):
        combo = self._formula_search_boxes.get(id(sender))
        if combo is None:
            return
        if args.Key == Key.Down:
            args.Handled = True
            combo.Dispatcher.BeginInvoke(
                DispatcherPriority.Input,
                Action(lambda box=combo: self._focus_formula_item(box, True)))
        elif args.Key == Key.Up:
            args.Handled = True
            combo.Dispatcher.BeginInvoke(
                DispatcherPriority.Input,
                Action(lambda box=combo: self._focus_formula_item(box, False)))
        elif args.Key == Key.Enter and combo.SelectedIndex >= 0:
            args.Handled = True
            combo.IsDropDownOpen = False
        elif args.Key == Key.Escape:
            args.Handled = True
            combo.IsDropDownOpen = False

    def _focus_formula_item(self, combo, first):
        popup = combo.Template.FindName("Popup", combo) if combo.Template else None
        if popup is None or popup.Child is None:
            return
        items = []
        self._collect_combo_items(popup.Child, items)
        if not items:
            return
        item = items[0] if first else items[-1]
        item.Focus()
        combo.SelectedIndex = 0 if first else combo.Items.Count - 1

    def _collect_combo_items(self, parent, items):
        if parent is None:
            return
        if isinstance(parent, ComboBoxItem):
            items.append(parent)
        count = VisualTreeHelper.GetChildrenCount(parent)
        for index in range(count):
            self._collect_combo_items(VisualTreeHelper.GetChild(parent, index), items)

    def FormulaPart_Changed(self, sender, args):
        if self._formula_loading:
            return
        tag = self._part_tag(sender)
        editor = self.editors.get(tag.key) if tag is not None else None
        if editor is None or tag.index >= len(editor["parts"]):
            return
        part = self._part_for_label(editor, sender.SelectedItem)
        if part is None:
            return
        if "separator" in editor["parts"][tag.index]:
            part["separator"] = editor["parts"][tag.index].get("separator") or ""
        editor["parts"][tag.index] = part
        self._save_formulas()
        self._refresh_formula_examples()

    def FormulaSeparator_Changed(self, sender, args):
        """Refresh the example on each keystroke. The model is saved when the box is left."""
        if self._formula_loading:
            return
        tag = sender.Tag
        editor = self.editors.get(tag.key) if tag is not None else None
        if editor is None or tag.index >= len(editor["parts"]):
            return
        editor["parts"][tag.index]["separator"] = sender.Text or ""
        self._fit_separator(sender)
        self._refresh_formula_examples()

    def FormulaSeparator_LostFocus(self, sender, args):
        tag = sender.Tag
        editor = self.editors.get(tag.key) if tag is not None else None
        if self._formula_loading or editor is None or tag.index >= len(editor["parts"]):
            return
        editor["parts"][tag.index]["separator"] = sender.Text or ""
        self._save_formulas()
        self._refresh_formula_examples()

    def AddFormulaPart_Click(self, sender, args):
        editor = self.editors.get(sender.Tag)
        if editor is None:
            return
        part = self._part_to_add(editor)
        parts = editor["parts"]
        if parts and parts[-1].get("source") == editor["default_source"]:
            part["separator"] = ""
            parts.insert(len(parts) - 1, part)
        else:
            if parts:
                parts[-1]["separator"] = parts[-1].get("separator") or ""
            parts.append(part)
        self._rebuild_formula_panel(editor["key"])
        self._save_formulas()
        self._refresh_formula_examples()

    def RemoveFormulaPart_Click(self, sender, args):
        tag = sender.Tag
        editor = self.editors.get(tag.key) if tag is not None else None
        if editor is None:
            return
        if 0 <= tag.index < len(editor["parts"]):
            del editor["parts"][tag.index]
        if not editor["parts"]:
            editor["parts"] = [{"source": editor["default_source"]}]
        self._rebuild_formula_panel(editor["key"])
        self._save_formulas()
        self._refresh_formula_examples()

    def _refresh_formula_examples(self):
        view = None
        sheet = None
        for item in self.all_items or []:
            if view is None:
                view = item.view
            if item.sheet is not None:
                sheet = item.sheet
                view = item.view
                break
        self._set_formula_example("view", view, sheet)
        self._set_formula_example("sheet", view, sheet)

    def _set_formula_example(self, key, view, sheet):
        editor = self.editors[key]
        if key == "view":
            if view is None:
                editor["example"].Text = "The view's own name, left to right."
                return
            text = view_references.compose_view_name(
                doc, view, sheet, self._formula_dict("view"))
            self._show_example(editor["example"], text)
            return
        if sheet is None:
            editor["example"].Text = (
                "Left to right. Choose sheet parameters or project information.")
            return
        text = view_references.compose_sheet_number(
            doc, sheet, self._formula_dict("sheet"))
        self._show_example(editor["example"], text)

    def _show_example(self, block, text):
        """Show 'Example:' in regular weight and the composed value in bold."""
        block.Inlines.Clear()
        label = Run(u"Example: ")
        label.FontWeight = FontWeights.Normal
        value = Run(text or u"-")
        value.FontWeight = FontWeights.Bold
        block.Inlines.Add(label)
        block.Inlines.Add(value)



    def _sheet_parameter_name(self):
        name = self.sheetParameterComboBox.SelectedItem
        return None if not name or name == NO_SHEET_PARAMETER else name

    def _load_views(self):
        """Read all views from the model, keeping which rows were unticked."""
        ticked = set(item.view.UniqueId for item in (self.all_items or []) if item.IsSelected)
        sheet_lookup = view_references.build_sheet_lookup(doc)
        existing = view_references.find_existing_references(doc)

        self.all_items = []
        for view in sorted(view_references.collect_views(doc), key=lambda v: v.Name):
            sheet = sheet_lookup.get(get_element_id_value(view.Id))
            item = ViewItemData(view, sheet, view.UniqueId in existing)
            item.IsSelected = view.UniqueId in ticked
            self.all_items.append(item)
        self._update_sheet_parameter_column()
        self._apply_filters()
        if getattr(self, "editors", None):
            self._refresh_formula_examples()

    def _update_sheet_parameter_column(self):
        name = self._sheet_parameter_name()
        self.sheetParameterColumn.Header = name or "Sheet Parameter"
        self.sheetParameterColumn.Visibility = (
            Visibility.Visible if name else Visibility.Collapsed)
        for item in self.all_items:
            item.SheetParameterValue = view_references.get_parameter_text(item.sheet, name)

    def _apply_filters(self):
        """Show the rows that pass the category, search, sheet and reference filters."""
        kinds = self._selected_kinds()
        words = (self.searchTextBox.Text or "").lower().split()
        sheet_filter = self.sheetFilterComboBox.SelectedItem
        reference_filter = self.referenceFilterComboBox.SelectedItem
        try:
            max_area = self._max_area()
            self.area_filter_error = False
        except ValueError:
            max_area = None
            self.area_filter_error = True

        self.views_data.Clear()
        for item in self.all_items:
            if item.kind not in kinds or not item.matches(words):
                continue
            if sheet_filter == SHEET_ON and item.sheet is None:
                continue
            if sheet_filter == SHEET_OFF and item.sheet is not None:
                continue
            if reference_filter == REFERENCE_PLACED and not item.has_reference:
                continue
            if reference_filter == REFERENCE_MISSING and item.has_reference:
                continue
            # a view without a crop box has no area and cannot get a reference anyway
            if max_area is not None and (item.area is None or item.area > max_area + 1e-6):
                continue
            self.views_data.Add(item)
        self._update_status()

    def _update_status(self):
        text = "Showing {} of {} views".format(self.views_data.Count, len(self.all_items))
        headers = dict((c.SortMemberPath, c.Header) for c in self.viewsDataGrid.Columns)
        levels = [u"{} {}".format(headers.get(d.PropertyName, d.PropertyName),
                                  u"\u2191" if d.Direction == ListSortDirection.Ascending else u"\u2193")
                  for d in self._sort_descriptions()]
        if levels:
            text += u"  \u00b7  Sorted by " + u", then ".join(levels)
        if self.area_filter_error:
            text += u"  \u00b7  Max area is not a number, not applied"
        self.countTextBlock.Text = text

    def _sort_descriptions(self):
        return CollectionViewSource.GetDefaultView(self.views_data).SortDescriptions

    def ViewsDataGrid_Sorting(self, sender, args):
        """Header clicks stack: each new column sorts within the ones clicked before it.

        First click adds the column ascending, the second makes it descending,
        the third takes it out of the sort again.
        """
        args.Handled = True
        column = args.Column
        path = column.SortMemberPath
        descriptions = self._sort_descriptions()
        index = -1
        for i, description in enumerate(descriptions):
            if description.PropertyName == path:
                index = i
                break

        if index < 0:
            descriptions.Add(SortDescription(path, ListSortDirection.Ascending))
            column.SortDirection = ListSortDirection.Ascending
        elif descriptions[index].Direction == ListSortDirection.Ascending:
            descriptions[index] = SortDescription(path, ListSortDirection.Descending)
            column.SortDirection = ListSortDirection.Descending
        else:
            descriptions.RemoveAt(index)
            column.SortDirection = None
        self._update_status()

    def _selected_kinds(self):
        """The view categories the category dropdown is showing."""
        label = self.categoryFilterComboBox.SelectedItem
        if not label or label == CATEGORY_ALL:
            return set(kind for kind, _kind_label in view_references.VIEW_KINDS)
        for kind, kind_label in view_references.VIEW_KINDS:
            if kind_label == label:
                return set([kind])
        return set(kind for kind, _kind_label in view_references.VIEW_KINDS)

    def Filter_Changed(self, sender, args):
        """Search text or a filter dropdown changed."""
        if getattr(self, "_formula_loading", False):
            return
        if self.all_items is not None:
            self._apply_filters()

    def SheetParameter_Changed(self, sender, args):
        """Another sheet parameter was picked for the extra column."""
        if self.all_items is None:
            return
        config = script.get_config()
        config.set_option("sheet_parameter", self._sheet_parameter_name() or NO_SHEET_PARAMETER)
        script.save_config()
        self._update_sheet_parameter_column()
        self._apply_filters()

    def SelectAll_Checked(self, sender, args):
        """Handle select all checkbox checked."""
        for view_data in self.views_data:
            view_data.IsSelected = True

    def SelectAll_Unchecked(self, sender, args):
        """Handle select all checkbox unchecked."""
        for view_data in self.views_data:
            view_data.IsSelected = False

    def ViewsDataGrid_PreviewMouseLeftButtonDown(self, sender, args):
        """Ticking one checkbox inside a multi-row selection ticks every selected row.

        Handled on mouse down: a plain click inside a multi-selection collapses the
        selection to the clicked row before the checkbox gets its Click.
        """
        checkbox = self._find_checkbox(args.OriginalSource)
        if checkbox is None:
            return
        if self._toggle_selected_rows(checkbox.DataContext):
            args.Handled = True

    def _find_checkbox(self, element):
        return self._find_ancestor(element, CheckBox)

    def _find_ancestor(self, element, element_type):
        while element is not None:
            if isinstance(element, element_type):
                return element
            try:
                element = VisualTreeHelper.GetParent(element)
            except Exception:
                return None  # not a visual, e.g. a text run
        return None

    def ViewsDataGrid_PreviewMouseRightButtonDown(self, sender, args):
        """Remember which row the context menu is opened on."""
        row = self._find_ancestor(args.OriginalSource, DataGridRow)
        self.context_item = row.Item if row is not None else None
        self.goToViewMenuItem.IsEnabled = self.context_item is not None
        self.goToSheetMenuItem.IsEnabled = (
            self.context_item is not None and self.context_item.sheet is not None)

    def GoToView_Click(self, sender, args):
        if self.context_item is not None:
            self._go_to(self.context_item.view)

    def GoToSheet_Click(self, sender, args):
        if self.context_item is not None and self.context_item.sheet is not None:
            self._go_to(self.context_item.sheet)

    def _go_to(self, view):
        """Close the dialog and have the view opened.

        Revit's UI is blocked while this modal dialog is open, so the view can
        only be looked at once it has closed. The window comes back as it was
        the next time the tool is started, see _save_state.
        """
        self.go_to_element = view
        self.Close()

    def _state_file(self):
        """Per-model file for the window state, in pyRevit's data folder."""
        return script.get_document_data_file("create_references_state", "json")

    def _save_state(self, sender, args):
        """Keep ticks, search, filters and sort for the next start. Runs on every close."""
        state = {
            "search": self.searchTextBox.Text or "",
            "sheet_filter": self.sheetFilterComboBox.SelectedIndex,
            "reference_filter": self.referenceFilterComboBox.SelectedIndex,
            "category_filter": self.categoryFilterComboBox.SelectedItem or CATEGORY_ALL,
            "ticked": [item.view.UniqueId for item in self.all_items if item.IsSelected],
            "sort": [[d.PropertyName, d.Direction == ListSortDirection.Ascending]
                     for d in self._sort_descriptions()],
            "show_depth": self.showDepthCheckbox.IsChecked == True,
        }
        try:
            with open(self._state_file(), "w") as state_file:
                json.dump(state, state_file)
        except Exception as ex:
            logger.debug("Could not save the window state: {}".format(ex))
        self._save_formulas()

    def _read_state(self):
        try:
            with open(self._state_file(), "r") as state_file:
                return json.load(state_file)
        except Exception:
            return None  # first start in this model, or an unreadable file

    def _restore_filters(self, state):
        if not state:
            return
        self.searchTextBox.Text = state.get("search", "")
        self.sheetFilterComboBox.SelectedIndex = state.get("sheet_filter", 0)
        self.referenceFilterComboBox.SelectedIndex = state.get("reference_filter", 0)
        self.showDepthCheckbox.IsChecked = bool(state.get("show_depth", False))
        self._restore_category_filter(state)

    def _restore_category_filter(self, state):
        """Keep a saved category. An older window stored the categories that were unticked."""
        labels = [label for label in self.categoryFilterComboBox.Items]
        saved = state.get("category_filter")
        if saved in labels:
            self.categoryFilterComboBox.SelectedItem = saved
            return
        kinds_off = set(state.get("kinds_off") or [])
        remaining = [label for kind, label in view_references.VIEW_KINDS if kind not in kinds_off]
        if len(remaining) == 1 and remaining[0] in labels:
            self.categoryFilterComboBox.SelectedItem = remaining[0]
            return
        self.categoryFilterComboBox.SelectedIndex = 0

    def _restore_rows(self, state):
        if not state:
            return
        ticked = set(state.get("ticked", []))
        for item in self.all_items:
            item.IsSelected = item.view.UniqueId in ticked
        columns = dict((c.SortMemberPath, c) for c in self.viewsDataGrid.Columns)
        descriptions = self._sort_descriptions()
        for path, ascending in state.get("sort", []):
            if path not in columns:
                continue
            direction = ListSortDirection.Ascending if ascending else ListSortDirection.Descending
            descriptions.Add(SortDescription(path, direction))
            columns[path].SortDirection = direction
        self._update_status()

    def _toggle_selected_rows(self, clicked_item):
        """Give all highlighted rows the clicked row's new state. False if not applicable."""
        highlighted = list(self.viewsDataGrid.SelectedItems)
        if len(highlighted) < 2 or not any(item is clicked_item for item in highlighted):
            return False
        new_state = not clicked_item.IsSelected
        for item in highlighted:
            item.IsSelected = new_state
        return True

    def CreateViewReferences_Click(self, sender, args):
        """Handle create button click."""
        selected_views = [view_data.view for view_data in self.views_data if view_data.IsSelected]
        if not selected_views:
            forms.alert("No views selected. Please select at least one view.", title="Warning")
            return

        # The family may have been loaded since the window opened
        if not self.family_symbol:
            self.family_symbol = view_references.find_family_symbol(doc)
            if not self.family_symbol:
                self._alert_family_missing()
                return

        show_depth = self.showDepthCheckbox.IsChecked == True
        try:
            manual_depth = self._manual_depth()
        except ValueError:
            forms.alert("'{}' is not a depth. Type a number of millimetres, or leave the "
                        "box empty.".format(self.manualDepthTextBox.Text), title="Depth")
            return
        script.get_config().set_option("manual_depth_mm", self.manualDepthTextBox.Text.strip())
        script.save_config()
        sheet_formula = self._formula_dict("sheet")
        view_formula = self._formula_dict("view")
        self._save_formulas()
        with revit.Transaction("Create 3D View References"):
            result = view_references.sync_view_references(
                doc, selected_views, self.family_symbol, show_depth, manual_depth,
                sheet_formula=sheet_formula, view_formula=view_formula)

        self.created_elements = result.element_ids
        self.isolateButton.Content = "Isolate {} references".format(len(self.created_elements))
        self.isolateButton.IsEnabled = len(self.created_elements) > 0
        self._load_views()

        lines = ["Created {} and updated {} 3D View References.".format(
            len(result.created), len(result.updated))]
        if result.skipped:
            lines.append("")
            lines.append("Skipped {}:".format(len(result.skipped)))
            lines.extend("- {}: {}".format(name, reason) for name, reason in result.skipped)
        forms.alert("\n".join(lines), title="3D View References")

    def IsolateElements_Click(self, sender, args):
        """Isolate created elements in current view."""
        if not self.created_elements:
            return

        try:
            element_ids = List[ElementId](self.created_elements)
            with revit.Transaction("Isolate View References"):
                doc.ActiveView.IsolateElementsTemporary(element_ids)
        except Exception as ex:
            logger.error("Error isolating elements: {}".format(ex))

    def Cancel_Click(self, sender, args):
        """Handle cancel button click."""
        self.Close()


# Run the window
if __name__ == '__main__':
    window = Generate3DViewReferencesWindow()
    window.ShowDialog()
    if window.go_to_element is not None:
        __revit__.ActiveUIDocument.ActiveView = window.go_to_element
