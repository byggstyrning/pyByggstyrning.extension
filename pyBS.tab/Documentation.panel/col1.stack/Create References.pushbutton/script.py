# -*- coding: utf-8 -*-
__title__ = "Create References"
__author__ = "Jonatan Jacobsson"
__doc__ = """Places a 3D View Reference family at the location and extent of the selected views.
Views that already have a reference get it updated instead of getting a second one.
"""

# Import .NET libraries
import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Thickness, Visibility
from System.Windows.Controls import CheckBox
from System.Windows.Media import VisualTreeHelper

# Import Revit API
from Autodesk.Revit.DB import ElementId

# Import pyRevit libraries
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


class ViewItemData(forms.Reactive):
    """Class for view data binding with WPF UI."""

    def __init__(self, view, sheet, has_reference):
        """Initialize with a Revit view and the sheet it is placed on, if any."""
        super(ViewItemData, self).__init__()
        self.view = view
        self.sheet = sheet
        self.kind = view_references.get_view_kind(view)
        self.has_reference = has_reference
        self._is_selected = True
        self._sheet_parameter_value = ""
        self.view_name = view.Name
        self.view_category = view_references.get_view_kind_label(view)
        self.view_scale = "1:{}".format(view.Scale) if view.Scale else "Unknown"
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

        self._setup_view_categories()
        self._setup_filters()
        self._load_views()

        # Bind views to DataGrid
        self.viewsDataGrid.ItemsSource = self.views_data
        self.selectAllCheckbox.IsChecked = True
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

    def _setup_view_categories(self):
        """Set up view category checkboxes."""
        self.view_kinds = {}
        for kind, label in view_references.VIEW_KINDS:
            checkbox = CheckBox()
            checkbox.Content = label
            checkbox.IsChecked = True
            checkbox.Margin = Thickness(0, 0, 15, 0)
            checkbox.SetResourceReference(CheckBox.StyleProperty, "StandardCheckBoxStyle")
            checkbox.SetResourceReference(CheckBox.ForegroundProperty, "TextBrush")
            checkbox.Checked += self.ViewCategory_CheckedChanged
            checkbox.Unchecked += self.ViewCategory_CheckedChanged

            self.viewCategoriesPanel.Children.Add(checkbox)
            self.view_kinds[kind] = checkbox

    def _setup_filters(self):
        """Fill the filter dropdowns. The sheet parameter choice is remembered."""
        self.sheetFilterComboBox.ItemsSource = List[str](SHEET_FILTERS)
        self.sheetFilterComboBox.SelectedIndex = 0
        self.referenceFilterComboBox.ItemsSource = List[str](REFERENCE_FILTERS)
        self.referenceFilterComboBox.SelectedIndex = 0

        self.manualDepthTextBox.Text = script.get_config().get_option("manual_depth_mm", "")

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

    def _sheet_parameter_name(self):
        name = self.sheetParameterComboBox.SelectedItem
        return None if not name or name == NO_SHEET_PARAMETER else name

    def _load_views(self):
        """Read all views from the model, keeping which rows were unticked."""
        deselected = set(item.view.UniqueId for item in (self.all_items or [])
                         if not item.IsSelected)
        sheet_lookup = view_references.build_sheet_lookup(doc)
        existing = view_references.find_existing_references(doc)

        self.all_items = []
        for view in sorted(view_references.collect_views(doc), key=lambda v: v.Name):
            sheet = sheet_lookup.get(get_element_id_value(view.Id))
            item = ViewItemData(view, sheet, view.UniqueId in existing)
            item.IsSelected = view.UniqueId not in deselected
            self.all_items.append(item)
        self._update_sheet_parameter_column()
        self._apply_filters()

    def _update_sheet_parameter_column(self):
        name = self._sheet_parameter_name()
        self.sheetParameterColumn.Header = name or "Sheet Parameter"
        self.sheetParameterColumn.Visibility = (
            Visibility.Visible if name else Visibility.Collapsed)
        for item in self.all_items:
            item.SheetParameterValue = view_references.get_parameter_text(item.sheet, name)

    def _apply_filters(self):
        """Show the rows that pass the category, search, sheet and reference filters."""
        kinds = set(kind for kind, cb in self.view_kinds.items() if cb.IsChecked == True)
        words = (self.searchTextBox.Text or "").lower().split()
        sheet_filter = self.sheetFilterComboBox.SelectedItem
        reference_filter = self.referenceFilterComboBox.SelectedItem

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
            self.views_data.Add(item)
        self.countTextBlock.Text = "Showing {} of {} views".format(
            self.views_data.Count, len(self.all_items))

    def ViewCategory_CheckedChanged(self, sender, args):
        """Handle view category checkbox changes."""
        if self.all_items is not None:
            self._apply_filters()

    def Filter_Changed(self, sender, args):
        """Search text or a filter dropdown changed."""
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
        while element is not None:
            if isinstance(element, CheckBox):
                return element
            try:
                element = VisualTreeHelper.GetParent(element)
            except Exception:
                return None  # not a visual, e.g. a text run
        return None

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
        with revit.Transaction("Create 3D View References"):
            result = view_references.sync_view_references(
                doc, selected_views, self.family_symbol, show_depth, manual_depth)

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
