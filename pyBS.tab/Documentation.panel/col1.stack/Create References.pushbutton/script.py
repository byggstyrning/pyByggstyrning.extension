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


class ViewItemData(forms.Reactive):
    """Class for view data binding with WPF UI."""

    def __init__(self, view, sheet_reference, has_reference):
        """Initialize with a Revit view."""
        super(ViewItemData, self).__init__()
        self.view = view
        self._is_selected = True
        self.view_name = view.Name
        self.view_category = view_references.get_view_kind_label(view)
        self.view_scale = "1:{}".format(view.Scale) if view.Scale else "Unknown"
        self.sheet_reference = sheet_reference
        self.reference_status = "Placed" if has_reference else ""

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

        # Initialize view data
        self.views_data = ObservableCollection[ViewItemData]()

        self.family_symbol = view_references.find_family_symbol(doc)
        if not self.family_symbol:
            self._alert_family_missing()

        self._setup_view_categories()
        self._populate_views()

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

    def _populate_views(self):
        """Populate the views data grid based on selected categories."""
        deselected = set(item.view.UniqueId for item in self.views_data if not item.IsSelected)
        self.views_data.Clear()

        checked_kinds = [kind for kind, cb in self.view_kinds.items() if cb.IsChecked == True]
        sheet_lookup = view_references.build_sheet_lookup(doc)
        existing = view_references.find_existing_references(doc)

        views = view_references.collect_views(doc, checked_kinds)
        for view in sorted(views, key=lambda v: v.Name):
            sheet = sheet_lookup.get(get_element_id_value(view.Id))
            sheet_reference = view_references.get_sheet_label(sheet) if sheet else "Not on sheet"
            item = ViewItemData(view, sheet_reference, view.UniqueId in existing)
            item.IsSelected = view.UniqueId not in deselected
            self.views_data.Add(item)

        logger.debug("Added {} views to data grid".format(self.views_data.Count))

    def ViewCategory_CheckedChanged(self, sender, args):
        """Handle view category checkbox changes."""
        self._populate_views()

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
        with revit.Transaction("Create 3D View References"):
            result = view_references.sync_view_references(
                doc, selected_views, self.family_symbol, show_depth)

        self.created_elements = result.element_ids
        self.isolateButton.Content = "Isolate {} references".format(len(self.created_elements))
        self.isolateButton.IsEnabled = len(self.created_elements) > 0
        self._populate_views()

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
