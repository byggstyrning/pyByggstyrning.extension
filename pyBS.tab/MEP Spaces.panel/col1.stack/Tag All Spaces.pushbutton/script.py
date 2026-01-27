# -*- coding: utf-8 -*-
"""Tag all MEP Spaces in selected views.

Select views and a space tag type to automatically tag all untagged
spaces in the selected floor plans and ceiling plans.
"""

__title__ = "Tag All\nSpaces"
__author__ = "Byggstyrning AB"
__doc__ = """Tag all MEP Spaces in selected views.

Select a space tag type and one or more floor plans or ceiling plans,
then click 'Tag Spaces' to automatically place tags on all untagged spaces.

Features:
- Filter and sort views by type, name, template, phase, and level
- Multi-select views using Shift+Click or Ctrl+Click
- Search to quickly find views
"""
__highlight__ = 'new'

# Standard library imports
import sys
import os.path as op

# .NET imports
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import ElementId

# pyRevit imports
from pyrevit import script, forms, revit

# Path setup for lib imports
script_dir = op.dirname(__file__)
pushbutton_dir = script_dir
stack_dir = op.dirname(pushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Import shared space tagging functionality from lib
from spaces import (
    TagTypeItem, ViewItem,
    get_space_tag_types, get_plan_views,
    tag_spaces_in_view
)

# Initialize logger
logger = script.get_logger()

# Get current document
doc = revit.doc
uidoc = revit.uidoc


class TagAllSpacesWindow(forms.WPFWindow):
    """WPF Window for selecting views and tag type."""
    
    def __init__(self, tag_types, view_items, pushbutton_dir):
        """Initialize the window.
        
        Args:
            tag_types: List of TagTypeItem objects
            view_items: List of ViewItem objects
            pushbutton_dir: Path to pushbutton directory for XAML
        """
        # Load styles before window initialization
        from styles import ensure_styles_loaded
        ensure_styles_loaded()
        
        # Load XAML
        xaml_path = op.join(pushbutton_dir, "TagAllSpacesWindow.xaml")
        forms.WPFWindow.__init__(self, xaml_path)
        
        # Store data
        self.tag_types = tag_types
        self.all_view_items = view_items
        self.filtered_view_items = list(view_items)
        self.selected_tag_type = None
        self.selected_views = []
        
        # Populate tag type combo box
        self.tagTypeComboBox.ItemsSource = self.tag_types
        if self.tag_types:
            self.tagTypeComboBox.SelectedIndex = 0
        
        # Populate views DataGrid
        self.viewsDataGrid.ItemsSource = self.filtered_view_items
        
        # Update status
        self._update_status()
    
    def _update_status(self):
        """Update status text with selection count."""
        selected_count = len(list(self.viewsDataGrid.SelectedItems))
        total_count = len(self.filtered_view_items)
        
        if selected_count == 0:
            self.statusText.Text = "{} views available - Select views to tag spaces".format(total_count)
        else:
            self.statusText.Text = "{} of {} views selected".format(selected_count, total_count)
    
    def searchTextBox_TextChanged(self, sender, args):
        """Handle search text changed event."""
        search_text = self.searchTextBox.Text or ""
        
        # Filter views
        self.filtered_view_items = [
            item for item in self.all_view_items
            if item.matches_search(search_text)
        ]
        
        # Update DataGrid
        self.viewsDataGrid.ItemsSource = self.filtered_view_items
        self._update_status()
    
    def tagButton_Click(self, sender, args):
        """Handle Tag Spaces button click."""
        # Validate tag type selection
        if self.tagTypeComboBox.SelectedItem is None:
            forms.alert("Please select a space tag type.", title="No Tag Type")
            return
        
        # Validate view selection
        selected_items = list(self.viewsDataGrid.SelectedItems)
        if not selected_items:
            forms.alert("Please select at least one view.", title="No Views Selected")
            return
        
        # Store selections
        self.selected_tag_type = self.tagTypeComboBox.SelectedItem
        self.selected_views = [item.view for item in selected_items]
        
        # Close with success
        self.DialogResult = True
        self.Close()
    
    def cancelButton_Click(self, sender, args):
        """Handle Cancel button click."""
        self.DialogResult = False
        self.Close()


def main():
    """Main entry point."""
    # Get space tag types
    tag_types = get_space_tag_types(doc)
    
    if not tag_types:
        forms.alert(
            "No space tag types are loaded in this project.\n\n"
            "Please load a space tag family before using this tool.",
            title="No Space Tags Available"
        )
        return
    
    # Get plan views
    view_items = get_plan_views(doc)
    
    if not view_items:
        forms.alert(
            "No floor plans or ceiling plans found in this project.",
            title="No Views Available"
        )
        return
    
    # Show dialog
    dialog = TagAllSpacesWindow(tag_types, view_items, pushbutton_dir)
    result = dialog.ShowDialog()
    
    if not result:
        return
    
    # Get selections
    selected_tag_type = dialog.selected_tag_type
    selected_views = dialog.selected_views
    
    if not selected_tag_type or not selected_views:
        return
    
    # Ensure tag type is activated
    tag_type_id = selected_tag_type.element_id
    
    # Tag spaces in selected views
    total_tagged = 0
    total_skipped = 0
    total_errors = 0
    views_processed = 0
    
    # Use progress bar for multiple views
    with forms.ProgressBar(title="Tagging Spaces...", cancellable=True) as pb:
        max_value = len(selected_views)
        
        with revit.Transaction("Tag All Spaces"):
            # Activate tag type if not already
            tag_symbol = doc.GetElement(tag_type_id)
            if tag_symbol and not tag_symbol.IsActive:
                tag_symbol.Activate()
                doc.Regenerate()
            
            for i, view in enumerate(selected_views):
                if pb.cancelled:
                    break
                
                pb.update_progress(i, max_value)
                
                tagged, skipped, errors = tag_spaces_in_view(doc, view, tag_type_id)
                
                total_tagged += tagged
                total_skipped += skipped
                total_errors += errors
                views_processed += 1
    
    # Show results
    if pb.cancelled:
        message = "Operation cancelled.\n\n"
    else:
        message = ""
    
    message += "Tagged {} spaces in {} views.\n".format(total_tagged, views_processed)
    
    if total_skipped > 0:
        message += "\n{} spaces were already tagged or invalid.".format(total_skipped)
    
    if total_errors > 0:
        message += "\n{} spaces could not be tagged (check log for details).".format(total_errors)
    
    if total_tagged > 0:
        forms.alert(message, title="Tagging Complete")
    else:
        forms.alert(message, title="No Spaces Tagged")


if __name__ == '__main__':
    main()
