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
from Autodesk.Revit.DB import (
    FilteredElementCollector, View, ViewType, FamilySymbol,
    BuiltInCategory, BuiltInParameter, ElementId, Transaction,
    IndependentTag, TagMode, TagOrientation, Reference, XYZ, UV
)
from Autodesk.Revit.DB.Mechanical import Space

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

# Initialize logger
logger = script.get_logger()

# Get current document
doc = revit.doc
uidoc = revit.uidoc


class TagTypeItem(object):
    """Wrapper for space tag family symbol."""
    
    def __init__(self, symbol):
        self.symbol = symbol
        self.element_id = symbol.Id
        self.family_name = symbol.Family.Name if symbol.Family else "Unknown"
        self.type_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if self.type_name:
            self.type_name = self.type_name.AsString() or symbol.Name
        else:
            self.type_name = symbol.Name
        self.display_name = "{}: {}".format(self.family_name, self.type_name)
    
    def __str__(self):
        return self.display_name


class ViewItem(object):
    """Wrapper for view information displayed in DataGrid."""
    
    def __init__(self, view, doc):
        self.view = view
        self.view_id = view.Id
        self.view_name = view.Name
        
        # View type display
        if view.ViewType == ViewType.FloorPlan:
            self.view_type = "Floor Plan"
        elif view.ViewType == ViewType.CeilingPlan:
            self.view_type = "Ceiling Plan"
        else:
            self.view_type = str(view.ViewType)
        
        # View template
        template_id = view.ViewTemplateId
        if template_id and template_id != ElementId.InvalidElementId:
            template = doc.GetElement(template_id)
            self.view_template = template.Name if template else "None"
        else:
            self.view_template = "None"
        
        # View phase
        phase_param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
        if phase_param:
            phase_id = phase_param.AsElementId()
            if phase_id and phase_id != ElementId.InvalidElementId:
                phase = doc.GetElement(phase_id)
                self.view_phase = phase.Name if phase else "?"
            else:
                self.view_phase = "?"
        else:
            self.view_phase = "?"
        
        # Associated level
        if hasattr(view, 'GenLevel') and view.GenLevel:
            self.level_name = view.GenLevel.Name
        else:
            self.level_name = "?"
    
    def matches_search(self, search_text):
        """Check if view matches search text."""
        if not search_text:
            return True
        search_lower = search_text.lower()
        return (search_lower in self.view_name.lower() or
                search_lower in self.view_type.lower() or
                search_lower in self.view_template.lower() or
                search_lower in self.view_phase.lower() or
                search_lower in self.level_name.lower())


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


def get_space_tag_types(doc):
    """Get all space tag family symbols.
    
    Args:
        doc: Revit document
        
    Returns:
        List of TagTypeItem objects
    """
    tag_types = []
    
    try:
        # Get all space tag family symbols
        collector = FilteredElementCollector(doc)\
            .OfCategory(BuiltInCategory.OST_MEPSpaceTags)\
            .OfClass(FamilySymbol)\
            .ToElements()
        
        for symbol in collector:
            tag_types.append(TagTypeItem(symbol))
        
        # Sort by display name
        tag_types.sort(key=lambda x: x.display_name)
        
    except Exception as e:
        logger.error("Error getting space tag types: {}".format(str(e)))
    
    return tag_types


def get_plan_views(doc):
    """Get all floor plans and ceiling plans (excluding templates and dependent views).
    
    Args:
        doc: Revit document
        
    Returns:
        List of ViewItem objects
    """
    view_items = []
    
    try:
        # Get all views
        collector = FilteredElementCollector(doc)\
            .OfClass(View)\
            .ToElements()
        
        for view in collector:
            # Skip templates
            if view.IsTemplate:
                continue
            
            # Only include floor plans and ceiling plans
            if view.ViewType not in [ViewType.FloorPlan, ViewType.CeilingPlan]:
                continue
            
            # Skip dependent views (only show parent views)
            # Dependent views have a valid primary view id
            try:
                primary_view_id = view.GetPrimaryViewId()
                if primary_view_id and primary_view_id != ElementId.InvalidElementId:
                    continue  # This is a dependent view, skip it
            except:
                pass  # GetPrimaryViewId might not exist in older versions
            
            view_items.append(ViewItem(view, doc))
        
        # Sort by view name
        view_items.sort(key=lambda x: x.view_name)
        
    except Exception as e:
        logger.error("Error getting views: {}".format(str(e)))
    
    return view_items


def tag_spaces_in_view(doc, view, tag_type_id):
    """Tag all untagged spaces in a view.
    
    Uses doc.Create.NewSpaceTag() which is faster than IndependentTag.Create().
    
    Args:
        doc: Revit document
        view: View to tag spaces in
        tag_type_id: ElementId of the tag type to use
        
    Returns:
        Tuple of (tagged_count, skipped_count, error_count)
    """
    tagged_count = 0
    skipped_count = 0
    error_count = 0
    
    try:
        # Get all spaces visible in this view
        spaces = list(FilteredElementCollector(doc, view.Id)\
            .OfCategory(BuiltInCategory.OST_MEPSpaces)\
            .WhereElementIsNotElementType()\
            .ToElements())
        
        if not spaces:
            return (0, 0, 0)
        
        # Get existing space tags to find already tagged spaces
        existing_tags = FilteredElementCollector(doc, view.Id)\
            .OfCategory(BuiltInCategory.OST_MEPSpaceTags)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        # Build set of already tagged space IDs
        tagged_space_ids = set()
        for tag in existing_tags:
            try:
                # Try different methods to get tagged element ID
                if hasattr(tag, 'TaggedLocalElementId'):
                    tagged_id = tag.TaggedLocalElementId
                    if tagged_id and tagged_id != ElementId.InvalidElementId:
                        tagged_space_ids.add(tagged_id)
                elif hasattr(tag, 'GetTaggedLocalElementIds'):
                    for tagged_id in tag.GetTaggedLocalElementIds():
                        if tagged_id and tagged_id != ElementId.InvalidElementId:
                            tagged_space_ids.add(tagged_id)
            except Exception:
                pass
        
        # Pre-filter spaces to tag and get their locations
        spaces_to_tag = []
        for space in spaces:
            if space.Id in tagged_space_ids:
                skipped_count += 1
                continue
            if space.Area <= 0:
                skipped_count += 1
                continue
            
            # Get location point - convert to UV for NewSpaceTag
            location = space.Location
            if location and hasattr(location, 'Point'):
                pt = location.Point
                spaces_to_tag.append((space, UV(pt.X, pt.Y)))
            else:
                # Only compute bounding box if no location point
                bb = space.get_BoundingBox(view)
                if bb:
                    uv_point = UV(
                        (bb.Min.X + bb.Max.X) / 2,
                        (bb.Min.Y + bb.Max.Y) / 2
                    )
                    spaces_to_tag.append((space, uv_point))
                else:
                    error_count += 1
        
        # Tag all spaces using NewSpaceTag (faster than IndependentTag.Create)
        for space, uv_point in spaces_to_tag:
            try:
                # NewSpaceTag takes (space, UV point, view) - simpler and faster
                space_tag = doc.Create.NewSpaceTag(space, uv_point, view)
                
                # Change to selected tag type
                if space_tag and tag_type_id:
                    space_tag.ChangeTypeId(tag_type_id)
                
                tagged_count += 1
            except Exception as e:
                logger.debug("Error tagging space {}: {}".format(space.Id, str(e)))
                error_count += 1
    
    except Exception as e:
        logger.error("Error processing view {}: {}".format(view.Name, str(e)))
    
    return (tagged_count, skipped_count, error_count)


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
