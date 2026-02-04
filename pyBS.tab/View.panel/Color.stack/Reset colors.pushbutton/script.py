# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, OverrideGraphicSettings, Transaction, ViewType, ViewSheet

__title__ = "Reset"
__author__ = ""
__doc__ = """Reset element color overrides in the active view.
If on a sheet, resets colors in all views placed on that sheet."""


def can_color_view(view):
    """Check if a view can have element colors applied."""
    try:
        # Check if view supports temporary visibility modes
        if not view.CanUseTemporaryVisibilityModes():
            return False
        
        # Exclude certain view types
        excluded_types = [
            ViewType.ProjectBrowser,
            ViewType.SystemBrowser,
            ViewType.Schedule,
            ViewType.Legend,
            ViewType.Report,
            ViewType.DraftingView,
            ViewType.DrawingSheet,  # The sheet itself, not the views on it
        ]
        
        if view.ViewType in excluded_types:
            return False
        
        return True
        
    except Exception:
        return False


def get_views_from_sheet(doc, sheet):
    """Get all colorable views placed on a sheet."""
    try:
        placed_view_ids = sheet.GetAllPlacedViews()
        
        views = []
        for view_id in placed_view_ids:
            view = doc.GetElement(view_id)
            if view and can_color_view(view):
                views.append(view)
        
        return views
        
    except Exception:
        return []


def find_sheet_containing_view(doc, view):
    """Find the sheet that contains the given view."""
    try:
        # Get all sheets in the document
        sheets = FilteredElementCollector(doc) \
                    .OfClass(ViewSheet) \
                    .ToElements()
        
        for sheet in sheets:
            placed_view_ids = sheet.GetAllPlacedViews()
            if view.Id in placed_view_ids:
                return sheet
        
        return None
        
    except Exception:
        return None


def get_target_views(doc, active_view):
    """Auto-detect target views based on context.
    
    If we're on a sheet, target all views on that sheet.
    If active view is placed on a sheet, target all views on that sheet.
    Otherwise, target just the active view.
    """
    try:
        # Check if active view is a sheet
        if active_view.ViewType == ViewType.DrawingSheet:
            # We're on a sheet - get all views on it
            views = get_views_from_sheet(doc, active_view)
            if views:
                return views
        else:
            # Check if active view is placed on a sheet
            sheet = find_sheet_containing_view(doc, active_view)
            if sheet:
                views = get_views_from_sheet(doc, sheet)
                if views:
                    return views
        
        # Regular view or fallback - just use the active view if it's colorable
        if can_color_view(active_view):
            return [active_view]
        
        return []
        
    except Exception:
        # Fallback to active view
        if can_color_view(active_view):
            return [active_view]
        return []


def reset_colors_in_view(doc, view):
    """Reset color overrides for all elements in a view."""
    ogs = OverrideGraphicSettings()
    collector = FilteredElementCollector(doc, view.Id).ToElements()
    reset_count = 0
    
    for element in collector:
        try:
            view.SetElementOverrides(element.Id, ogs)
            reset_count += 1
        except Exception:
            # Some elements may not support overrides
            pass
    
    return reset_count


if __name__ == '__main__':
    doc = __revit__.ActiveUIDocument.Document
    active_view = doc.ActiveView
    
    # Get target views (handles sheets automatically)
    target_views = get_target_views(doc, active_view)
    
    if not target_views:
        print("No views found that support color overrides.")
    else:
        with Transaction(doc, 'Reset Colors') as t:
            t.Start()
            
            total_reset = 0
            for view in target_views:
                count = reset_colors_in_view(doc, view)
                total_reset += count
                print("Reset {} element(s) in view: {}".format(count, view.Name))
            
            t.Commit()
            
            if len(target_views) > 1:
                print("\nTotal: Reset colors in {} views.".format(len(target_views)))

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.