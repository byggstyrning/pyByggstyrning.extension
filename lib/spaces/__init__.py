# -*- coding: utf-8 -*-
"""Shared space tagging functionality for pyByggstyrning.

This module provides common classes and functions for working with
MEP Spaces and space tags in Revit.

Exports:
    - TagTypeItem: Wrapper class for space tag family symbols
    - ViewItem: Wrapper class for view information
    - get_space_tag_types(doc): Get all space tag family symbols
    - get_plan_views(doc): Get all floor/ceiling plans (non-template, non-dependent)
    - get_views_with_space_tags(doc): Find views containing space tags
    - tag_spaces_in_view(doc, view, tag_type_id, space_ids=None): Tag spaces in a view
"""

from __future__ import print_function

# .NET imports
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FilteredElementCollector, View, ViewType, FamilySymbol,
    BuiltInCategory, BuiltInParameter, ElementId, UV
)
from Autodesk.Revit.DB.Mechanical import Space

# pyRevit imports
from pyrevit import script

# Initialize logger
logger = script.get_logger()


class TagTypeItem(object):
    """Wrapper for space tag family symbol."""
    
    def __init__(self, symbol):
        """Initialize tag type item.
        
        Args:
            symbol: FamilySymbol for the space tag
        """
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
    
    def __repr__(self):
        return self.display_name
    
    def ToString(self):
        """Explicit ToString for WPF binding in IronPython."""
        return self.display_name


class ViewItem(object):
    """Wrapper for view information displayed in DataGrid."""
    
    def __init__(self, view, doc):
        """Initialize view item.
        
        Args:
            view: Revit View element
            doc: Revit document
        """
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
        """Check if view matches search text.
        
        Args:
            search_text: Text to search for (case-insensitive)
            
        Returns:
            bool: True if view matches search text
        """
        if not search_text:
            return True
        search_lower = search_text.lower()
        return (search_lower in self.view_name.lower() or
                search_lower in self.view_type.lower() or
                search_lower in self.view_template.lower() or
                search_lower in self.view_phase.lower() or
                search_lower in self.level_name.lower())


def get_space_tag_types(doc):
    """Get all space tag family symbols.
    
    Args:
        doc: Revit document
        
    Returns:
        list: List of TagTypeItem objects sorted by display name
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
        list: List of ViewItem objects sorted by view name
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


def get_views_with_space_tags(doc):
    """Find all floor/ceiling plan views that contain space tags.
    
    OPTIMIZED: Uses a single document-wide FilteredElementCollector to find all
    space tags, then extracts their OwnerViewId. This is O(total_tags) instead
    of O(views × tags_per_view).
    
    Args:
        doc: Revit document
        
    Returns:
        dict: {view_id: True} for views containing space tags.
              We only need to know which views have tags, not which spaces are tagged.
    """
    views_with_tags = {}
    
    try:
        # OPTIMIZED: Single document-wide collection of ALL space tags
        # This is much faster than iterating through each view
        all_space_tags = FilteredElementCollector(doc)\
            .OfCategory(BuiltInCategory.OST_MEPSpaceTags)\
            .WhereElementIsNotElementType()\
            .ToElements()
        
        tags_list = list(all_space_tags) if all_space_tags else []
        
        # Get the set of valid plan view IDs (floor plans and ceiling plans only)
        plan_views = get_plan_views(doc)
        valid_view_ids = set(view_item.view.Id for view_item in plan_views)
        
        # Extract OwnerViewId from each tag
        for tag in tags_list:
            try:
                # SpaceTag inherits from SpatialElementTag which has OwnerViewId
                owner_view_id = tag.OwnerViewId
                if owner_view_id and owner_view_id != ElementId.InvalidElementId:
                    # Only include if it's a floor plan or ceiling plan
                    if owner_view_id in valid_view_ids:
                        views_with_tags[owner_view_id] = True
            except Exception:
                # Fallback: try View property if OwnerViewId doesn't work
                try:
                    if hasattr(tag, 'View') and tag.View:
                        view_id = tag.View.Id
                        if view_id in valid_view_ids:
                            views_with_tags[view_id] = True
                except Exception:
                    pass
        
    except Exception as e:
        logger.error("Error in get_views_with_space_tags: {}".format(str(e)))
    
    return views_with_tags


def tag_spaces_in_view(doc, view, tag_type_id, space_ids=None):
    """Tag all untagged spaces in a view.
    
    Uses doc.Create.NewSpaceTag() which is faster than IndependentTag.Create().
    
    Args:
        doc: Revit document
        view: View to tag spaces in
        tag_type_id: ElementId of the tag type to use
        space_ids: Optional set of space IDs to limit tagging to (e.g., newly created spaces).
                   If None, all untagged spaces in the view will be tagged.
        
    Returns:
        tuple: (tagged_count, skipped_count, error_count)
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
                # SpaceTag has a direct Space property that returns the tagged Space
                if hasattr(tag, 'Space') and tag.Space:
                    tagged_space_ids.add(tag.Space.Id)
                # Fallback for other tag types
                elif hasattr(tag, 'TaggedLocalElementId'):
                    tagged_id = tag.TaggedLocalElementId
                    if tagged_id and tagged_id != ElementId.InvalidElementId:
                        tagged_space_ids.add(tagged_id)
            except Exception:
                pass
        
        # Pre-filter spaces to tag and get their locations
        spaces_to_tag = []
        for space in spaces:
            # If space_ids filter is provided, only tag those specific spaces
            if space_ids is not None and space.Id not in space_ids:
                continue
            
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
