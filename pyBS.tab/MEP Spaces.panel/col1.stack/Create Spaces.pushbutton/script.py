# -*- coding: utf-8 -*-
"""Create MEP Spaces from Rooms in linked Revit models.

This tool creates MEP Spaces at the same locations as Rooms in linked models.
It handles level matching, phase handling, and coordinate transformation.
"""

__title__ = "Create Spaces\nfrom Link"
__author__ = "Byggstyrning AB"
__doc__ = """Create MEP Spaces from Rooms in linked Revit models.

Select a linked model containing Rooms, and this tool will create
MEP Spaces at the same locations in the current model.

Options:
- Write Name and Number from Room to Space
- Remove and recreate spaces (preserves other parameters when linked to a Room; keeps unlinked spaces)
"""


# Standard library imports
import sys
import os.path as op
import time

# .NET imports
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Windows import Visibility

# Revit API imports
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    SpatialElement,
    Level,
    Phase,
    UV,
    Transaction,
    ElementId,
    BuiltInParameter,
    View
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.DB.Mechanical import Space

# pyRevit imports
from pyrevit import script
from pyrevit import forms
from pyrevit import revit
from pyrevit.forms import WPFWindow
from pyrevit.revit.db import failure as revit_failure

# Set up paths for lib imports
script_dir = op.dirname(__file__)
pushbutton_dir = script_dir
stack_dir = op.dirname(pushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from revit.compat import get_element_id_value

# Initialize logger
logger = script.get_logger()

# Document references
doc = revit.doc
uidoc = revit.uidoc


class LinkedModelItem(object):
    """Represents a linked model with rooms for display in ComboBox."""
    
    def __init__(self, link_instance, room_count, phase_room_counts):
        """Initialize linked model item.
        
        Args:
            link_instance: RevitLinkInstance element
            room_count: Total number of placed rooms in the linked model
            phase_room_counts: Dict mapping phase names to room counts
        """
        self.link_instance = link_instance
        self.link_doc = link_instance.GetLinkDocument()
        self.room_count = room_count
        self.phase_room_counts = phase_room_counts
        
        # Extract display name from link name
        link_name = link_instance.Name
        # Remove file extension and instance info if present
        if ':' in link_name:
            link_name = link_name.split(':')[0].strip()
        self.display_name = "{} ({} rooms)".format(link_name, room_count)
    
    def __str__(self):
        return self.display_name
    
    def __repr__(self):
        return self.display_name
    
    def ToString(self):
        """Explicit ToString for WPF binding in IronPython."""
        return self.display_name


def _room_label(room):
    """Short name/number label for debug logs."""
    try:
        number_param = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        name_param = room.get_Parameter(BuiltInParameter.ROOM_NAME)
        number = number_param.AsString() if number_param else None
        name = name_param.AsString() if name_param else None
        if number and name:
            return "{} {}".format(number, name)
        return number or name or "?"
    except Exception:
        return "?"


def _collect_linked_rooms(link_doc):
    """Collect rooms from a linked document and classify placed vs skipped.

    Returns:
        tuple: (all_spatial, room_elements, placed_rooms, stats)
        stats keys: spatial, rooms, placed, zero_area, no_location
    """
    all_spatial = FilteredElementCollector(link_doc).OfClass(SpatialElement).ToElements()
    room_elements = [r for r in all_spatial if isinstance(r, Room)]
    placed_rooms = []
    zero_area = 0
    no_location = 0
    for room in room_elements:
        if not room.Location:
            no_location += 1
        if room.Area > 0:
            placed_rooms.append(room)
        else:
            zero_area += 1
    stats = {
        'spatial': len(all_spatial),
        'rooms': len(room_elements),
        'placed': len(placed_rooms),
        'zero_area': zero_area,
        'no_location': no_location,
    }
    return all_spatial, room_elements, placed_rooms, stats


def _log_found_rooms(link_name, link_doc, placed_rooms, stats, sample_limit=8):
    """Log room inventory for a linked model (counts + a small sample)."""
    logger.debug(
        "Found rooms in '{}': spatial={}, rooms={}, placed(area>0)={}, "
        "zero_area={}, no_location={}".format(
            link_name,
            stats['spatial'],
            stats['rooms'],
            stats['placed'],
            stats['zero_area'],
            stats['no_location']))

    phase_counts = {}
    level_counts = {}
    for room in placed_rooms:
        try:
            phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
            phase_id = phase_param.AsElementId() if phase_param else None
            phase = link_doc.GetElement(phase_id) if phase_id else None
            phase_name = phase.Name if phase else "(no phase)"
            phase_counts[phase_name] = phase_counts.get(phase_name, 0) + 1
        except Exception:
            phase_counts['(phase error)'] = phase_counts.get('(phase error)', 0) + 1
        try:
            level_name = room.Level.Name if room.Level else "(no level)"
            level_counts[level_name] = level_counts.get(level_name, 0) + 1
        except Exception:
            level_counts['(level error)'] = level_counts.get('(level error)', 0) + 1

    if phase_counts:
        logger.debug("  placed rooms by phase: {}".format(phase_counts))
    if level_counts:
        logger.debug("  placed rooms by level: {}".format(level_counts))

    for room in placed_rooms[:sample_limit]:
        try:
            level_name = room.Level.Name if room.Level else None
            level_elev = room.Level.Elevation if room.Level else None
            phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
            phase_id = phase_param.AsElementId() if phase_param else None
            phase = link_doc.GetElement(phase_id) if phase_id else None
            loc = room.Location.Point if room.Location else None
            logger.debug(
                "  sample room '{}' area={} level={} elev={} phase={} loc=({}, {}, {})".format(
                    _room_label(room),
                    room.Area,
                    level_name,
                    level_elev,
                    phase.Name if phase else None,
                    loc.X if loc else None,
                    loc.Y if loc else None,
                    loc.Z if loc else None))
        except Exception as e:
            logger.debug("  sample room log failed: {}".format(str(e)))


def get_linked_documents_with_rooms(doc):
    """Get all linked Revit documents that contain placed rooms.
    
    Args:
        doc: Host Revit document
        
    Returns:
        list: List of LinkedModelItem objects for links with rooms
    """
    linked_models = []
    
    try:
        link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
        logger.debug("Scanning {} Revit link instance(s) for rooms".format(len(link_instances)))
        
        for link in link_instances:
            try:
                link_doc = link.GetLinkDocument()
                if not link_doc:
                    logger.debug("Skipping unloaded link: {}".format(link.Name))
                    continue
                
                _, _, placed_rooms, stats = _collect_linked_rooms(link_doc)
                _log_found_rooms(link.Name, link_doc, placed_rooms, stats)
                
                if len(placed_rooms) > 0:
                    # Count rooms per phase
                    phase_counts = {}
                    for room in placed_rooms:
                        try:
                            phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
                            if phase_param:
                                phase_id = phase_param.AsElementId()
                                phase = link_doc.GetElement(phase_id)
                                if phase:
                                    phase_name = phase.Name
                                    phase_counts[phase_name] = phase_counts.get(phase_name, 0) + 1
                        except Exception:
                            pass
                    
                    linked_models.append(LinkedModelItem(link, len(placed_rooms), phase_counts))
                    
            except Exception as e:
                logger.debug("Error accessing linked document: {}".format(str(e)))
                continue
        
        # Sort by name
        linked_models.sort(key=lambda x: x.display_name)
        logger.debug("Linked models with placed rooms: {}".format(
            [item.display_name for item in linked_models]))
        
    except Exception as e:
        logger.error("Error getting linked documents: {}".format(str(e)))
    
    return linked_models


# Elevation tolerance for level matching (1 foot = ~300mm)
LEVEL_ELEVATION_TOLERANCE = 1.0


def get_host_levels_by_elevation(doc):
    """Get list of host model levels sorted by elevation.
    
    Args:
        doc: Host Revit document
        
    Returns:
        list: List of tuples (elevation, Level) sorted by elevation
    """
    levels = []
    try:
        all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
        for level in all_levels:
            levels.append((level.Elevation, level))
        # Sort by elevation for efficient lookup
        levels.sort(key=lambda x: x[0])
    except Exception as e:
        logger.error("Error getting levels: {}".format(str(e)))
    return levels


def find_level_by_elevation(host_levels, target_elevation, tolerance=LEVEL_ELEVATION_TOLERANCE):
    """Find host level matching the target elevation within tolerance.
    
    This mimics how Revit's "Place Spaces Automatically" matches levels
    by elevation rather than by name.
    
    Args:
        host_levels: List of (elevation, Level) tuples, sorted by elevation
        target_elevation: The elevation to match (in feet)
        tolerance: Maximum elevation difference allowed (default 1 foot)
    
    Returns:
        Level or None: The matching host level, or None if no match within tolerance
    """
    best_match = None
    best_diff = float('inf')
    
    for elev, level in host_levels:
        diff = abs(elev - target_elevation)
        if diff < best_diff and diff <= tolerance:
            best_diff = diff
            best_match = level
    
    return best_match


def get_host_phases(doc):
    """Get all phases in the host document.
    
    Args:
        doc: Host Revit document
        
    Returns:
        list: List of Phase elements
    """
    phases = []
    try:
        all_phases = FilteredElementCollector(doc).OfClass(Phase).ToElements()
        phases = list(all_phases)
    except Exception as e:
        logger.error("Error getting phases: {}".format(str(e)))
    return phases


def delete_existing_spaces(doc, exclude_ids=None):
    """Delete existing spaces in the document, optionally preserving some by id.
    
    Pinned spaces cannot be deleted by Revit - these are reported separately.
    Spaces listed in exclude_ids (no linked Room / cannot restore) are preserved.
    
    Args:
        doc: Revit document
        exclude_ids: Optional set of int ElementId values for spaces to skip
        
    Returns:
        dict: Results with 'deleted', 'failed', 'pinned_skipped', 'preserved_unlinked'
    """
    result = {
        'deleted': 0,
        'failed': 0,
        'pinned_skipped': 0,
        'preserved_unlinked': 0
    }
    exclude = exclude_ids if exclude_ids else set()
    
    try:
        spaces = FilteredElementCollector(doc).OfClass(SpatialElement).ToElements()
        space_elements = [s for s in spaces if isinstance(s, Space)]
        
        if space_elements:
            for space in space_elements:
                try:
                    sid = get_element_id_value(space.Id)
                    if sid in exclude:
                        result['preserved_unlinked'] += 1
                        continue
                    # Check if pinned - Revit won't let us delete pinned elements
                    if space.Pinned:
                        result['pinned_skipped'] += 1
                        continue
                    
                    doc.Delete(space.Id)
                    result['deleted'] += 1
                except Exception as e:
                    result['failed'] += 1
                    logger.debug("Error deleting space {}: {}".format(space.Id, str(e)))
                    
    except Exception as e:
        logger.error("Error deleting spaces: {}".format(str(e)))
    
    return result


def create_spaces_from_linked_rooms(doc, linked_item, write_params=True, 
                                     remove_existing=False, progress_bar=None):
    """Create MEP spaces from rooms in a linked model.
    
    IMPORTANT: This function should be called WITHOUT an active transaction.
    It manages its own transactions because NewSpace() uses the active view's 
    phase at transaction start, not the Phase parameter passed to it.
    
    Args:
        doc: Host Revit document
        linked_item: LinkedModelItem with the source linked model
        write_params: Whether to copy Name and Number from Room to Space
        remove_existing: Whether to delete existing spaces first
        progress_bar: Optional pyRevit progress bar
        
    Returns:
        dict: Results with counts and details
    """
    results = {
        'created': 0,
        'created_space_ids': [],  # Track created space IDs for re-tagging
        'skipped_no_level': 0,
        'skipped_no_phase': 0,
        'skipped_failed': 0,
        'deleted': 0,
        'delete_failed': 0,
        'pinned_skipped': 0,
        'preserved_unlinked': 0,
        'params_restored': 0,
        'spaces_matched': 0,
        'errors': [],
        'level_warnings': [],
        'phase_warnings': []
    }
    param_cache = {}
    
    link_instance = linked_item.link_instance
    link_doc = linked_item.link_doc
    
    # Get link transform for coordinate conversion
    link_transform = link_instance.GetTransform()
    
    # Get host levels by elevation for matching (like Revit's "Place Spaces Automatically")
    host_levels = get_host_levels_by_elevation(doc)
    
    # Get all rooms from linked document (placed rooms only)
    _, _, placed_rooms, room_stats = _collect_linked_rooms(link_doc)
    logger.debug("Creating spaces from selected link '{}'".format(linked_item.display_name))
    _log_found_rooms(linked_item.display_name, link_doc, placed_rooms, room_stats)
    
    total_rooms = len(placed_rooms)
    
    if progress_bar:
        progress_bar.update_progress(0, total_rooms)
    
    # Helper to configure transaction with warning suppression using pyRevit's FailureSwallower
    failure_swallower = revit_failure.FailureSwallower()
    def start_transaction_with_warning_suppression(trans):
        """Configure transaction to suppress warnings and start it."""
        opts = trans.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(failure_swallower)
        trans.SetFailureHandlingOptions(opts)
        trans.Start()
    
    # Delete existing spaces if requested (in its own transaction)
    if remove_existing:
        from spaces.params import (
            capture_space_parameters,
            restore_space_parameters,
            rematch_unlinked_spaces_to_rooms,
        )
        if progress_bar:
            progress_bar.update_progress(0, total_rooms)
        # Capture parameters from spaces linked to Rooms; preserve spaces without Space.Room
        existing_spaces = [
            s for s in FilteredElementCollector(doc).OfClass(SpatialElement).ToElements()
            if isinstance(s, Space)
        ]
        try:
            param_cache, unlinked_exclude = capture_space_parameters(existing_spaces)
        except Exception as e:
            logger.error("Error capturing space parameters: {}".format(str(e)))
            param_cache = {}
            # Preserve any space whose linked Room cannot be resolved so we
            # never silently delete user data when capture fails mid-way.
            unlinked_exclude = set()
            try:
                for s in existing_spaces:
                    try:
                        if s.Room is None:
                            unlinked_exclude.add(get_element_id_value(s.Id))
                    except Exception:
                        unlinked_exclude.add(get_element_id_value(s.Id))
            except Exception:
                pass
        # Space.Room is often None when rooms live only in the link. Rematch
        # those spaces to source rooms by number/level/phase or location so
        # recreate can delete them instead of stacking unenclosed copies.
        try:
            rooms_by_key = {}
            rooms_by_loc = []
            for room in placed_rooms:
                nparam = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                nval = nparam.AsString() if nparam else None
                lname = room.Level.Name if room.Level else None
                pparam = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
                pname = None
                if pparam:
                    ph = link_doc.GetElement(pparam.AsElementId())
                    pname = ph.Name if ph else None
                if nval and lname:
                    rooms_by_key[(nval, lname, pname)] = room
                    if (nval, lname, None) not in rooms_by_key:
                        rooms_by_key[(nval, lname, None)] = room
                if room.Location:
                    hp = link_transform.OfPoint(room.Location.Point)
                    rooms_by_loc.append((hp.X, hp.Y, hp.Z, room))
            rematch_unlinked_spaces_to_rooms(
                existing_spaces, unlinked_exclude, param_cache,
                rooms_by_key, rooms_by_loc)
        except Exception as e:
            logger.error("Error rematching unlinked spaces: {}".format(str(e)))
        t = Transaction(doc, "Delete Existing Spaces")
        start_transaction_with_warning_suppression(t)
        delete_results = delete_existing_spaces(doc, exclude_ids=unlinked_exclude)
        results['deleted'] = delete_results['deleted']
        results['delete_failed'] = delete_results['failed']
        results['pinned_skipped'] = delete_results['pinned_skipped']
        results['preserved_unlinked'] = delete_results.get('preserved_unlinked', 0)
        t.Commit()
    
    # Get host phases ONCE before the loop (performance optimization)
    host_phases = get_host_phases(doc)
    host_phases_dict = {phase.Name: phase for phase in host_phases}  # Dict for O(1) lookup
    
    # Group rooms by phase for batch processing
    rooms_by_phase = {}
    level_warnings_dict = {}
    phase_warnings_dict = {}
    
    for room in placed_rooms:
        try:
            # Get room phase
            room_phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
            room_phase_id = room_phase_param.AsElementId() if room_phase_param else None
            room_phase = link_doc.GetElement(room_phase_id) if room_phase_id else None
            room_phase_name = room_phase.Name if room_phase else None
            
            if not room_phase_name:
                results['skipped_no_phase'] += 1
                continue
            
            # Check if host has matching phase
            host_phase = host_phases_dict.get(room_phase_name)
            if not host_phase:
                if room_phase_name not in phase_warnings_dict:
                    phase_warnings_dict[room_phase_name] = 0
                phase_warnings_dict[room_phase_name] += 1
                results['skipped_no_phase'] += 1
                continue
            
            # Get room location
            room_location = room.Location
            if not room_location:
                results['skipped_failed'] += 1
                continue
            room_point = room_location.Point
            
            # Transform room point to host coordinates
            host_point = link_transform.OfPoint(room_point)
            
            # Find matching host level by elevation
            host_level = find_level_by_elevation(host_levels, room.Level.Elevation if room.Level else None)
            if not host_level:
                level_name = room.Level.Name if room.Level else "Unknown"
                if level_name not in level_warnings_dict:
                    level_warnings_dict[level_name] = 0
                level_warnings_dict[level_name] += 1
                results['skipped_no_level'] += 1
                continue
            
            # Add to phase group
            if room_phase_name not in rooms_by_phase:
                rooms_by_phase[room_phase_name] = []
            rooms_by_phase[room_phase_name].append({
                'room': room,
                'host_level': host_level,
                'host_phase': host_phase,
                'host_point': host_point
            })
        except Exception as e:
            results['skipped_failed'] += 1
            results['errors'].append("Error preprocessing room: {}".format(str(e)))
    
    logger.debug(
        "Selected link rooms after filters: ready={} skipped_no_phase={} "
        "skipped_no_level={} skipped_failed={} by_phase={}".format(
            sum(len(v) for v in rooms_by_phase.values()),
            results['skipped_no_phase'],
            results['skipped_no_level'],
            results['skipped_failed'],
            dict((name, len(items)) for name, items in rooms_by_phase.items())))
    if phase_warnings_dict:
        logger.debug("  rooms skipped, host missing phase: {}".format(phase_warnings_dict))
    if level_warnings_dict:
        logger.debug("  rooms skipped, no host level match: {}".format(level_warnings_dict))
    
    # Create temporary views for each phase we need
    # We need to ACTIVATE a view with the correct phase before creating spaces
    # because NewSpace() ignores the Phase parameter and uses the active view's phase
    from Autodesk.Revit.DB import ViewPlan, ViewFamily, ViewFamilyType
    
    # Get a floor plan view family type to use for creating views
    view_family_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    floor_plan_type = None
    for vft in view_family_types:
        if vft.ViewFamily == ViewFamily.FloorPlan:
            floor_plan_type = vft
            break
    
    # Get first level for creating views (host_levels is list of tuples: (elevation, Level))
    first_level = host_levels[0][1] if host_levels else None
    
    # Resolve a floor plan per phase so NewSpace() picks up the view phase.
    # Prefer existing plans (avoids ViewPlan.Create "Sequence contains no elements").
    # Only create a temp view for phases that have no existing plan.
    phase_to_view = {}
    temp_views_created = []
    phases_needed = dict((name, host_phases_dict[name]) for name in rooms_by_phase)

    def _phase_id_value(phase_or_id):
        try:
            if hasattr(phase_or_id, "Id"):
                return get_element_id_value(phase_or_id.Id)
            return get_element_id_value(phase_or_id)
        except Exception:
            return None

    plans_by_phase_id = {}
    try:
        active_id = uidoc.ActiveView.Id if uidoc.ActiveView else None
        for view in FilteredElementCollector(doc).OfClass(ViewPlan).ToElements():
            try:
                if getattr(view, "IsTemplate", False):
                    continue
                view_phase_param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
                if not view_phase_param:
                    continue
                pid = _phase_id_value(view_phase_param.AsElementId())
                if pid is None:
                    continue
                if pid not in plans_by_phase_id:
                    plans_by_phase_id[pid] = view
                elif active_id and view.Id == active_id:
                    plans_by_phase_id[pid] = view
            except Exception:
                continue
    except Exception as e:
        logger.debug("Error collecting phase views: {}".format(str(e)))

    for phase_name, host_phase in phases_needed.items():
        pid = _phase_id_value(host_phase)
        if pid in plans_by_phase_id:
            phase_to_view[phase_name] = plans_by_phase_id[pid]

    missing = [n for n in phases_needed if n not in phase_to_view]
    if missing and floor_plan_type and first_level:
        t_views = Transaction(doc, "Create Temp Views for Phases")
        try:
            start_transaction_with_warning_suppression(t_views)
            for phase_name in missing:
                host_phase = phases_needed[phase_name]
                try:
                    new_view = ViewPlan.Create(doc, floor_plan_type.Id, first_level.Id)
                    try:
                        new_view.Name = "_TempSpaceCreation_{}".format(phase_name)
                    except Exception as e:
                        logger.debug("Could not rename temp view for {}: {}".format(
                            phase_name, str(e)))
                    view_phase_param = new_view.get_Parameter(BuiltInParameter.VIEW_PHASE)
                    if view_phase_param and not view_phase_param.IsReadOnly:
                        view_phase_param.Set(host_phase.Id)
                    set_id = _phase_id_value(view_phase_param.AsElementId()) if view_phase_param else None
                    if set_id == _phase_id_value(host_phase):
                        phase_to_view[phase_name] = new_view
                        temp_views_created.append(new_view.Id)
                except Exception as e:
                    logger.debug("Could not create temp view for {}: {}".format(
                        phase_name, str(e)))
            t_views.Commit()
        except Exception as e:
            logger.debug("Temp view transaction failed: {}".format(str(e)))
            try:
                if t_views.HasStarted() and not t_views.HasEnded():
                    t_views.RollBack()
            except Exception:
                pass
    
    # Store original active view to restore later
    original_active_view = uidoc.ActiveView
    
    # Process each phase group
    processed_count = 0
    progress_update_interval = 50
    last_progress_time = time.time()
    progress_time_interval = 2.0
    
    for phase_name, room_data_list in rooms_by_phase.items():
        host_phase = host_phases_dict[phase_name]
        
        # Try to activate a view with the correct phase
        phase_view = phase_to_view.get(phase_name)
        if phase_view and phase_view.Id != uidoc.ActiveView.Id:
            try:
                uidoc.ActiveView = phase_view
            except Exception:
                pass  # Continue with current view if activation fails
        
        # Create all spaces for this phase
        batch_idx = 0
        t2 = Transaction(doc, "Create Spaces - {}".format(phase_name))
        start_transaction_with_warning_suppression(t2)
        
        for room_data in room_data_list:
            room = room_data['room']
            host_level = room_data['host_level']
            host_point = room_data['host_point']
            
            # Create UV from XYZ (2D point on level)
            uv_point = UV(host_point.X, host_point.Y)
            
            try:
                # Create space using the standard method - the active view should now have the correct phase
                new_space = doc.Create.NewSpace(host_level, host_phase, uv_point)
                
                if new_space:
                    results['created'] += 1
                    results['created_space_ids'].append(new_space.Id)  # Track for re-tagging
                    batch_idx += 1
                    
                    # Write parameters if requested
                    if write_params:
                        try:
                            # Get room name and number
                            room_name = room.get_Parameter(BuiltInParameter.ROOM_NAME)
                            room_number = room.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                            
                            # Set space name
                            if room_name and room_name.AsString():
                                space_name_param = new_space.get_Parameter(BuiltInParameter.ROOM_NAME)
                                if space_name_param and not space_name_param.IsReadOnly:
                                    space_name_param.Set(room_name.AsString())
                            
                            # Set space number
                            if room_number and room_number.AsString():
                                space_number_param = new_space.get_Parameter(BuiltInParameter.ROOM_NUMBER)
                                if space_number_param and not space_number_param.IsReadOnly:
                                    space_number_param.Set(room_number.AsString())
                                    
                        except Exception as e:
                            logger.debug("Error setting space parameters: {}".format(str(e)))
                    
                    # Restore parameters from old space (keyed by linked Room UniqueId)
                    if remove_existing and param_cache and new_space:
                        try:
                            room_uid = room.UniqueId
                            if room_uid in param_cache:
                                n_rest = restore_space_parameters(
                                    new_space, param_cache[room_uid])
                                results['params_restored'] += n_rest
                                results['spaces_matched'] += 1
                        except Exception as e:
                            logger.debug("Error restoring space parameters: {}".format(str(e)))
                else:
                    results['skipped_failed'] += 1
                    
            except Exception as e:
                results['skipped_failed'] += 1
                results['errors'].append("Failed to create space: {}".format(str(e)))
            
            processed_count += 1
            
            # Throttled progress update
            if progress_bar:
                should_update = False
                if processed_count % progress_update_interval == 0:
                    should_update = True
                elif time.time() - last_progress_time >= progress_time_interval:
                    should_update = True
                elif processed_count == total_rooms:
                    should_update = True
                
                if should_update:
                    progress_bar.update_progress(processed_count, total_rooms)
                    last_progress_time = time.time()
        
        # Commit the batch transaction
        t2.Commit()
    
    # Restore original active view
    if original_active_view and original_active_view.Id != uidoc.ActiveView.Id:
        try:
            uidoc.ActiveView = original_active_view
        except Exception:
            pass  # Ignore errors restoring view
    
    # Delete temporary views we created (no warning suppression needed for view deletion)
    if temp_views_created:
        t_cleanup = Transaction(doc, "Delete Temp Views")
        t_cleanup.Start()
        deleted_views = 0
        for view_id in temp_views_created:
            try:
                doc.Delete(view_id)
                deleted_views += 1
            except:
                pass
        t_cleanup.Commit()
    
    # Convert warning dicts back to lists for results
    results['level_warnings'] = [(k, v) for k, v in level_warnings_dict.items()]
    results['phase_warnings'] = [(k, v) for k, v in phase_warnings_dict.items()]
    
    logger.debug(
        "Create Spaces finished for '{}': found_placed={} created={} "
        "skipped_no_phase={} skipped_no_level={} skipped_failed={}".format(
            linked_item.display_name,
            total_rooms,
            results['created'],
            results['skipped_no_phase'],
            results['skipped_no_level'],
            results['skipped_failed']))
    
    return results


def show_results(results):
    """Show results summary.
    
    Args:
        results: Results dictionary from create_spaces_from_linked_rooms
    """
    message_parts = []
    
    # Report on deleted spaces
    if results['deleted'] > 0:
        message_parts.append("{} existing spaces deleted".format(results['deleted']))
    
    # Report on pinned spaces that couldn't be deleted
    if results.get('pinned_skipped', 0) > 0:
        message_parts.append("{} pinned spaces could not be deleted".format(results['pinned_skipped']))
    
    # Report on other failed deletions
    if results.get('delete_failed', 0) > 0:
        message_parts.append("{} spaces failed to delete".format(results['delete_failed']))
    
    message_parts.append("{} spaces created".format(results['created']))
    
    if results.get('preserved_unlinked', 0) > 0:
        message_parts.append(
            "{} spaces preserved (no Room link from Space.Room)".format(
                results['preserved_unlinked']))
    
    if results.get('spaces_matched', 0) > 0:
        message_parts.append(
            "{} spaces recreated with parameters preserved ({} parameter values restored)".format(
                results['spaces_matched'], results.get('params_restored', 0)))
    elif results.get('params_restored', 0) > 0:
        message_parts.append(
            "{} parameter values restored".format(results['params_restored']))
    
    # Show tagging results if re-tagging was performed
    if results.get('tagged', 0) > 0:
        message_parts.append("{} spaces tagged in {} views".format(
            results['tagged'], results.get('views_tagged', 0)))
    
    if results['skipped_no_level'] > 0:
        message_parts.append("{} skipped (no matching level)".format(results['skipped_no_level']))
    
    if results['skipped_no_phase'] > 0:
        message_parts.append("{} skipped (no matching phase)".format(results['skipped_no_phase']))
    
    if results['skipped_failed'] > 0:
        message_parts.append("{} failed".format(results['skipped_failed']))
    
    message = "\n".join(message_parts)
    
    # Show level warnings if any (now shows elevation info)
    if results['level_warnings']:
        message += "\n\nNo matching level elevation in host model:"
        for level_info, count in results['level_warnings']:
            message += "\n  - {} ({} rooms)".format(level_info, count)
    
    # Show phase warnings if any
    if results['phase_warnings']:
        message += "\n\nMissing phases in host model:"
        for phase_name, count in results['phase_warnings']:
            message += "\n  - '{}' ({} rooms)".format(phase_name, count)
    
    # Build balloon text
    balloon_lines = []
    
    if results['deleted'] > 0:
        balloon_lines.append("{} deleted".format(results['deleted']))
    
    if results.get('pinned_skipped', 0) > 0:
        balloon_lines.append("{} pinned (not deleted)".format(results['pinned_skipped']))
    
    balloon_lines.append("{} created".format(results['created']))
    
    if results.get('preserved_unlinked', 0) > 0:
        balloon_lines.append("{} preserved (no Room link)".format(results['preserved_unlinked']))
    
    if results.get('spaces_matched', 0) > 0:
        balloon_lines.append("{} w/ params restored".format(results['spaces_matched']))
    
    if results.get('tagged', 0) > 0:
        balloon_lines.append("{} tagged".format(results['tagged']))
    
    if results['skipped_no_level'] > 0:
        balloon_lines.append("{} skipped (no matching level)".format(
            results['skipped_no_level']))
    
    if results['skipped_no_phase'] > 0:
        balloon_lines.append("{} skipped (no matching phase)".format(
            results['skipped_no_phase']))
    
    if results['skipped_failed'] > 0:
        balloon_lines.append("{} failed".format(results['skipped_failed']))
    
    balloon_text = ", ".join(balloon_lines) + "."
    
    if results['phase_warnings']:
        missing_phases = ", ".join(
            "'{}'".format(name) for name, _count in results['phase_warnings'])
        balloon_text += (
            " Host is missing phase(s): {}. "
            "Phase names in the host must match the linked model."
        ).format(missing_phases)
    
    forms.show_balloon(
        header="Spaces",
        text=balloon_text,
        is_new=True
    )


class CreateSpacesWindow(WPFWindow):
    """WPF Window for creating spaces from linked model rooms."""
    
    def __init__(self, xaml_file):
        """Initialize the window.
        
        Args:
            xaml_file: Path to the XAML file
        """
        # Initialize attributes that may be accessed in OnClosing
        self.linked_models = []
        
        try:
            # Initialize WPF window
            WPFWindow.__init__(self, xaml_file)
            
            # Load styles AFTER window initialization (window-scoped, isolated from Revit UI)
            from styles import load_styles_to_window
            load_styles_to_window(self)
            
            # Hide busy overlay initially
            self.busyOverlay.Visibility = Visibility.Collapsed
            
            # Load linked models with rooms
            self.load_linked_models()
            
            # Set up ComboBox selection changed event
            self.linkedModelComboBox.SelectionChanged += self.on_linked_model_changed
            
        except Exception as e:
            logger.error("Error initializing window: {}".format(str(e)))
            forms.alert("Failed to initialize: {}".format(str(e)), exitscript=True)
    
    def set_busy(self, is_busy, message="Creating spaces..."):
        """Show or hide the busy overlay indicator.
        
        Args:
            is_busy: Whether to show the busy indicator
            message: Message to display
        """
        try:
            if is_busy:
                self.busyOverlay.Visibility = Visibility.Visible
                self.busyTextBlock.Text = message
                self.createButton.IsEnabled = False
            else:
                self.busyOverlay.Visibility = Visibility.Collapsed
                self.createButton.IsEnabled = True
        except Exception as e:
            logger.debug("Error setting busy indicator: {}".format(str(e)))
    
    def load_linked_models(self):
        """Load linked models with rooms into the ComboBox."""
        try:
            self.linked_models = get_linked_documents_with_rooms(doc)
            
            if not self.linked_models:
                self.createButton.IsEnabled = False
                return
            
            self.linkedModelComboBox.ItemsSource = self.linked_models
            self.linkedModelComboBox.SelectedIndex = 0
            
        except Exception as e:
            logger.error("Error loading linked models: {}".format(str(e)))
    
    def on_linked_model_changed(self, sender, args):
        """Handle linked model selection change."""
        try:
            selected = self.linkedModelComboBox.SelectedItem
            if selected:
                logger.debug("Selected linked model: {} ({} rooms)".format(
                    selected.display_name, selected.room_count))
        except Exception as e:
            logger.debug("Error handling selection change: {}".format(str(e)))
    
    def RetagCheckBox_Changed(self, sender, args):
        """Enable/disable tag type selector based on checkbox state."""
        try:
            is_checked = self.retagSpacesCheckBox.IsChecked
            self.tagTypeComboBox.IsEnabled = is_checked
            
            # Load tag types on first enable if not already loaded
            if is_checked and self.tagTypeComboBox.ItemsSource is None:
                self.load_tag_types()
        except Exception as e:
            logger.debug("Error handling retag checkbox change: {}".format(str(e)))
    
    def load_tag_types(self):
        """Load space tag types into the ComboBox."""
        try:
            from spaces import get_space_tag_types
            tag_types = get_space_tag_types(doc)
            if tag_types:
                self.tagTypeComboBox.ItemsSource = tag_types
                self.tagTypeComboBox.SelectedIndex = 0
            else:
                logger.warning("No space tag types found in project")
        except Exception as e:
            logger.error("Error loading tag types: {}".format(str(e)))
    
    def CreateButton_Click(self, sender, args):
        """Handle Create button click."""
        try:
            selected = self.linkedModelComboBox.SelectedItem
            if not selected:
                forms.alert("Please select a linked model.")
                return
            
            write_params = self.writeParamsCheckBox.IsChecked
            remove_existing = self.removeExistingCheckBox.IsChecked
            retag_spaces = self.retagSpacesCheckBox.IsChecked
            selected_tag_type = self.tagTypeComboBox.SelectedItem
            logger.debug(
                "Create Spaces clicked: link='{}' ui_room_count={} "
                "write_params={} remove_existing={}".format(
                    selected.display_name,
                    selected.room_count,
                    bool(write_params),
                    bool(remove_existing)))
            
            # Validate tag type selection if re-tagging is enabled
            if retag_spaces and not selected_tag_type:
                forms.alert("Please select a tag type for re-tagging.", title="No Tag Type")
                return
            
            # Close window before showing progress bar (forms.ProgressBar is modal)
            self.Close()
            
            # Run the space creation with progress bar
            # Note: create_spaces_from_linked_rooms manages its own transactions
            # because NewSpace() uses the view's phase at transaction start
            try:
                # If re-tagging is enabled and we're removing existing spaces,
                # we need to capture which views have space tags BEFORE deleting
                views_with_tags_before = None
                if retag_spaces and remove_existing:
                    from spaces import get_views_with_space_tags
                    # Fast lookup - no progress bar needed (~50ms)
                    views_with_tags_before = get_views_with_space_tags(doc)
                
                with forms.ProgressBar(title="Creating Spaces...") as pb:
                    results = create_spaces_from_linked_rooms(
                        doc,
                        selected,
                        write_params=write_params,
                        remove_existing=remove_existing,
                        progress_bar=pb
                    )
                
                # Re-tag new spaces if option is enabled
                if retag_spaces and results['created'] > 0 and selected_tag_type:
                    self._retag_new_spaces(results, selected_tag_type, views_with_tags_before)
                
                # Show results (same format as Update Spaces)
                show_results(results)
                    
            except Exception as e:
                forms.alert("Error creating spaces: {}".format(str(e)))
                logger.error("Error creating spaces: {}".format(str(e)))
                
        except Exception as e:
            forms.alert("Error: {}".format(str(e)))
            logger.error("Create button error: {}".format(str(e)))
    
    def _retag_new_spaces(self, results, selected_tag_type, views_with_tags_before=None):
        """Re-tag newly created spaces in views that have existing space tags.
        
        Args:
            results: Results dict from create_spaces_from_linked_rooms (will be modified)
            selected_tag_type: TagTypeItem with the selected tag type
            views_with_tags_before: Optional dict of views that had space tags before deletion.
                                    If provided, uses these views instead of looking up current tags.
        """
        try:
            from spaces import get_views_with_space_tags, tag_spaces_in_view
            
            # Use pre-captured views if provided (when spaces were deleted before creation)
            # Otherwise, find views that currently have space tags
            if views_with_tags_before is not None:
                views_with_tags = views_with_tags_before
            else:
                # Find views that have existing space tags (fast lookup - no progress bar needed)
                views_with_tags = get_views_with_space_tags(doc)
            
            if not views_with_tags:
                logger.debug("No views with existing space tags found - skipping re-tag")
                forms.alert(
                    "No views with existing space tags were found.\n\n"
                    "To use re-tagging, first tag spaces in desired views using 'Tag All Spaces', "
                    "then run 'Create Spaces' with re-tagging enabled.\n\n"
                    "The views that had tags will be remembered and new spaces will be tagged in them.",
                    title="Re-tagging Skipped"
                )
                return
            
            tag_type_id = selected_tag_type.element_id
            new_space_ids = set(results['created_space_ids'])
            
            logger.debug("Re-tagging {} new spaces in {} views".format(
                len(new_space_ids), len(views_with_tags)))
            
            # Tag newly created spaces in views that already have tags (with progress bar)
            total_tagged = 0
            view_ids_list = list(views_with_tags.keys())
            total_views = len(view_ids_list)
            
            with forms.ProgressBar(title="Re-tagging spaces in views...", cancellable=True) as pb:
                with revit.Transaction("Re-tag New Spaces"):
                    # Activate tag type if not already
                    tag_symbol = doc.GetElement(tag_type_id)
                    if tag_symbol and not tag_symbol.IsActive:
                        tag_symbol.Activate()
                        doc.Regenerate()
                    
                    for i, view_id in enumerate(view_ids_list):
                        if pb.cancelled:
                            break
                        
                        pb.update_progress(i, total_views)
                        
                        view = doc.GetElement(view_id)
                        if view:
                            tagged, _, _ = tag_spaces_in_view(doc, view, tag_type_id, 
                                                              space_ids=new_space_ids)
                            total_tagged += tagged
                    
                    pb.update_progress(total_views, total_views)
            
            # Add tagging results
            results['tagged'] = total_tagged
            results['views_tagged'] = len(views_with_tags)
            
            logger.debug("Re-tagged {} spaces in {} views".format(
                total_tagged, len(views_with_tags)))
                
        except Exception as e:
            logger.error("Error re-tagging spaces: {}".format(str(e)))
            results['errors'].append("Re-tagging error: {}".format(str(e)))
    
    def CancelButton_Click(self, sender, args):
        """Handle Cancel button click."""
        self.Close()


# Import System.Windows.Threading for dispatcher
import System.Windows.Threading

# Main execution
if __name__ == '__main__':
    # Check for linked models with rooms
    linked_models = get_linked_documents_with_rooms(doc)
    
    if not linked_models:
        forms.alert("No linked models with rooms found.\n\n"
                   "Please ensure you have linked Revit models that contain placed rooms.",
                   title="No Linked Rooms",
                   exitscript=True)
    
    # Show the window
    xaml_file = op.join(script_dir, 'CreateSpacesWindow.xaml')
    window = CreateSpacesWindow(xaml_file)
    window.ShowDialog()
