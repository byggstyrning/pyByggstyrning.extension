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
- Remove existing spaces before creating
"""
__highlight__ = 'new'

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
        
        for link in link_instances:
            try:
                link_doc = link.GetLinkDocument()
                if not link_doc:
                    continue
                
                # Get all rooms in linked document
                rooms = FilteredElementCollector(link_doc).OfClass(SpatialElement).ToElements()
                placed_rooms = [r for r in rooms if isinstance(r, Room) and r.Area > 0]
                
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


def delete_existing_spaces(doc):
    """Delete all existing spaces in the document.
    
    Args:
        doc: Revit document
        
    Returns:
        int: Number of spaces deleted
    """
    deleted_count = 0
    try:
        spaces = FilteredElementCollector(doc).OfClass(SpatialElement).ToElements()
        space_ids = [s.Id for s in spaces if isinstance(s, Space)]
        
        if space_ids:
            for space_id in space_ids:
                try:
                    doc.Delete(space_id)
                    deleted_count += 1
                except Exception:
                    pass
    except Exception as e:
        logger.error("Error deleting spaces: {}".format(str(e)))
    
    return deleted_count


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
        'skipped_no_level': 0,
        'skipped_no_phase': 0,
        'skipped_failed': 0,
        'deleted': 0,
        'errors': [],
        'level_warnings': [],
        'phase_warnings': []
    }
    
    link_instance = linked_item.link_instance
    link_doc = linked_item.link_doc
    
    # Get link transform for coordinate conversion
    link_transform = link_instance.GetTransform()
    
    # Get host levels by elevation for matching (like Revit's "Place Spaces Automatically")
    host_levels = get_host_levels_by_elevation(doc)
    
    # Get all rooms from linked document (placed rooms only)
    all_rooms = FilteredElementCollector(link_doc).OfClass(SpatialElement).ToElements()
    placed_rooms = [r for r in all_rooms if isinstance(r, Room) and r.Area > 0]
    
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
        if progress_bar:
            progress_bar.update_progress(0, total_rooms)
        t = Transaction(doc, "Delete Existing Spaces")
        start_transaction_with_warning_suppression(t)
        results['deleted'] = delete_existing_spaces(doc)
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
    
    # Create a temporary view for each phase
    phase_to_view = {}
    temp_views_created = []
    
    if floor_plan_type and first_level:
        t_views = Transaction(doc, "Create Temp Views for Phases")
        start_transaction_with_warning_suppression(t_views)
        
        for phase_name, host_phase in host_phases_dict.items():
            try:
                # Create a new floor plan view
                new_view = ViewPlan.Create(doc, floor_plan_type.Id, first_level.Id)
                new_view.Name = "_TempSpaceCreation_{}".format(phase_name)
                
                # Set the view's phase
                view_phase_param = new_view.get_Parameter(BuiltInParameter.VIEW_PHASE)
                if view_phase_param and not view_phase_param.IsReadOnly:
                    view_phase_param.Set(host_phase.Id)
                
                phase_to_view[phase_name] = new_view
                temp_views_created.append(new_view.Id)
            except Exception:
                pass  # Skip phases where we can't create a view
        
        t_views.Commit()
    
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
    
    return results


def show_results(results):
    """Show results summary.
    
    Args:
        results: Results dictionary from create_spaces_from_linked_rooms
    """
    message_parts = []
    
    if results['deleted'] > 0:
        message_parts.append("{} existing spaces removed".format(results['deleted']))
    
    message_parts.append("{} spaces created".format(results['created']))
    
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
    
    # Use balloon for success, alert for issues
    has_skips = (results['skipped_no_level'] > 0 or 
                 results['skipped_no_phase'] > 0 or 
                 results['skipped_failed'] > 0)
    
    if results['created'] > 0 and not has_skips:
        forms.show_balloon(
            header="Spaces Created",
            text="{} spaces created successfully.".format(results['created']),
            is_new=True
        )
    else:
        forms.alert(message, title="Create Spaces Results")


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
            # Load styles before WPFWindow.__init__
            self.load_styles()
            
            # Initialize WPF window
            WPFWindow.__init__(self, xaml_file)
            
            # Hide busy overlay initially
            self.busyOverlay.Visibility = Visibility.Collapsed
            
            # Load linked models with rooms
            self.load_linked_models()
            
            # Set up ComboBox selection changed event
            self.linkedModelComboBox.SelectionChanged += self.on_linked_model_changed
            
        except Exception as e:
            logger.error("Error initializing window: {}".format(str(e)))
            forms.alert("Failed to initialize: {}".format(str(e)), exitscript=True)
    
    def load_styles(self):
        """Load the common styles ResourceDictionary with theme support."""
        try:
            from styles import load_styles_to_window
            result = load_styles_to_window(self)
            if result:
                logger.debug("Loaded styles with theme support")
            else:
                logger.warning("Could not load styles with theme support")
        except ImportError as e:
            logger.error("Failed to import styles: {}".format(str(e)))
        except Exception as e:
            logger.warning("Could not load styles: {}".format(str(e)))
    
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
    
    def CreateButton_Click(self, sender, args):
        """Handle Create button click."""
        try:
            selected = self.linkedModelComboBox.SelectedItem
            if not selected:
                forms.alert("Please select a linked model.")
                return
            
            write_params = self.writeParamsCheckBox.IsChecked
            remove_existing = self.removeExistingCheckBox.IsChecked
            
            # Close window before showing progress bar (forms.ProgressBar is modal)
            self.Close()
            
            # Run the space creation with progress bar
            # Note: create_spaces_from_linked_rooms manages its own transactions
            # because NewSpace() uses the view's phase at transaction start
            try:
                with forms.ProgressBar(title="Creating Spaces...") as pb:
                    results = create_spaces_from_linked_rooms(
                        doc,
                        selected,
                        write_params=write_params,
                        remove_existing=remove_existing,
                        progress_bar=pb
                    )
                
                # Show results (same format as Update Spaces)
                show_results(results)
                    
            except Exception as e:
                forms.alert("Error creating spaces: {}".format(str(e)))
                logger.error("Error creating spaces: {}".format(str(e)))
                
        except Exception as e:
            forms.alert("Error: {}".format(str(e)))
            logger.error("Create button error: {}".format(str(e)))
    
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
