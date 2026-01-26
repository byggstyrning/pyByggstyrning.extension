# -*- coding: utf-8 -*-
"""Update MEP Space Name and Number from Room parameters.

This tool copies the "Room Name" and "Room Number" parameter values
to the Space's built-in Name and Number parameters.
"""

__title__ = "Update\nSpaces"
__author__ = "Byggstyrning AB"
__doc__ = """Update MEP Space Name and Number from Room parameters.

For each Space in the current model, this tool copies:
- "Room Name" parameter -> Space Name
- "Room Number" parameter -> Space Number

This updates the built-in Name/Number from the room-bounding
parameters that Revit populates automatically.
"""
__highlight__ = 'new'

# Standard library imports
import time

# .NET imports
import clr
clr.AddReference('RevitAPI')

# Revit API imports
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    SpatialElement,
    BuiltInParameter
)
from Autodesk.Revit.DB.Mechanical import Space

# pyRevit imports
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Initialize logger
logger = script.get_logger()

# Document references
doc = revit.doc


def get_all_spaces(doc):
    """Get all placed MEP Spaces in the document.
    
    Args:
        doc: Revit document
        
    Returns:
        list: List of Space elements with Area > 0
    """
    spaces = []
    try:
        all_spatial = FilteredElementCollector(doc).OfClass(SpatialElement).ToElements()
        spaces = [s for s in all_spatial if isinstance(s, Space) and s.Area > 0]
    except Exception as e:
        logger.error("Error getting spaces: {}".format(str(e)))
    return spaces


def update_spaces_from_room_params(doc, progress_bar=None):
    """Update Space Name and Number from Room Name/Number parameters.
    
    Args:
        doc: Revit document
        progress_bar: Optional pyRevit progress bar
        
    Returns:
        dict: Results with counts
    """
    results = {
        'updated': 0,
        'no_room_data': 0,
        'no_change': 0,
        'errors': 0,
        'total': 0
    }
    
    # Get all spaces
    spaces = get_all_spaces(doc)
    results['total'] = len(spaces)
    
    if not spaces:
        return results
    
    # Process each space
    # Throttle progress updates to reduce UI overhead (update every 50 spaces or every 2 seconds)
    progress_update_interval = 50  # Update every N spaces
    last_progress_time = time.time()
    progress_time_interval = 2.0  # Or every N seconds
    
    for idx, space in enumerate(spaces):
        # Throttle progress callback to reduce UI overhead
        if progress_bar:
            should_update = False
            # Update every N spaces
            if (idx + 1) % progress_update_interval == 0:
                should_update = True
            # Or if enough time has passed
            elif time.time() - last_progress_time >= progress_time_interval:
                should_update = True
            # Always update on last space
            elif idx + 1 == results['total']:
                should_update = True
            
            if should_update:
                progress_bar.update_progress(idx + 1, results['total'])
                last_progress_time = time.time()
        
        try:
            # Get "Room Name" and "Room Number" parameters (from room bounding)
            room_name_param = space.LookupParameter("Room Name")
            room_number_param = space.LookupParameter("Room Number")
            
            room_name = room_name_param.AsString() if room_name_param else None
            room_number = room_number_param.AsString() if room_number_param else None
            
            # Check if we have room data
            if not room_name and not room_number:
                results['no_room_data'] += 1
                continue
            
            # Get current space name and number (built-in parameters)
            space_name_param = space.get_Parameter(BuiltInParameter.ROOM_NAME)
            space_number_param = space.get_Parameter(BuiltInParameter.ROOM_NUMBER)
            
            current_name = space_name_param.AsString() if space_name_param else ""
            current_number = space_number_param.AsString() if space_number_param else ""
            
            # Check if update is needed
            name_changed = room_name and room_name != current_name
            number_changed = room_number and room_number != current_number
            
            if not name_changed and not number_changed:
                results['no_change'] += 1
                continue
            
            # Update space parameters
            updated = False
            
            if name_changed and space_name_param and not space_name_param.IsReadOnly:
                space_name_param.Set(room_name)
                updated = True
            
            if number_changed and space_number_param and not space_number_param.IsReadOnly:
                space_number_param.Set(room_number)
                updated = True
            
            if updated:
                results['updated'] += 1
            else:
                results['no_change'] += 1
                
        except Exception as e:
            results['errors'] += 1
            logger.debug("Error updating space: {}".format(str(e)))
    
    return results


# Main execution
if __name__ == '__main__':
    # Get all spaces
    spaces = get_all_spaces(doc)
    
    if not spaces:
        forms.alert("No MEP Spaces found in the model.\n\n"
                   "Use 'Create Spaces' to create spaces first.",
                   title="No Spaces Found",
                   exitscript=True)
    
    # Confirm action
    confirm_msg = ("This will update Name and Number for {} spaces\n"
                   "from their 'Room Name' and 'Room Number' parameters.\n\n"
                   "Continue?").format(len(spaces))
    if not forms.alert(confirm_msg, yes=True, no=True, title="Update Spaces"):
        script.exit()
    
    # Run the update with progress bar
    with forms.ProgressBar(title="Updating Spaces...") as pb:
        with revit.Transaction("Update Spaces from Room Parameters"):
            results = update_spaces_from_room_params(doc, progress_bar=pb)
    
    # Build result message
    message_parts = []
    message_parts.append("{} spaces processed".format(results['total']))
    message_parts.append("{} updated".format(results['updated']))
    
    if results['no_room_data'] > 0:
        message_parts.append("{} had no room data".format(results['no_room_data']))
    
    if results['no_change'] > 0:
        message_parts.append("{} already matched".format(results['no_change']))
    
    if results['errors'] > 0:
        message_parts.append("{} errors".format(results['errors']))
    
    message = "\n".join(message_parts)
    
    # Show results
    if results['updated'] > 0:
        forms.show_balloon(
            header="Spaces Updated",
            text="{} spaces updated from room parameters.".format(results['updated']),
            is_new=True
        )
    else:
        forms.alert(message, title="Update Spaces Results")
