# -*- coding: utf-8 -*-
"""Create 3D Zone Generic Model family instances from FilledRegion boundaries.

Creates Generic Model family instances using the 3DZone.rfa template,
replacing the extrusion profile with each filled region's boundary loops.
Instances are placed in the active view respecting the view's phase.
"""

__title__ = "Create 3D Zones from Regions"
__author__ = "Byggstyrning AB"
__doc__ = "Create Generic Model family instances from FilledRegion boundaries using 3DZone.rfa template"
__highlight__ = 'new'

# Import standard libraries
import sys
import os.path as op

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Add the extension directory to the path
logger_temp = script.get_logger()
script_dir = script.get_script_path()
# Calculate extension directory by going up from script directory
# Structure: extension_root/pyBS.tab/3D Zone.panel/col2.stack/3D Zones from Region.pushbutton/
pushbutton_dir = script_dir
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)  # Go up one more level to get extension root
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Initialize logger
logger = script.get_logger()

# Import shared modules
from zone3d.spatial_adapter import RegionAdapter
from zone3d.zone_creator import create_zones_from_spatial_elements


def show_region_filter_dialog(regions, doc):
    """Simple filter function that returns selected regions (no dialog).
    
    This function mimics the interface expected by create_zones_from_spatial_elements
    but doesn't show a dialog - just returns the regions directly.
    
    Args:
        regions: List of FilledRegion elements
        doc: Revit document
        
    Returns:
        List of FilledRegion elements (same as input)
    """
    # No dialog needed - just return the regions
    return regions


# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Get selected elements
    selection = revit.get_selection()
    selected_ids = selection.element_ids
    
    if not selected_ids or len(selected_ids) == 0:
        forms.alert("Please select at least one FilledRegion element before running this tool.",
                   title="No Selection", exitscript=True)
    
    # Get selected elements and filter for FilledRegion
    selected_elements = [doc.GetElement(eid) for eid in selected_ids]
    filled_regions = [elem for elem in selected_elements if isinstance(elem, FilledRegion)]
    
    if not filled_regions:
        forms.alert("No FilledRegion elements found in selection.\n\nPlease select FilledRegion elements and try again.",
                   title="Invalid Selection", exitscript=True)
    
    logger.debug("Found {} FilledRegion element(s) in selection".format(len(filled_regions)))
    
    # Get view from first region (all regions should be in the same view)
    # Use this view for phase handling
    active_view = None
    if filled_regions:
        try:
            first_region = filled_regions[0]
            view_id = first_region.OwnerViewId
            if view_id and view_id != ElementId.InvalidElementId:
                active_view = doc.GetElement(view_id)
        except Exception as e:
            logger.debug("Error getting view from region: {}".format(e))
    
    # Fallback to active view if region view not found
    if not active_view:
        try:
            active_view = revit.active_view
            if not active_view:
                active_view = doc.ActiveView
        except:
            pass
    
    if not active_view:
        logger.warning("Could not determine view - phase may not be set correctly")
    else:
        logger.debug("Using view '{}' (ID: {}) for phase handling".format(
            active_view.Name if hasattr(active_view, 'Name') else 'Unknown',
            active_view.Id if active_view else 'None'))
    
    # Create adapter with view for phase handling
    adapter = RegionAdapter(active_view=active_view)
    
    # Create zones using shared orchestration function
    # Note: We bypass the dialog by providing a simple filter function
    success_count, fail_count, failed_elements, created_instance_ids = create_zones_from_spatial_elements(
        filled_regions,
        doc,
        adapter,
        extension_dir,
        pushbutton_dir,
        show_region_filter_dialog,
        "Regions"
    )
    
    # Select newly created instances
    if created_instance_ids:
        try:
            from System.Collections.Generic import List
            uidoc = revit.uidoc
            element_ids = List[ElementId]()
            for instance_id in created_instance_ids:
                element_ids.Add(instance_id)
            
            uidoc.Selection.SetElementIds(element_ids)
            logger.debug("Selected {} newly created zone instance(s)".format(len(created_instance_ids)))
        except Exception as select_error:
            logger.debug("Error selecting created instances: {}".format(select_error))
            # Don't fail the script if selection fails

