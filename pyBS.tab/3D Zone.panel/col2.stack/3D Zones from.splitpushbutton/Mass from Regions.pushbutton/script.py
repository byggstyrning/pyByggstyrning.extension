# -*- coding: utf-8 -*-
"""Create Mass family instances from FilledRegion boundaries.

Creates Mass family instances using the MassZone.rfa template,
replacing the extrusion profile with each filled region's boundary loops.
Instances are placed in the active view respecting the view's phase.

NOTE: This tool requires a MassZone.rfa template created from a Conceptual Mass
family in Revit. The template should contain a simple Form element.
"""

__title__ = "Mass from\nRegions"
__author__ = "Byggstyrning AB"
__doc__ = "Create Mass family instances from FilledRegion boundaries using MassZone.rfa template"
__highlight__ = 'new'

# Import standard libraries
import sys
import os.path as op

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Add the extension directory to the path
script_dir = script.get_script_path()
# Calculate extension directory by going up from script directory
# Structure: extension_root/pyBS.tab/3D Zone.panel/col2.stack/3D Zones from.splitpushbutton/Mass from Regions.pushbutton/
pushbutton_dir = script_dir
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Initialize logger
logger = script.get_logger()

# Import shared modules
from zone3d.spatial_adapter import RegionAdapter
from zone3d.mass_creator import create_masses_from_spatial_elements


def show_region_filter_dialog(regions, doc):
    """Simple filter function that returns selected regions (no dialog).
    
    This function mimics the interface expected by create_masses_from_spatial_elements
    but doesn't show a dialog - just returns the regions directly.
    
    Args:
        regions: List of FilledRegion elements
        doc: Revit document
        
    Returns:
        List of FilledRegion elements (same as input)
    """
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
    
    # Get view from first region for phase handling
    active_view = None
    if filled_regions:
        try:
            first_region = filled_regions[0]
            view_id = first_region.OwnerViewId
            if view_id and view_id != ElementId.InvalidElementId:
                active_view = doc.GetElement(view_id)
        except Exception as e:
            logger.debug("Error getting view from region: {}".format(e))
    
    # Fallback to active view
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
        logger.debug("Using view '{}' for phase handling".format(
            active_view.Name if hasattr(active_view, 'Name') else 'Unknown'))
    
    # Create adapter with view for phase handling
    adapter = RegionAdapter(active_view=active_view)
    
    # Create Mass elements using orchestration function
    success_count, fail_count, failed_elements, created_instance_ids = create_masses_from_spatial_elements(
        filled_regions,
        doc,
        adapter,
        extension_dir,
        pushbutton_dir,
        show_region_filter_dialog,
        "Regions",
        "MassZone.rfa"
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
            logger.debug("Selected {} newly created Mass instance(s)".format(len(created_instance_ids)))
        except Exception as select_error:
            logger.debug("Error selecting created instances: {}".format(select_error))
    
    # Show summary
    if success_count > 0:
        forms.alert("Created {} Mass element(s) from FilledRegion boundaries.\n\n{} failed.".format(
            success_count, fail_count), title="Mass Creation Complete")
    elif fail_count > 0:
        forms.alert("Failed to create Mass elements.\n\nCheck the output window for details.",
                   title="Mass Creation Failed", warn_icon=True)
