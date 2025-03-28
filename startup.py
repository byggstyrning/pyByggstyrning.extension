# -*- coding: utf-8 -*-
"""
Startup script for pyByggstyrning extension.
Clones the MMI panel to the Modify tab when a document is opened.
Creates an API endpoint to select an element by ID in a specific 3D view.
"""

import clr
clr.AddReference('AdWindows')
import Autodesk.Windows as AdWindows

from pyrevit import HOST_APP, routes
from System import EventHandler
from Autodesk.Revit.DB import (
    Events, ElementId, Transaction, View3D, ViewFamilyType,
    FilteredElementCollector, BuiltInCategory, BoundingBoxXYZ,
    ViewFamily, XYZ
)
from Autodesk.Revit.UI import UIDocument, Selection
from System.Collections.Generic import List
from pyrevit.coreutils import logger
import time # Keep time for total duration calculation

# Create a logger instance for this script
script_logger = logger.get_logger('switchback_api')

def find_and_clone_mmi_panel():
    ribbon = AdWindows.ComponentManager.Ribbon
    if not ribbon:
        return
    
    # Find the Modify tab
    modify_tab = next((tab for tab in ribbon.Tabs 
                      if tab.Title == "Modify" and tab.IsVisible), None)
    
    if not modify_tab:
        return
        
    # Find the source MMI panel
    source_panel = None
    for tab in ribbon.Tabs:
        if not tab.IsVisible:
            continue
        for panel in tab.Panels:
            if panel.Source and "MMI" in panel.Source.Title:
                source_panel = panel
                break
        if source_panel:
            break
            
    if not source_panel:
        return
        
    # Check if panel already exists in Modify tab
    if any(panel.Source and panel.Source.Title == source_panel.Source.Title 
           for panel in modify_tab.Panels):
        return
        
    # Clone panel to Modify tab
    modify_tab.Panels.Add(AdWindows.RibbonPanel())
    new_panel = modify_tab.Panels[modify_tab.Panels.Count - 1]
    new_panel.Source = source_panel.Source.Clone()
    new_panel.IsEnabled = True

# Run on startup / reload
find_and_clone_mmi_panel()

# Run on document open on event
def doc_opening_handler(sender, args):
    find_and_clone_mmi_panel()

# Register the event handler to load the MMI panel when a document is opened
HOST_APP.app.DocumentOpening += \
    EventHandler[Events.DocumentOpeningEventArgs](
        doc_opening_handler
    )

# Initialize API
api = routes.API('switchback')

# --- Constants ---
SWITCHBACK_VIEW_NAME = "Switchback"

@api.route('/id/<int:element_id_param>')
def select_by_id(doc, request, element_id_param):
    """
    Selects an element by ID. Finds/creates a 'Switchback' 3D view,
    activates it, applies a section box, selects the element, and zooms.
    Expects URL like /id/12345
    """
    if not doc:
        script_logger.error('No active Revit document found.')
        return routes.make_response(
            data={'error': 'No active Revit document found.'},
            status=400
        )

    uidoc = UIDocument(doc)
    if not uidoc:
         script_logger.error('Could not get UIDocument from active document.')
         return routes.make_response(
            data={'error': 'Could not get UIDocument from active document.'},
            status=500
        )
        
    element_id_long = element_id_param
    if element_id_long <= 0:
        script_logger.error('Invalid Element ID provided: {}'.format(element_id_long))
        return routes.make_response(
             data={'error': 'Invalid Element ID provided in URL path: {}.'.format(element_id_long)},
             status=400
         )

    element_id = ElementId(element_id_long)

    try:
        element = doc.GetElement(element_id)
        if not element:
            script_logger.warning('Element with ID {} not found.'.format(element_id_long))
            return routes.make_response(
                data={'status': 'failure', 'message': 'Element with ID {} not found.'.format(element_id_long)},
                status=404
            )

        # --- Step 1: Find or Create Switchback View ---
        switchback_view = None
        try:
            collector = FilteredElementCollector(doc).OfClass(View3D)

            for view in collector:
                if not view.IsTemplate and view.Name == SWITCHBACK_VIEW_NAME:
                    switchback_view = view
                    break

            if not switchback_view:
                t_create = Transaction(doc, "Create Switchback View")
                try:
                    t_create.Start()
                    script_logger.debug("Creating new view: {}".format(SWITCHBACK_VIEW_NAME))
                    vft_collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
                    vft_3d = next((vft for vft in vft_collector if vft.ViewFamily == ViewFamily.ThreeDimensional), None)

                    if not vft_3d:
                        script_logger.error("Could not find a 3D ViewFamilyType.")
                        t_create.RollBack()
                        return routes.make_response({'error': 'Could not find a 3D ViewFamilyType.'}, status=500)

                    switchback_view = View3D.CreateIsometric(doc, vft_3d.Id)
                    switchback_view.Name = SWITCHBACK_VIEW_NAME

                    t_create.Commit()
                except Exception as create_ex:
                    script_logger.error("Error during view creation transaction: {}".format(create_ex))
                    if t_create.HasStarted(): t_create.RollBack()
                    raise
            else:
                 script_logger.debug("Found existing view: {}".format(SWITCHBACK_VIEW_NAME))

            if not switchback_view:
                 script_logger.error("Failed to find or create switchback view after process.")
                 return routes.make_response({'error': 'Failed to find or create switchback view.'}, status=500)

        except Exception as find_create_ex:
             script_logger.error("Error finding/creating view: {}".format(find_create_ex))
             return routes.make_response({'error': 'Error finding/creating view: {}'.format(find_create_ex)}, status=500)

        # --- Step 2: Regenerate, Apply Section Box & Select ---
        t_modify = Transaction(doc, "Regenerate, Apply Section Box & Select")
        try:
            t_modify.Start()

            doc.Regenerate()

            # 2a. Get Bounding Box
            bbox = element.get_BoundingBox(switchback_view)
            if not bbox:
                script_logger.warning("Element {} has no bounding box...".format(element_id_long))
            else:
                # 2b. Apply Section Box
                padding = 3
                min_pt = bbox.Min - XYZ(padding, padding, padding)
                max_pt = bbox.Max + XYZ(padding, padding, padding)
                padded_bbox = BoundingBoxXYZ()
                padded_bbox.Min = min_pt
                padded_bbox.Max = max_pt

                switchback_view.IsSectionBoxActive = True
                switchback_view.SetSectionBox(padded_bbox)

            # 2c. Select Element
            selection = uidoc.Selection
            selection.SetElementIds(List[ElementId]([element_id]))

            t_modify.Commit()

        except Exception as modify_ex:
            script_logger.error("Error during modification transaction: {}".format(modify_ex))
            if t_modify.HasStarted(): t_modify.RollBack()
            return routes.make_response(
                data={'error': 'An error occurred during section box/selection update: {}'.format(modify_ex)},
                status=500
            )
            
        # --- Step 3: Activate View ---
        try:
            uidoc.ActiveView = switchback_view
        except Exception as activate_ex:
             script_logger.error("Error activating view: {}".format(activate_ex))
             return routes.make_response({'error': 'Error activating view: {}'.format(activate_ex)}, status=500)

        # --- Step 4: Zoom to Element ---
        try:
            uidoc.ShowElements(element_id) # Operates on the now active view
        except Exception as zoom_ex:
            script_logger.warning("Failed to zoom to element: {}".format(zoom_ex))

        # Success
        return routes.make_response(
            data={'status': 'success', 'selected_id': element_id_long}
        )

    except Exception as e:
        # Catch errors before transactions start or other unexpected issues
        script_logger.error("Error selecting element ID {}: {}".format(element_id_long, e))
        return routes.make_response(
            data={'error': 'An unexpected error occurred: {}'.format(str(e))},
            status=500
        ) 