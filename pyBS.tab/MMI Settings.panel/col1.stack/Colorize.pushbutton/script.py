# -*- coding: utf-8 -*-
"""Toggles the MMI Colorer on and off.

Normal Click:
- When active, colors elements in the active view based on their MMI values.
- When inactive, removes the color overrides.

Shift+Click:
- Creates view filters for MMI ranges and applies them to the active view.
- This provides a more permanent, reusable solution compared to temporary overrides.
"""

__title__ = "Colorizer"
__author__ = "Byggstyrning AB"
__doc__ = "Toggle MMI-based coloring on/off. Shift+Click: Create view filters"
__highlight__ = 'new'

# Import standard libraries
import sys
import os
import os.path as op

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit
from pyrevit.userconfig import user_config
from pyrevit.coreutils.ribbon import ICON_MEDIUM
from pyrevit.revit import ui
import pyrevit.extensions as exts

# Add the extension directory to the path
script_path = __file__
pushbutton_dir = op.dirname(script_path)
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Try direct import from current directory's parent path
sys.path.append(op.dirname(op.dirname(panel_dir)))

# Initialize logger
logger = script.get_logger()

# Import MMI libraries - now using centralized colorizer module
from mmi import (
    get_mmi_parameter_name,
    get_element_mmi_value,
    MMI_COLOR_RANGES,
    get_color_for_mmi,
    is_colorer_active,
    set_colorer_active,
    get_colored_view_id,
    set_colored_view_id,
    get_colored_element_ids,
    set_colored_element_ids,
    clear_colorizer_state
)

# All MMI color ranges, state management, and helper functions
# are now imported from the mmi.colorizer library module

# Performance optimization: Cache for color lookups
_color_cache = {}
_solid_fill_pattern_cache = {}

def _get_solid_fill_pattern_id(doc):
    """Get solid fill pattern ID with caching."""
    if doc not in _solid_fill_pattern_cache:
        patterns = FilteredElementCollector(doc).OfClass(FillPatternElement)
        for pat in patterns:
            fill_pattern = pat.GetFillPattern()
            if fill_pattern.IsSolidFill:
                _solid_fill_pattern_cache[doc] = pat.Id
                break
        else:
            _solid_fill_pattern_cache[doc] = None
    return _solid_fill_pattern_cache[doc]

def _get_color_for_mmi_cached(mmi_value):
    """Get color for MMI value with caching."""
    if mmi_value not in _color_cache:
        _color_cache[mmi_value] = get_color_for_mmi(mmi_value)
    return _color_cache[mmi_value]

def apply_mmi_colors(doc, view):
    """Apply colors to elements in the view based on their MMI values.
    
    Returns:
        bool: Success status
    """
    try:
        # Get MMI parameter name
        mmi_param_name = get_mmi_parameter_name(doc)
        if not mmi_param_name:
            forms.alert("No MMI parameter configured. Please configure MMI settings first.", 
                       title="MMI Parameter Not Configured")
            return False
        
        logger.debug("Using MMI parameter: {}".format(mmi_param_name))
        
        # Get solid fill pattern (cached)
        solid_fill_id = _get_solid_fill_pattern_id(doc)
        
        # Get all elements in the view
        collector = FilteredElementCollector(doc, view.Id) \
                    .WhereElementIsNotElementType() \
                    .WhereElementIsViewIndependent()
        
        # Group elements by color using tuple key (faster than string concatenation)
        elements_by_color = {}
        colored_element_ids = []
        
        # Process elements
        for element in collector:
            # Skip elements that can't have overrides (early check)
            if not hasattr(element, 'Id'):
                continue
            
            # Get MMI value
            mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
            
            if mmi_value is not None:
                # Use cached color lookup
                color, range_name = _get_color_for_mmi_cached(mmi_value)
                
                # Group by color using tuple key (faster than string formatting)
                color_key = (color.Red, color.Green, color.Blue)
                if color_key not in elements_by_color:
                    elements_by_color[color_key] = {
                        "color": color,
                        "element_ids": []
                    }
                
                elements_by_color[color_key]["element_ids"].append(element.Id)
                colored_element_ids.append(element.Id)
        
        if not colored_element_ids:
            forms.alert("No elements with MMI values found in the active view.", 
                       title="No MMI Elements")
            return False
        
        # Apply color overrides in a transaction
        # Batch operations by creating override settings once per color group
        with Transaction(doc, "Apply MMI Colors") as t:
            t.Start()
            
            for color_data in elements_by_color.values():
                color = color_data["color"]
                element_ids = color_data["element_ids"]
                
                # Create override settings once per color group
                ogs = OverrideGraphicSettings()
                ogs.SetProjectionLineColor(color)
                ogs.SetCutLineColor(color)
                ogs.SetSurfaceForegroundPatternColor(color)
                ogs.SetCutForegroundPatternColor(color)
                
                if solid_fill_id:
                    ogs.SetSurfaceForegroundPatternId(solid_fill_id)
                    ogs.SetCutForegroundPatternId(solid_fill_id)
                
                # Apply overrides to all elements with this color
                # OverrideGraphicSettings is reused for all elements in this group
                for element_id in element_ids:
                    try:
                        view.SetElementOverrides(element_id, ogs)
                    except Exception as ex:
                        logger.debug("Could not apply override to element {}: {}".format(element_id, ex))
            
            t.Commit()
        
        # Store the colored elements and view
        set_colored_view_id(view.Id)
        set_colored_element_ids(colored_element_ids)
        
        return True
        
    except Exception as ex:
        logger.error("Error applying MMI colors: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
        return False

def reset_mmi_colors(doc):
    """Reset color overrides for previously colored elements.
    
    Returns:
        bool: Success status
    """
    try:
        # Get the view and elements that were colored
        view_id = get_colored_view_id()
        element_ids = get_colored_element_ids()
        
        if not view_id or not element_ids:
            logger.debug("No colored elements to reset")
            return True
        
        # Get the view
        view = doc.GetElement(view_id)
        if not view:
            logger.warning("Previously colored view not found")
            # Clear stored data anyway
            set_colored_view_id(ElementId.InvalidElementId)
            set_colored_element_ids([])
            return True
        
        # Reset overrides in a transaction
        with Transaction(doc, "Reset MMI Colors") as t:
            t.Start()
            
            # Create default override settings once (resets to defaults)
            ogs = OverrideGraphicSettings()
            
            # Reset all elements with same override settings
            for element_id in element_ids:
                try:
                    view.SetElementOverrides(element_id, ogs)
                except Exception as ex:
                    logger.debug("Could not reset override for element {}: {}".format(element_id, ex))
            
            t.Commit()
        
        # Clear stored data
        set_colored_view_id(ElementId.InvalidElementId)
        set_colored_element_ids([])
        
        return True
        
    except Exception as ex:
        logger.error("Error resetting MMI colors: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
        return False

# --- Button Initialization --- 

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    """Initialize the button icon based on the current active state."""
    try:
        # Use the same approach as Monitor script
        on_icon = ui.resolve_icon_file(script_cmp.directory, exts.DEFAULT_ON_ICON_FILE)
        off_icon = ui.resolve_icon_file(script_cmp.directory, exts.DEFAULT_OFF_ICON_FILE)

        button_icon = script_cmp.get_bundle_file(
            on_icon if is_colorer_active() else off_icon
        )
        ui_button_cmp.set_icon(button_icon, icon_size=ICON_MEDIUM)
    except Exception as e:
        logger.error("Error initializing MMI Colorer button: {}".format(e))

# --- Main Execution --- 

if __name__ == '__main__':
    was_active = is_colorer_active()
    new_active_state = not was_active
    
    doc = revit.doc
    active_view = doc.ActiveView
    
    # Check if we have a valid view
    if not active_view:
        forms.alert("No active view found.", title="Error")
    elif active_view.ViewType == ViewType.DrawingSheet:
        forms.alert("MMI Colorer cannot be used on sheets. Please open a model view.", 
                   title="Invalid View Type")
    else:
        success = False
        
        if new_active_state:
            # Activate: Apply colors
            logger.debug("Activating MMI Colorer...")
            
            success = apply_mmi_colors(doc, active_view)
            
            if success:
                set_colorer_active(True)
                script.toggle_icon(new_active_state)
        else:
            # Deactivate: Reset colors
            logger.debug("Deactivating MMI Colorer...")
            
            success = reset_mmi_colors(doc)
            
            if success:
                set_colorer_active(False)
                script.toggle_icon(new_active_state)
        
        if success:
            logger.debug("MMI Colorer state toggled to: {}".format("ON" if new_active_state else "OFF"))
        else:
            logger.error("Failed to toggle MMI Colorer state.")

