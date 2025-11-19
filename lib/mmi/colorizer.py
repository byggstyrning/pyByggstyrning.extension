# -*- coding: utf-8 -*-
"""Configuration and utilities for MMI Colorizer."""

from Autodesk.Revit.DB import Color, ElementId
from pyrevit import script
from pyrevit.userconfig import user_config

# Initialize logger
logger = script.get_logger()

# Configuration section
CONFIG_SECTION = 'MMIColorizer'
CONFIG_KEY_ACTIVE = 'isActive'
CONFIG_KEY_VIEW_ID = 'coloredViewId'
CONFIG_KEY_ELEMENT_IDS = 'coloredElementIds'

# Official MMI Color codes from https://mmi-veilederen.no/?page_id=85
# RGB values as defined in the MMI guide
MMI_COLOR_RANGES = [
    {"value": 0, "color": Color(215, 50, 150), "name": "000 - Tidligfase"},
    {"value": 100, "color": Color(190, 40, 35), "name": "100 - Grunnlagsinformasjon"},
    {"value": 125, "color": Color(210, 75, 70), "name": "125 - Etablert konsept"},
    {"value": 150, "color": Color(225, 120, 115), "name": "150 - Tverrfaglig kontrollert konsept"},
    {"value": 175, "color": Color(240, 170, 170), "name": "175 - Valgt konsept"},
    {"value": 200, "color": Color(230, 150, 55), "name": "200 - Ferdig konsept"},
    {"value": 225, "color": Color(235, 175, 100), "name": "225 - Etablert prinsipielle løsninger"},
    {"value": 250, "color": Color(240, 200, 140), "name": "250 - Tverrfaglig kontrollert prinsipielle løsninger"},
    {"value": 275, "color": Color(245, 230, 215), "name": "275 - Valgt prinsipielle løsninger"},
    {"value": 300, "color": Color(250, 240, 80), "name": "300 - Underlag for detaljering"},
    {"value": 325, "color": Color(215, 205, 65), "name": "325 - Etablert detaljerte løsninger"},
    {"value": 350, "color": Color(185, 175, 60), "name": "350 - Tverrfaglig kontrollert detaljerte løsninger"},
    {"value": 375, "color": Color(150, 150, 50), "name": "375 - Detaljerte løsninger (anbud/bestilling)"},
    {"value": 400, "color": Color(55, 130, 70), "name": "400 - Arbeidsgrunnlag"},
    {"value": 425, "color": Color(75, 170, 90), "name": "425 - Etablert/utført"},
    {"value": 450, "color": Color(100, 195, 125), "name": "450 - Kontrollert utførelse"},
    {"value": 475, "color": Color(155, 215, 165), "name": "475 - Godkjent utförelse"},
    {"value": 500, "color": Color(30, 70, 175), "name": "500 - Som bygget"},
    {"value": 600, "color": Color(175, 50, 205), "name": "600 - I drift"},
]

def get_color_for_mmi(mmi_value):
    """Get the color for a given MMI value.
    
    Returns the color of the exact match or the closest lower MMI level.
    For example: 140 would use 125's color, 360 would use 350's color.
    
    Args:
        mmi_value: Numeric MMI value
        
    Returns:
        tuple: (Color, name) or (Gray, "Unknown") if no match
    """
    # Find the closest MMI level that is <= the value
    closest_range = None
    
    for range_def in MMI_COLOR_RANGES:
        if range_def["value"] <= mmi_value:
            closest_range = range_def
        else:
            # Since the list is sorted, we can break when we find a higher value
            break
    
    if closest_range:
        return closest_range["color"], closest_range["name"]
    
    # If no match found (value < 0), use the first defined color
    if MMI_COLOR_RANGES:
        return MMI_COLOR_RANGES[0]["color"], MMI_COLOR_RANGES[0]["name"]
    
    return Color(128, 128, 128), "Unknown"  # Gray fallback

def is_colorer_active():
    """Check if the MMI colorizer is currently active."""
    try:
        if not hasattr(user_config, CONFIG_SECTION):
            return False
        
        section = getattr(user_config, CONFIG_SECTION)
        return section.get_option(CONFIG_KEY_ACTIVE, default_value=False)
    except Exception as ex:
        logger.debug("Error checking colorizer state: {}".format(ex))
        return False

def set_colorer_active(is_active):
    """Set the MMI colorizer active state."""
    try:
        # Make sure the section exists
        if not hasattr(user_config, CONFIG_SECTION):
            user_config.add_section(CONFIG_SECTION)
        
        section = getattr(user_config, CONFIG_SECTION)
        section.set_option(CONFIG_KEY_ACTIVE, is_active)
        user_config.save_changes()
        
        logger.debug("MMI Colorizer active state set to: {}".format(is_active))
    except Exception as ex:
        logger.error("Error setting colorizer state: {}".format(ex))

def get_colored_view_id():
    """Get the ID of the view that was colored."""
    try:
        if not hasattr(user_config, CONFIG_SECTION):
            return None
        
        section = getattr(user_config, CONFIG_SECTION)
        view_id_str = section.get_option(CONFIG_KEY_VIEW_ID, default_value="")
        
        if view_id_str:
            try:
                return ElementId(int(view_id_str))
            except:
                return None
        return None
    except Exception as ex:
        logger.debug("Error getting colored view ID: {}".format(ex))
        return None

def set_colored_view_id(view_id):
    """Store the ID of the colored view."""
    try:
        if not hasattr(user_config, CONFIG_SECTION):
            user_config.add_section(CONFIG_SECTION)
        
        section = getattr(user_config, CONFIG_SECTION)
        section.set_option(CONFIG_KEY_VIEW_ID, str(view_id.IntegerValue))
        user_config.save_changes()
    except Exception as ex:
        logger.error("Error setting colored view ID: {}".format(ex))

def get_colored_element_ids():
    """Get the list of element IDs that were colored."""
    try:
        if not hasattr(user_config, CONFIG_SECTION):
            return []
        
        section = getattr(user_config, CONFIG_SECTION)
        ids_str = section.get_option(CONFIG_KEY_ELEMENT_IDS, default_value="")
        
        if ids_str:
            try:
                id_list = [ElementId(int(x)) for x in ids_str.split(",") if x.strip()]
                return id_list
            except Exception as ex:
                logger.debug("Error parsing colored element IDs: {}".format(ex))
                return []
        return []
    except Exception as ex:
        logger.debug("Error getting colored element IDs: {}".format(ex))
        return []

def set_colored_element_ids(element_ids):
    """Store the list of colored element IDs."""
    try:
        if not hasattr(user_config, CONFIG_SECTION):
            user_config.add_section(CONFIG_SECTION)
        
        section = getattr(user_config, CONFIG_SECTION)
        ids_str = ",".join([str(eid.IntegerValue) for eid in element_ids])
        section.set_option(CONFIG_KEY_ELEMENT_IDS, ids_str)
        user_config.save_changes()
    except Exception as ex:
        logger.error("Error setting colored element IDs: {}".format(ex))

def clear_colorizer_state():
    """Clear all colorizer state (view ID and element IDs)."""
    try:
        set_colored_view_id(ElementId.InvalidElementId)
        set_colored_element_ids([])
        logger.debug("Cleared colorizer state")
    except Exception as ex:
        logger.error("Error clearing colorizer state: {}".format(ex))


