# -*- coding: utf-8 -*-
"""Clash-specific adapter for revit.view_markers (TemporaryGraphicsManager)."""

from revit.view_markers import (
    MarkerStyle,
    DEFAULT_MARKER_STYLE,
    ViewMarkerDriver,
    is_temporary_graphics_available,
    get_temporary_graphics_manager,
    ensure_temporary_graphics_handler,
    ensure_marker_bitmaps_ready,
    get_marker_image_path,
    register_session,
    get_session,
    is_session_toggle_active,
    set_session_toggle_active,
    find_marker_driver as _find_marker_driver,
    start_or_get_driver as _start_marker_driver,
    refresh_session,
    clean_session,
    cluster_points_model_space,
    _is_view_sheet,
    _sheet_clash_viewport_markers,
)

CLASH_SESSION_ID = 'clash'

CLASH_MARKER_STYLE = MarkerStyle(
    cache_subdir='pyBS_view_markers',
    cache_version='v8',
    marker_bmp_size=40,
    dot_diameter=14,
)


def register_marker_session(document, view_points_map, toggle_active=None):
    register_session(
        document, CLASH_SESSION_ID, view_points_map,
        toggle_active=toggle_active,
        marker_style=CLASH_MARKER_STYLE)


def get_marker_session(document):
    return get_session(document, CLASH_SESSION_ID)


def is_marker_toggle_active(document):
    return is_session_toggle_active(document, CLASH_SESSION_ID)


def set_marker_toggle_active(document, active):
    set_session_toggle_active(document, CLASH_SESSION_ID, active)


def find_marker_driver(document):
    return _find_marker_driver(document, CLASH_SESSION_ID)


def refresh_marker_session(document):
    return refresh_session(document, CLASH_SESSION_ID)


def clean_marker_session(document):
    return clean_session(document, CLASH_SESSION_ID)


def start_or_get_driver(uiapp, document, view_points_map,
                        get_element_id_value, logger=None):
    return _start_marker_driver(
        uiapp, document, view_points_map, get_element_id_value,
        session_id=CLASH_SESSION_ID,
        marker_style=CLASH_MARKER_STYLE,
        sheet_entry_builder=_sheet_clash_viewport_markers,
        sheet_tooltip='Double-click to view clashes',
        single_tooltip='Clash',
        cluster_tooltip_template='{count} clashes',
        logger=logger)


class ClashMarkerDriver(ViewMarkerDriver):
    """Clash views marker driver (session_id=clash, sheet viewport badges)."""

    def __init__(self, uiapp, document, view_points_map,
                 get_element_id_value, logger=None):
        ViewMarkerDriver.__init__(
            self, uiapp, document, view_points_map, get_element_id_value,
            session_id=CLASH_SESSION_ID,
            marker_style=CLASH_MARKER_STYLE,
            sheet_entry_builder=_sheet_clash_viewport_markers,
            sheet_tooltip='Double-click to view clashes',
            single_tooltip='Clash',
            cluster_tooltip_template='{count} clashes',
            logger=logger)
