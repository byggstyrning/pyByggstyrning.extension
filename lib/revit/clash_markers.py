# -*- coding: utf-8 -*-
"""Zoom-aware clash markers via TemporaryGraphicsManager (Revit 2022+)."""

import clr
import os
import time

clr.AddReference("System.Drawing")
from System.Drawing import (
    Bitmap, Graphics, Font, FontStyle, SolidBrush, Color, StringFormat,
    StringAlignment,
)
from System.Drawing.Imaging import ImageFormat

from Autodesk.Revit.DB import (
    TemporaryGraphicsManager,
    InCanvasControlData,
    XYZ,
)

from revit.compat import make_element_id

_MARKER_SYS_KEY = '_pyBS_clash_marker_drivers'
_TEMP_SUBDIR = 'pyBS_clash_markers'
_THROTTLE_MS = 250
_SPAN_FACTOR = 0.10
_MAX_MARKERS_PER_VIEW = 500
_TRANSPARENT_BG = Color.FromArgb(0, 128, 128)


def is_temporary_graphics_available():
    """Return True when TemporaryGraphicsManager API exists (Revit 2022+)."""
    try:
        return TemporaryGraphicsManager is not None
    except Exception:
        return False


def get_temporary_graphics_manager(document):
    try:
        return TemporaryGraphicsManager.GetTemporaryGraphicsManager(document)
    except Exception:
        return None


def _ensure_bitmap_cache_dir():
    temp = os.environ.get('TEMP', os.path.expanduser('~'))
    path = os.path.join(temp, _TEMP_SUBDIR)
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except Exception:
            pass
    return path


def get_marker_image_path(count):
    """Return absolute path to a cached BMP for single or cluster marker."""
    cache_dir = _ensure_bitmap_cache_dir()
    if count <= 1:
        key = 'single'
    elif count > 99:
        key = 'cluster_99plus'
    else:
        key = 'cluster_{}'.format(count)
    path = os.path.join(cache_dir, key + '.bmp')
    if os.path.isfile(path):
        return os.path.abspath(path)
    _generate_marker_bitmap(path, count)
    return os.path.abspath(path)


def _generate_marker_bitmap(path, count):
    size = 64
    bmp = Bitmap(size, size)
    g = Graphics.FromImage(bmp)
    g.Clear(_TRANSPARENT_BG)
    if count <= 1:
        brush = SolidBrush(Color.FromArgb(230, 50, 50))
        g.FillEllipse(brush, 8, 8, size - 16, size - 16)
        brush.Dispose()
    else:
        brush = SolidBrush(Color.FromArgb(40, 170, 70))
        g.FillEllipse(brush, 4, 4, size - 8, size - 8)
        brush.Dispose()
        label = str(count) if count <= 99 else '99+'
        font = Font("Arial", 14, FontStyle.Bold)
        text_brush = SolidBrush(Color.White)
        sf = StringFormat()
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        g.DrawString(label, font, text_brush, size / 2.0, size / 2.0, sf)
        font.Dispose()
        text_brush.Dispose()
        sf.Dispose()
    g.Dispose()
    bmp.Save(path, ImageFormat.Bmp)
    bmp.Dispose()


def cluster_points_model_space(xyz_list, cluster_radius):
    """Greedy model-space clustering. Returns list of dicts with xyz, count."""
    if not xyz_list:
        return []
    unassigned = list(range(len(xyz_list)))
    clusters = []
    while unassigned:
        seed = unassigned.pop(0)
        members = [seed]
        seed_pt = xyz_list[seed]
        new_unassigned = []
        for i in unassigned:
            if seed_pt.DistanceTo(xyz_list[i]) <= cluster_radius:
                members.append(i)
            else:
                new_unassigned.append(i)
        unassigned = new_unassigned
        cx = sum(xyz_list[m].X for m in members) / float(len(members))
        cy = sum(xyz_list[m].Y for m in members) / float(len(members))
        cz = sum(xyz_list[m].Z for m in members) / float(len(members))
        clusters.append({
            'xyz': XYZ(cx, cy, cz),
            'count': len(members),
        })
    return clusters


def _view_span_from_corners(corners):
    if corners is None or len(corners) < 2:
        return 1.0
    a = corners[0]
    b = corners[1]
    dx = abs(b.X - a.X)
    dy = abs(b.Y - a.Y)
    dz = abs(b.Z - a.Z)
    return max(dx, dy, dz, 1.0)


def _corners_key(corners):
    if corners is None or len(corners) < 2:
        return None
    a = corners[0]
    b = corners[1]
    return (
        round(a.X, 4), round(a.Y, 4), round(a.Z, 4),
        round(b.X, 4), round(b.Y, 4), round(b.Z, 4),
    )


def _get_uiview_for_view(uiapp, view_id_value, get_element_id_value):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    for uiv in uidoc.GetOpenUIViews():
        try:
            if get_element_id_value(uiv.ViewId) == view_id_value:
                return uiv
        except Exception:
            continue
    return None


def _points_from_dicts(point_dicts):
    out = []
    for p in point_dicts or []:
        try:
            out.append(XYZ(float(p['x']), float(p['y']), float(p['z'])))
        except Exception:
            continue
    return out


class ClashMarkerDriver(object):
    """Idling-driven temporary clash markers for created clash 3D views."""

    def __init__(self, uiapp, document, view_points_map, get_element_id_value,
                 logger=None):
        self._uiapp = uiapp
        self._doc = document
        self._sessions = dict(view_points_map or {})
        self._get_element_id_value = get_element_id_value
        self._logger = logger
        self._enabled = True
        self._control_indices = {}
        self._last_corners = {}
        self._last_refresh_ms = 0
        self._idling_handler = None
        self._view_activated_handler = None

    def set_sessions(self, view_points_map):
        self._sessions = dict(view_points_map or {})

    def set_enabled(self, enabled):
        self._enabled = bool(enabled)
        if not self._enabled:
            self._clear_all_markers()

    def start(self):
        if not self._sessions:
            return
        self._idling_handler = self._on_idling
        self._view_activated_handler = self._on_view_activated
        try:
            self._uiapp.Idling += self._idling_handler
        except Exception:
            pass
        try:
            self._uiapp.ViewActivated += self._view_activated_handler
        except Exception:
            pass
        import sys
        if not hasattr(sys, _MARKER_SYS_KEY):
            setattr(sys, _MARKER_SYS_KEY, [])
        getattr(sys, _MARKER_SYS_KEY).append(self)

    def stop(self):
        self._clear_all_markers()
        if self._idling_handler is not None:
            try:
                self._uiapp.Idling -= self._idling_handler
            except Exception:
                pass
            self._idling_handler = None
        if self._view_activated_handler is not None:
            try:
                self._uiapp.ViewActivated -= self._view_activated_handler
            except Exception:
                pass
            self._view_activated_handler = None
        import sys
        try:
            getattr(sys, _MARKER_SYS_KEY).remove(self)
        except (AttributeError, ValueError):
            pass

    def _log(self, message):
        if self._logger is not None:
            try:
                self._logger.debug(message)
            except Exception:
                pass

    def _clear_all_markers(self):
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            self._control_indices.clear()
            self._last_corners.clear()
            return
        for indices in self._control_indices.values():
            for idx in indices:
                try:
                    tgm.RemoveControl(idx)
                except Exception:
                    pass
        self._control_indices.clear()
        self._last_corners.clear()

    def _clear_view_markers(self, view_id_value):
        tgm = get_temporary_graphics_manager(self._doc)
        indices = self._control_indices.get(view_id_value, [])
        if tgm is not None:
            for idx in indices:
                try:
                    tgm.RemoveControl(idx)
                except Exception:
                    pass
        self._control_indices[view_id_value] = []
        self._last_corners.pop(view_id_value, None)

    def _on_view_activated(self, sender, args):
        if not self._enabled:
            return
        try:
            view = args.CurrentActiveView
            if view is None:
                return
            vid = self._get_element_id_value(view.Id)
            if vid in self._sessions:
                self._refresh_view_markers(vid, force=True)
            else:
                self._hide_non_active_markers(vid)
        except Exception as ex:
            self._log("ViewActivated marker refresh failed: {}".format(ex))

    def _on_idling(self, sender, args):
        if not self._enabled:
            return
        now_ms = int(time.time() * 1000)
        if now_ms - self._last_refresh_ms < _THROTTLE_MS:
            return
        self._last_refresh_ms = now_ms
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None:
            return
        try:
            active_view = uidoc.ActiveView
            if active_view is None:
                return
            active_vid = self._get_element_id_value(active_view.Id)
        except Exception:
            return
        self._hide_non_active_markers(active_vid)
        if active_vid not in self._sessions:
            return
        uiview = _get_uiview_for_view(
            self._uiapp, active_vid, self._get_element_id_value)
        if uiview is None:
            return
        try:
            corners = uiview.GetZoomCorners()
        except Exception:
            corners = None
        key = _corners_key(corners)
        if key is not None and self._last_corners.get(active_vid) == key:
            return
        self._last_corners[active_vid] = key
        self._refresh_view_markers(active_vid, corners=corners)

    def _hide_non_active_markers(self, active_vid):
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            return
        for vid, indices in self._control_indices.items():
            if vid == active_vid:
                continue
            for idx in indices:
                try:
                    tgm.SetVisibility(idx, False)
                except Exception:
                    pass

    def _refresh_view_markers(self, view_id_value, corners=None, force=False):
        point_dicts = self._sessions.get(view_id_value)
        if not point_dicts:
            self._clear_view_markers(view_id_value)
            return
        if len(point_dicts) > _MAX_MARKERS_PER_VIEW:
            self._log(
                "Clash markers capped at {} for view {}".format(
                    _MAX_MARKERS_PER_VIEW, view_id_value))
            point_dicts = point_dicts[:_MAX_MARKERS_PER_VIEW]

        uiview = _get_uiview_for_view(
            self._uiapp, view_id_value, self._get_element_id_value)
        if uiview is None:
            return
        if corners is None:
            try:
                corners = uiview.GetZoomCorners()
            except Exception:
                corners = None

        xyz_list = _points_from_dicts(point_dicts)
        if not xyz_list:
            self._clear_view_markers(view_id_value)
            return

        span = _view_span_from_corners(corners)
        radius = _SPAN_FACTOR * span
        clusters = cluster_points_model_space(xyz_list, radius)

        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            return

        owner_view_id = make_element_id(view_id_value)
        self._clear_view_markers(view_id_value)
        new_indices = []
        for cluster in clusters:
            count = cluster['count']
            xyz = cluster['xyz']
            try:
                img_path = get_marker_image_path(count)
                data = InCanvasControlData(img_path, xyz)
                ctrl_idx = tgm.AddControl(data, owner_view_id)
                new_indices.append(ctrl_idx)
                try:
                    if count > 1:
                        tgm.SetTooltip(ctrl_idx, "{} clashes".format(count))
                    else:
                        tgm.SetTooltip(ctrl_idx, "Clash")
                    tgm.SetVisibility(ctrl_idx, True)
                except Exception:
                    pass
            except Exception as ex:
                self._log("AddControl failed: {}".format(ex))
        self._control_indices[view_id_value] = new_indices
        if force and corners is not None:
            self._last_corners[view_id_value] = _corners_key(corners)
