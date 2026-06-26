# -*- coding: utf-8 -*-
"""Reusable zoom-aware view markers via TemporaryGraphicsManager (Revit 2022+).

Point markers in 3D views with optional clustering; sheet mode shows one badge
per registered viewport. Use MarkerStyle + session_id for multiple tools.
"""

import clr
import os
import time

clr.AddReference("System.Drawing")
from System.Drawing import (
    Bitmap, Graphics, Font, FontStyle, SolidBrush, Color, StringFormat,
    StringAlignment, Pen,
)
from System.Drawing.Imaging import ImageFormat, PixelFormat
from System.Drawing.Drawing2D import SmoothingMode, InterpolationMode, CompositingQuality
from System.Drawing.Text import TextRenderingHint

from Autodesk.Revit.DB import (
    TemporaryGraphicsManager,
    InCanvasControlData,
    XYZ,
    ViewSheet,
    ElementId,
)

from revit.compat import make_element_id

_MARKER_SYS_KEY = '_pyBS_view_marker_drivers'
_HANDLER_SYS_KEY = '_pyBS_view_marker_tgm_handler'
_CLICK_REGISTRY_KEY = '_pyBS_tgm_click_registry'
_SESSION_CLICK_HANDLERS_KEY = '_pyBS_tgm_session_click_handlers'
_HANDLER_SERVER_GUID_STR = 'B3C4D5E6-F7A8-4901-BC02-EF1234567890'
_THROTTLE_MS = 500
_VIEW_REFRESH_MS = 1000
_BADGE_BORDER_COLOR = Color.FromArgb(255, 48, 50, 56)
_SHADOW_COLOR = Color.FromArgb(255, 110, 110, 118)
_TEXT_COLOR = Color.FromArgb(255, 28, 30, 36)


class MarkerStyle(object):
    """Bitmap and clustering settings for one marker session / tool."""

    def __init__(
        self,
        cache_subdir='pyBS_view_markers',
        cache_version='v8',
        marker_bmp_size=40,
        dot_diameter=14,
        span_factor=0.10,
        max_markers_per_view=500,
        badge_margin=3,
        shadow_offset=2,
    ):
        self.cache_subdir = cache_subdir
        self.cache_version = cache_version
        self.marker_bmp_size = marker_bmp_size
        self.dot_diameter = dot_diameter
        self.span_factor = span_factor
        self.max_markers_per_view = max_markers_per_view
        self.badge_margin = badge_margin
        self.shadow_offset = shadow_offset


DEFAULT_MARKER_STYLE = MarkerStyle()
# Revit clears exact RGB(0,128,128) as transparent (see InCanvasControlData remarks).
_CHROMA_R = 0
_CHROMA_G = 128
_CHROMA_B = 128
_TRANSPARENT_BG = Color.FromArgb(_CHROMA_R, _CHROMA_G, _CHROMA_B)


def is_temporary_graphics_available():
    """Return True when TemporaryGraphicsManager API exists (Revit 2022+)."""
    try:
        return (
            TemporaryGraphicsManager is not None
            and hasattr(TemporaryGraphicsManager, 'GetTemporaryGraphicsManager')
        )
    except Exception:
        return False


def _get_tgm_handler_service_id():
    """Resolve TemporaryGraphicsHandlerService across Revit API layouts."""
    try:
        from Autodesk.Revit.DB.ExternalService import ExternalServices
        return (
            ExternalServices.BuiltInExternalServices
            .TemporaryGraphicsHandlerService
        )
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB.ExternalService import BuiltInExternalServices
        return BuiltInExternalServices.TemporaryGraphicsHandlerService
    except Exception:
        pass
    return None


def _create_tgm_handler_instance():
    """Build ITemporaryGraphicsHandler server (lazy; needs RevitAPIUI)."""
    service_id = _get_tgm_handler_service_id()
    if service_id is None:
        return None
    try:
        clr.AddReference('RevitAPIUI')
        from Autodesk.Revit.UI import ITemporaryGraphicsHandler
        from System import Guid

        server_guid = Guid(_HANDLER_SERVER_GUID_STR)

        class ViewMarkerTemporaryGraphicsHandler(ITemporaryGraphicsHandler):
            """No-op click handler for TGM in-canvas controls."""

            def GetName(self):
                return "pyBS View Markers"

            def GetDescription(self):
                return "pyBS temporary view marker graphics handler"

            def GetServerId(self):
                return server_guid

            def GetServiceId(self):
                return service_id

            def GetVendorId(self):
                return "pyBS"

            def OnClick(self, data):
                _dispatch_tgm_click(data)

        return ViewMarkerTemporaryGraphicsHandler()
    except Exception:
        return None


def get_temporary_graphics_manager(document):
    try:
        return TemporaryGraphicsManager.GetTemporaryGraphicsManager(document)
    except Exception:
        return None


def _click_registry():
    import sys
    if not hasattr(sys, _CLICK_REGISTRY_KEY):
        setattr(sys, _CLICK_REGISTRY_KEY, {})
    return getattr(sys, _CLICK_REGISTRY_KEY)


def clear_control_clicks(document):
    """Drop click payloads for a document (before marker redraw)."""
    key = _document_key(document)
    reg = _click_registry()
    if key in reg:
        reg[key] = {}


def register_control_click(document, control_index, payload):
    """Map a TGM control index to click metadata for OnClick dispatch."""
    key = _document_key(document)
    reg = _click_registry()
    doc_reg = reg.get(key) or {}
    doc_reg[int(control_index)] = dict(payload or {})
    reg[key] = doc_reg


def lookup_control_click(document, control_index):
    reg = _click_registry().get(_document_key(document)) or {}
    return reg.get(int(control_index))


def unregister_control_click(document, control_index):
    """Remove one TGM control from the click registry."""
    key = _document_key(document)
    reg = _click_registry()
    doc_reg = reg.get(key)
    if not doc_reg:
        return
    doc_reg.pop(int(control_index), None)
    reg[key] = doc_reg


def register_session_click_handler(session_id, handler):
    """Register ``handler(payload, command_data)`` for a marker session."""
    import sys
    if not hasattr(sys, _SESSION_CLICK_HANDLERS_KEY):
        setattr(sys, _SESSION_CLICK_HANDLERS_KEY, {})
    handlers = getattr(sys, _SESSION_CLICK_HANDLERS_KEY)
    handlers[str(session_id)] = handler


def unregister_session_click_handler(session_id):
    import sys
    if not hasattr(sys, _SESSION_CLICK_HANDLERS_KEY):
        return
    handlers = getattr(sys, _SESSION_CLICK_HANDLERS_KEY)
    handlers.pop(str(session_id), None)


def _dispatch_tgm_click(command_data):
    try:
        document = command_data.Document
        idx = command_data.Index
        payload = lookup_control_click(document, idx)
        if not payload:
            return
        session_id = payload.get('session_id') or DEFAULT_SESSION_ID
        import sys
        handlers = getattr(sys, _SESSION_CLICK_HANDLERS_KEY, {})
        handler = handlers.get(str(session_id))
        if handler is not None:
            handler(payload, command_data)
    except Exception:
        pass


def _ensure_bitmap_cache_dir(marker_style=None):
    """Stable cache under LOCALAPPDATA (avoids pyRevit isolated TEMP + short paths)."""
    style = marker_style or DEFAULT_MARKER_STYLE
    localappdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    path = os.path.join(localappdata, 'pyBS', style.cache_subdir)
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except Exception:
            pass
    return path


def _normalize_image_path(path):
    """Absolute long path for Revit image loader (no 8.3 short names)."""
    path = os.path.abspath(path)
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(512)
        if ctypes.windll.kernel32.GetLongPathNameW(path, buf, 512):
            return buf.value
    except Exception:
        pass
    return path


def _bitmap_cache_key(count, style_key, marker_style=None):
    style = marker_style or DEFAULT_MARKER_STYLE
    ver = style.cache_version
    if style_key == 'dot' or count <= 1:
        return 'dot_' + ver
    if count > 99:
        return 'cluster_99plus_' + ver
    return 'cluster_{}_{}'.format(count, ver)


def _bitmap_size_for_key(key, marker_style=None):
    style = marker_style or DEFAULT_MARKER_STYLE
    return style.marker_bmp_size


def _invalidate_stale_bitmap_cache(marker_style=None):
    """Remove legacy BMPs (wrong size/version) so Revit reloads fresh assets."""
    style = marker_style or DEFAULT_MARKER_STYLE
    cache_dir = _ensure_bitmap_cache_dir(style)
    try:
        for name in os.listdir(cache_dir):
            if not name.endswith('.bmp'):
                continue
            path = os.path.join(cache_dir, name)
            if style.cache_version in name:
                key = name[:-4]
                if _cached_bitmap_valid(path, _bitmap_size_for_key(key, style), style):
                    continue
            try:
                os.remove(path)
            except Exception:
                pass
    except Exception:
        pass


def _cached_bitmap_valid(path, expected_size, marker_style=None):
    style = marker_style or DEFAULT_MARKER_STYLE
    try:
        probe = Bitmap(path)
        ok = (
            probe.Width == expected_size
            and probe.Height == expected_size
            and probe.Width == style.marker_bmp_size)
        probe.Dispose()
        return ok
    except Exception:
        return False


def get_marker_image_path(count, style_key='dot', marker_style=None):
    """Return absolute path to cached BMP (dot or cluster badge)."""
    style = marker_style or DEFAULT_MARKER_STYLE
    cache_dir = _ensure_bitmap_cache_dir(style)
    key = _bitmap_cache_key(count, style_key, style)
    expected_size = _bitmap_size_for_key(key, style)
    path = os.path.join(cache_dir, key + '.bmp')
    if os.path.isfile(path) and not _cached_bitmap_valid(path, expected_size, style):
        try:
            os.remove(path)
        except Exception:
            pass
    if not os.path.isfile(path):
        _generate_marker_bitmap(path, count, style_key, style)
    if not os.path.isfile(path):
        raise IOError("Marker bitmap missing after generation: {}".format(path))
    if os.path.getsize(path) < 32:
        raise IOError("Marker bitmap too small: {}".format(path))
    return _normalize_image_path(path)


def ensure_marker_bitmaps_ready(marker_style=None):
    """Pre-create BMP marker images (TGM requires bitmap, not PNG)."""
    style = marker_style or DEFAULT_MARKER_STYLE
    _invalidate_stale_bitmap_cache(style)
    errors = []
    paths = []
    for count in (2, 5, 99):
        try:
            paths.append(get_marker_image_path(count, 'cluster', style))
        except Exception as ex:
            errors.append(str(ex))
    try:
        paths.append(get_marker_image_path(1, 'dot', style))
    except Exception as ex:
        errors.append(str(ex))
    return paths, errors


def _badge_diameter(size, marker_style=None):
    style = marker_style or DEFAULT_MARKER_STYLE
    return size - style.badge_margin * 2 - style.shadow_offset


def _make_badge_font(size_pt):
    for family in ('Segoe UI', 'Tahoma', 'Arial'):
        try:
            return Font(family, size_pt, FontStyle.Bold)
        except Exception:
            continue
    return Font('Arial', size_pt, FontStyle.Bold)


def _begin_marker_graphics(bmp):
    g = Graphics.FromImage(bmp)
    g.SmoothingMode = SmoothingMode.AntiAlias
    g.InterpolationMode = InterpolationMode.HighQualityBicubic
    g.CompositingQuality = CompositingQuality.HighQuality
    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit
    g.Clear(_TRANSPARENT_BG)
    return g


def _draw_circle_badge(g, size, count=None, marker_style=None):
    """White circle with drop shadow, border, and optional count label."""
    style = marker_style or DEFAULT_MARKER_STYLE
    x = float(style.badge_margin)
    y = float(style.badge_margin)
    d = float(_badge_diameter(size, style))
    sh = style.shadow_offset

    shadow_brush = SolidBrush(_SHADOW_COLOR)
    g.FillEllipse(
        shadow_brush,
        x + sh, y + sh, d, d)
    shadow_brush.Dispose()

    fill_brush = SolidBrush(Color.FromArgb(255, 255, 255, 255))
    g.FillEllipse(fill_brush, x, y, d, d)
    fill_brush.Dispose()

    border_pen = Pen(_BADGE_BORDER_COLOR, 1.0)
    g.DrawEllipse(border_pen, x + 0.5, y + 0.5, d - 1.0, d - 1.0)
    border_pen.Dispose()

    if count is not None and count > 1:
        label = str(count) if count <= 99 else '99+'
        font_size = 10.0 if count <= 9 else 8.0
        if count > 99:
            font_size = 7.0
        font = _make_badge_font(font_size)
        text_brush = SolidBrush(_TEXT_COLOR)
        sf = StringFormat()
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center
        g.DrawString(label, font, text_brush, x + d / 2.0, y + d / 2.0, sf)
        font.Dispose()
        text_brush.Dispose()
        sf.Dispose()


def _draw_dot_marker(g, size, diameter):
    """Filled dot with light border (no icon)."""
    d = float(diameter)
    x = (float(size) - d) / 2.0
    y = (float(size) - d) / 2.0
    fill_brush = SolidBrush(_TEXT_COLOR)
    g.FillEllipse(fill_brush, x, y, d, d)
    fill_brush.Dispose()
    border_pen = Pen(Color.FromArgb(255, 255, 255, 255), 1.5)
    g.DrawEllipse(border_pen, x, y, d, d)
    border_pen.Dispose()


def _generate_marker_bitmap(path, count, style_key, marker_style=None):
    style = marker_style or DEFAULT_MARKER_STYLE
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        try:
            os.makedirs(parent)
        except Exception:
            pass
    size = style.marker_bmp_size
    bmp = Bitmap(size, size, PixelFormat.Format24bppRgb)
    g = _begin_marker_graphics(bmp)
    try:
        if style_key == 'dot' or count <= 1:
            _draw_dot_marker(g, size, style.dot_diameter)
        else:
            _draw_circle_badge(g, size, count=count, marker_style=style)
    finally:
        g.Dispose()
    bmp.Save(path, ImageFormat.Bmp)
    bmp.Dispose()
    if not os.path.isfile(path) or os.path.getsize(path) < 32:
        raise IOError("Failed to write marker bitmap: {}".format(path))
    if not _cached_bitmap_valid(path, size, style):
        raise IOError(
            "Marker bitmap wrong dimensions after write: {}".format(path))


def ensure_temporary_graphics_handler(document, logger=None):
    """Register and activate ITemporaryGraphicsHandler for this document."""
    import sys

    def _warn(msg):
        if logger is not None:
            try:
                logger.warning(msg)
            except Exception:
                pass

    try:
        from Autodesk.Revit.DB.ExternalService import ExternalServiceRegistry
        from System import Guid
        from System.Collections.Generic import List

        service_id = _get_tgm_handler_service_id()
        if service_id is None:
            _warn("TemporaryGraphicsHandler: service id not found")
            return False

        ext_service = ExternalServiceRegistry.GetService(service_id)
        if ext_service is None:
            _warn("TemporaryGraphicsHandler: GetService returned None")
            return False

        handler = getattr(sys, _HANDLER_SYS_KEY, None)
        if handler is None:
            handler = _create_tgm_handler_instance()
            if handler is None:
                _warn("TemporaryGraphicsHandler: could not create handler")
                return False
            setattr(sys, _HANDLER_SYS_KEY, handler)

        server_guid = Guid(_HANDLER_SERVER_GUID_STR)
        try:
            ext_service.AddServer(handler)
        except Exception as add_ex:
            add_msg = str(add_ex)
            if 'already' not in add_msg.lower() and 'registered' not in add_msg.lower():
                _warn(
                    "TemporaryGraphicsHandler AddServer failed: {}".format(
                        add_msg))

        id_list = List[Guid]()
        has_ours = False
        try:
            active_ids = ext_service.GetActiveServerIds()
            for gid in active_ids:
                id_list.Add(gid)
                if gid == server_guid:
                    has_ours = True
        except Exception as ex:
            _warn(
                "TemporaryGraphicsHandler GetActiveServerIds: {}".format(ex))
        if not has_ours:
            id_list.Add(server_guid)

        activated = False
        last_err = ''

        # App-level activation (required — see Revit forum TemporaryGraphicsManager)
        try:
            ext_service.SetActiveServers(id_list)
            activated = True
        except Exception as ex:
            last_err = "app: {}".format(ex)

        # Optional per-document activation
        if document is not None:
            try:
                doc_ids = List[Guid]()
                doc_has_ours = False
                try:
                    for gid in ext_service.GetActiveServerIds(document):
                        doc_ids.Add(gid)
                        if gid == server_guid:
                            doc_has_ours = True
                except Exception:
                    pass
                if not doc_has_ours:
                    doc_ids.Add(server_guid)
                ext_service.SetActiveServers(doc_ids, document)
            except Exception as ex:
                if last_err:
                    last_err += "; "
                last_err += "doc: {}".format(ex)

        if activated:
            return True

        _warn("TemporaryGraphicsHandler activation failed: {}".format(last_err))
        return False
    except Exception as ex:
        _warn("TemporaryGraphicsHandler registration failed: {}".format(ex))
        return False


def _points_centroid(xyz_list):
    if not xyz_list:
        return None
    n = float(len(xyz_list))
    return XYZ(
        sum(p.X for p in xyz_list) / n,
        sum(p.Y for p in xyz_list) / n,
        sum(p.Z for p in xyz_list) / n,
    )


def _marker_anchor_for_clash_view(document, view_id_value, xyz_list):
    """Model anchor above clash section box (reads well on sheet viewports)."""
    try:
        view = document.GetElement(make_element_id(view_id_value))
        if view is not None:
            sb = view.GetSectionBox()
            if sb is not None:
                cx = (sb.Min.X + sb.Max.X) / 2.0
                cy = (sb.Min.Y + sb.Max.Y) / 2.0
                dz = sb.Max.Z - sb.Min.Z
                lift = max(dz * 0.08, 0.5)
                return XYZ(cx, cy, sb.Max.Z + lift)
    except Exception:
        pass
    return _points_centroid(xyz_list)


def _sheet_viewport_marker_entries(sheet, document, get_element_id_value,
                                   sessions, anchor_for_view=None):
    """One grouped marker per viewport whose view id is in sessions."""
    session_map = sessions or {}
    session_set = set(_normalize_view_id(v) for v in session_map.keys())
    entries = []
    try:
        for vp_id in sheet.GetAllViewports():
            vp = document.GetElement(vp_id)
            if vp is None:
                continue
            view_vid = _normalize_view_id(get_element_id_value(vp.ViewId))
            if view_vid not in session_set:
                continue
            point_dicts = session_map.get(view_vid) or []
            xyz_list = _points_from_dicts(point_dicts)
            if not xyz_list:
                continue
            if anchor_for_view is not None:
                xyz = anchor_for_view(document, view_vid, xyz_list)
            else:
                xyz = _points_centroid(xyz_list)
            if xyz is None:
                continue
            entries.append({
                'view_id': view_vid,
                'count': len(xyz_list),
                'xyz': xyz,
            })
    except Exception:
        pass
    return entries


def _clash_sheet_anchor(document, view_id_value, xyz_list):
    """Clash views: anchor above section box on sheet viewports."""
    return _marker_anchor_for_clash_view(document, view_id_value, xyz_list)


def _sheet_clash_viewport_markers(sheet, document, get_element_id_value,
                                  sessions):
    return _sheet_viewport_marker_entries(
        sheet, document, get_element_id_value, sessions,
        anchor_for_view=_clash_sheet_anchor)


def _owner_view_id_for_control(view_id_value):
    """TGM owner: markers scoped to a single view (never all views)."""
    return make_element_id(view_id_value)


def cluster_points_model_space(xyz_list, cluster_radius):
    """Greedy model-space clustering. Marker at seed clash point, not centroid."""
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
        clusters.append({
            'xyz': seed_pt,
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


def _view_span_from_view3d(document, view_id_value):
    """Cluster radius fallback when UIView zoom corners are unavailable."""
    try:
        view = document.GetElement(make_element_id(view_id_value))
        if view is None:
            return 1.0
        try:
            sb = view.GetSectionBox()
            if sb is not None:
                dx = abs(sb.Max.X - sb.Min.X)
                dy = abs(sb.Max.Y - sb.Min.Y)
                dz = abs(sb.Max.Z - sb.Min.Z)
                return max(dx, dy, dz, 1.0)
        except Exception:
            pass
        try:
            cb = view.CropBox
            if cb is not None:
                dx = abs(cb.Max.X - cb.Min.X)
                dy = abs(cb.Max.Y - cb.Min.Y)
                dz = abs(cb.Max.Z - cb.Min.Z)
                return max(dx, dy, dz, 1.0)
        except Exception:
            pass
    except Exception:
        pass
    return 1.0


def _is_view_sheet(view):
    try:
        return isinstance(view, ViewSheet)
    except Exception:
        return False


def _clash_view_ids_on_sheet(sheet, document, get_element_id_value, session_ids):
    """Return clash view id values placed on a sheet."""
    session_set = set(_normalize_view_id(v) for v in session_ids)
    found = []
    try:
        for vp_id in sheet.GetAllViewports():
            vp = document.GetElement(vp_id)
            if vp is None:
                continue
            vid = _normalize_view_id(get_element_id_value(vp.ViewId))
            if vid in session_set:
                found.append(vid)
    except Exception:
        pass
    return found


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
    target_vid = _normalize_view_id(view_id_value)
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    for uiv in uidoc.GetOpenUIViews():
        try:
            if _normalize_view_id(get_element_id_value(uiv.ViewId)) == target_vid:
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


def _validate_marker_image(path):
    """Return True when BMP exists and is square (TGM requirement)."""
    try:
        if not path or not os.path.isfile(path):
            return False
        if os.path.getsize(path) < 32:
            return False
        probe = Bitmap(path)
        try:
            return probe.Width == probe.Height and probe.Width >= 16
        finally:
            probe.Dispose()
    except Exception:
        return False


def update_control_image(document, control_index, image_path, xyz):
    """Swap one TGM control bitmap in place (after toggle)."""
    if not _validate_marker_image(image_path):
        return False
    tgm = get_temporary_graphics_manager(document)
    if tgm is None:
        return False
    try:
        data = InCanvasControlData(image_path, xyz)
        tgm.UpdateControl(int(control_index), data)
        return True
    except Exception:
        return False


def _normalize_view_id(view_id_value):
    """Normalize ElementId values (int/long/Int64) for consistent dict keys."""
    try:
        return int(view_id_value)
    except Exception:
        try:
            return int(str(view_id_value))
        except Exception:
            return view_id_value


def _normalize_view_points_map(view_points_map):
    out = {}
    for vid, points in (view_points_map or {}).items():
        out[_normalize_view_id(vid)] = points
    return out


_SESSION_REGISTRY_KEY = '_pyBS_view_marker_registry'
DEFAULT_SESSION_ID = 'default'


def _get_session_registry():
    import sys
    if not hasattr(sys, _SESSION_REGISTRY_KEY):
        setattr(sys, _SESSION_REGISTRY_KEY, {})
    return getattr(sys, _SESSION_REGISTRY_KEY)


def _document_key(document):
    try:
        path = document.PathName
        if path:
            return path
    except Exception:
        pass
    try:
        title = document.Title
        if title:
            return title
    except Exception:
        pass
    return 'unknown'


def register_session(document, session_id, view_points_map,
                     toggle_active=None, marker_style=None):
    """Store per-document marker data for a named session (tool / feature)."""
    key = _document_key(document)
    reg = _get_session_registry()
    doc_entry = reg.get(key) or {}
    entry = doc_entry.get(session_id) or {
        'view_points': {},
        'toggle_active': False,
        'marker_style': None,
    }
    entry['view_points'] = _normalize_view_points_map(view_points_map)
    if toggle_active is not None:
        entry['toggle_active'] = bool(toggle_active)
    if marker_style is not None:
        entry['marker_style'] = marker_style
    doc_entry[session_id] = entry
    reg[key] = doc_entry


def get_session(document, session_id=DEFAULT_SESSION_ID):
    key = _document_key(document)
    doc_entry = _get_session_registry().get(key) or {}
    entry = doc_entry.get(session_id)
    if not entry:
        return {}
    return _normalize_view_points_map(entry.get('view_points') or {})


def is_session_toggle_active(document, session_id=DEFAULT_SESSION_ID):
    key = _document_key(document)
    doc_entry = _get_session_registry().get(key) or {}
    entry = doc_entry.get(session_id)
    if not entry:
        return False
    return bool(entry.get('toggle_active', False))


def set_session_toggle_active(document, session_id, active):
    key = _document_key(document)
    reg = _get_session_registry()
    doc_entry = reg.get(key) or {}
    entry = doc_entry.get(session_id) or {
        'view_points': {},
        'toggle_active': False,
        'marker_style': None,
    }
    entry['toggle_active'] = bool(active)
    doc_entry[session_id] = entry
    reg[key] = doc_entry


def _session_marker_style(document, session_id):
    key = _document_key(document)
    doc_entry = _get_session_registry().get(key) or {}
    entry = doc_entry.get(session_id) or {}
    style = entry.get('marker_style')
    if style is not None:
        return style
    return DEFAULT_MARKER_STYLE


def find_marker_driver(document, session_id=DEFAULT_SESSION_ID):
    key = _document_key(document)
    import sys
    try:
        drivers = getattr(sys, _MARKER_SYS_KEY)
    except AttributeError:
        return None
    for driver in drivers:
        try:
            if (_document_key(driver._doc) == key
                    and driver._session_id == session_id):
                return driver
        except Exception:
            continue
    return None


def start_or_get_driver(uiapp, document, view_points_map, get_element_id_value,
                        session_id=DEFAULT_SESSION_ID,
                        marker_style=None,
                        sheet_entry_builder=None,
                        sheet_tooltip='',
                        single_tooltip='Marker',
                        cluster_tooltip_template='{count} markers',
                        logger=None,
                        refresh_active_view=True):
    """Return running driver for document session, creating one if needed."""
    driver = find_marker_driver(document, session_id)
    if driver is not None:
        if view_points_map:
            driver.set_sessions(view_points_map)
        return driver
    if not view_points_map:
        return None
    style = marker_style or _session_marker_style(document, session_id)
    driver = ViewMarkerDriver(
        uiapp, document, view_points_map, get_element_id_value,
        session_id=session_id,
        marker_style=style,
        sheet_entry_builder=sheet_entry_builder,
        sheet_tooltip=sheet_tooltip,
        single_tooltip=single_tooltip,
        cluster_tooltip_template=cluster_tooltip_template,
        logger=logger,
        refresh_active_view=refresh_active_view)
    driver.start()
    return driver


def refresh_session(document, session_id=DEFAULT_SESSION_ID):
    driver = find_marker_driver(document, session_id)
    if driver is None:
        return False
    driver.refresh_all()
    return True


def clean_session(document, session_id=DEFAULT_SESSION_ID):
    """Remove marker graphics, stop driver, mark session toggle inactive."""
    driver = find_marker_driver(document, session_id)
    if driver is not None:
        driver.stop()
    set_session_toggle_active(document, session_id, False)
    return True


class ViewMarkerDriver(object):
    """Idling-driven temporary markers for registered 3D views."""

    def __init__(self, uiapp, document, view_points_map, get_element_id_value,
                 session_id=DEFAULT_SESSION_ID,
                 marker_style=None,
                 sheet_entry_builder=None,
                 sheet_tooltip='',
                 single_tooltip='Marker',
                 cluster_tooltip_template='{count} markers',
                 logger=None,
                 refresh_active_view=True):
        self._uiapp = uiapp
        self._doc = document
        self._session_id = session_id
        self._style = marker_style or DEFAULT_MARKER_STYLE
        self._sheet_entry_builder = sheet_entry_builder
        self._sheet_tooltip = sheet_tooltip or ''
        self._single_tooltip = single_tooltip or 'Marker'
        self._cluster_tooltip_template = cluster_tooltip_template or '{count}'
        self._refresh_active_view = bool(refresh_active_view)
        self._sessions = _normalize_view_points_map(view_points_map or {})
        self._get_element_id_value = get_element_id_value
        self._logger = logger
        self._enabled = True
        self._control_indices = {}
        self._sheet_control_indices = {}
        self._last_corners = {}
        self._last_refresh_ms = 0
        self._last_view_refresh_ms = {}
        self._last_sheet_refresh_ms = {}
        self._refresh_in_progress = False
        self._idling_handler = None
        self._view_activated_handler = None

    def set_sessions(self, view_points_map):
        self._sessions = _normalize_view_points_map(view_points_map or {})

    def set_enabled(self, enabled):
        self._enabled = bool(enabled)
        if not self._enabled:
            self._clear_all_markers()

    def refresh_all(self):
        """Redraw markers for the current active view context only."""
        if not self._enabled:
            return
        self._last_view_refresh_ms.clear()
        self._last_corners.clear()
        self._last_sheet_refresh_ms.clear()
        self._sync_markers_to_active_context(force=True)

    def _sync_markers_to_active_context(self, force=False):
        """Sheet: one badge per clash viewport. 3D clash view: markers in that view only."""
        try:
            uidoc = self._uiapp.ActiveUIDocument
            if uidoc is None or uidoc.ActiveView is None:
                return
            active_view = uidoc.ActiveView
            if _is_view_sheet(active_view):
                self._clear_clash_view_markers_only()
                self._refresh_sheet_markers(active_view, force=force)
                return
            active_vid = _normalize_view_id(
                self._get_element_id_value(active_view.Id))
            if active_vid in self._sessions:
                self._clear_sheet_markers()
                self._clear_clash_view_markers_except(active_vid)
                uiview = _get_uiview_for_view(
                    self._uiapp, active_vid, self._get_element_id_value)
                corners = None
                if uiview is not None:
                    try:
                        corners = uiview.GetZoomCorners()
                    except Exception:
                        corners = None
                self._refresh_view_markers(
                    active_vid, corners=corners, force=force)
                self._maybe_refresh_active_view_for_session()
            else:
                self._clear_all_markers()
        except Exception:
            pass

    def _maybe_refresh_active_view_for_session(self):
        """Redraw viewport when active view is a session clash view."""
        if not self._refresh_active_view:
            return
        try:
            uidoc = self._uiapp.ActiveUIDocument
            if uidoc is None or uidoc.ActiveView is None:
                return
            if _is_view_sheet(uidoc.ActiveView):
                return
            active_vid = _normalize_view_id(
                self._get_element_id_value(uidoc.ActiveView.Id))
            if active_vid not in self._sessions:
                return
            uidoc.RefreshActiveView()
        except Exception:
            pass

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
        _, bmp_errors = ensure_marker_bitmaps_ready(self._style)
        if bmp_errors and self._logger is not None:
            try:
                self._logger.warning(
                    "Marker bitmap errors: {}".format(bmp_errors[:2]))
            except Exception:
                pass
        if not ensure_temporary_graphics_handler(self._doc, logger=self._logger):
            if self._logger is not None:
                try:
                    self._logger.warning(
                        "TemporaryGraphicsHandler not active (markers may not display)")
                except Exception:
                    pass
        self._trigger_initial_refresh()

    def _trigger_initial_refresh(self):
        """Refresh markers for whatever view is currently active."""
        self._sync_markers_to_active_context(force=True)

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

    def _clear_all_markers(self):
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            self._control_indices.clear()
            self._sheet_control_indices.clear()
            self._last_corners.clear()
            return
        for indices in self._control_indices.values():
            for idx in indices:
                try:
                    tgm.RemoveControl(idx)
                except Exception:
                    pass
        for indices in self._sheet_control_indices.values():
            for idx in indices:
                try:
                    tgm.RemoveControl(idx)
                except Exception:
                    pass
        self._control_indices.clear()
        self._sheet_control_indices.clear()
        self._last_corners.clear()

    def _clear_clash_view_markers_only(self):
        for vid in list(self._sessions.keys()):
            self._clear_view_markers(vid)

    def _clear_clash_view_markers_except(self, keep_view_id_value):
        keep = _normalize_view_id(keep_view_id_value)
        for vid in list(self._sessions.keys()):
            if _normalize_view_id(vid) != keep:
                self._clear_view_markers(vid)

    def _clear_sheet_markers(self, sheet_id_value=None):
        tgm = get_temporary_graphics_manager(self._doc)
        if sheet_id_value is not None:
            keys = [_normalize_view_id(sheet_id_value)]
        else:
            keys = list(self._sheet_control_indices.keys())
        for sid in keys:
            indices = self._sheet_control_indices.get(sid, [])
            if tgm is not None:
                for idx in indices:
                    try:
                        tgm.RemoveControl(idx)
                    except Exception:
                        pass
            self._sheet_control_indices.pop(sid, None)

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
            self._sync_markers_to_active_context(force=True)
        except Exception:
            pass

    def _refresh_sheet_markers(self, sheet, force=False):
        sheet_vid = _normalize_view_id(self._get_element_id_value(sheet.Id))
        now_ms = int(time.time() * 1000)
        last_ms = self._last_sheet_refresh_ms.get(sheet_vid, 0)
        if not force and now_ms - last_ms < _VIEW_REFRESH_MS:
            return

        ensure_temporary_graphics_handler(self._doc, logger=None)

        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            return

        if self._sheet_entry_builder is not None:
            entries = self._sheet_entry_builder(
                sheet, self._doc, self._get_element_id_value, self._sessions)
        else:
            entries = _sheet_viewport_marker_entries(
                sheet, self._doc, self._get_element_id_value, self._sessions)
        self._clear_sheet_markers(sheet_vid)
        owner_view_id = _owner_view_id_for_control(sheet_vid)
        new_indices = []
        for entry in entries:
            count = entry['count']
            xyz = entry['xyz']
            try:
                img_path = get_marker_image_path(
                    count, 'cluster', self._style)
                data = InCanvasControlData(img_path, xyz)
                ctrl_idx = tgm.AddControl(data, owner_view_id)
                new_indices.append(ctrl_idx)
                try:
                    tgm.SetVisibility(ctrl_idx, True)
                except Exception:
                    pass
                try:
                    if self._sheet_tooltip:
                        tgm.SetTooltip(ctrl_idx, self._sheet_tooltip)
                except Exception:
                    pass
            except Exception:
                pass

        self._sheet_control_indices[sheet_vid] = new_indices
        self._last_sheet_refresh_ms[sheet_vid] = now_ms

    def _refresh_clash_views_on_sheet(self, sheet, force=False):
        self._refresh_sheet_markers(sheet, force=force)

    def _on_idling(self, sender, args):
        if not self._enabled or self._refresh_in_progress:
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
            if _is_view_sheet(active_view):
                self._refresh_sheet_markers(active_view, force=False)
                return
            active_vid = _normalize_view_id(
                self._get_element_id_value(active_view.Id))
        except Exception:
            return

        if active_vid not in self._sessions:
            return

        uiview = _get_uiview_for_view(
            self._uiapp, active_vid, self._get_element_id_value)
        corners = None
        if uiview is not None:
            try:
                corners = uiview.GetZoomCorners()
            except Exception:
                corners = None
            key = _corners_key(corners)
            if key is not None and self._last_corners.get(active_vid) == key:
                return
            self._last_corners[active_vid] = key
        self._clear_sheet_markers()
        self._clear_clash_view_markers_except(active_vid)
        self._refresh_view_markers(active_vid, corners=corners)

    def _refresh_view_markers(self, view_id_value, corners=None, force=False):
        view_id_value = _normalize_view_id(view_id_value)
        if self._refresh_in_progress:
            return
        now_ms = int(time.time() * 1000)
        last_ms = self._last_view_refresh_ms.get(view_id_value, 0)
        if not force and now_ms - last_ms < _VIEW_REFRESH_MS:
            return

        point_dicts = self._sessions.get(view_id_value)
        if not point_dicts:
            self._clear_view_markers(view_id_value)
            return
        if len(point_dicts) > self._style.max_markers_per_view:
            point_dicts = point_dicts[:self._style.max_markers_per_view]

        custom_dicts = [
            p for p in point_dicts if p.get('image_path')]
        plain_dicts = [
            p for p in point_dicts if not p.get('image_path')]

        uiview = _get_uiview_for_view(
            self._uiapp, view_id_value, self._get_element_id_value)
        if corners is None and uiview is not None:
            try:
                corners = uiview.GetZoomCorners()
            except Exception:
                corners = None

        ensure_temporary_graphics_handler(self._doc, logger=None)
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            return

        clear_control_clicks(self._doc)
        self._refresh_in_progress = True
        try:
            owner_view_id = _owner_view_id_for_control(view_id_value)
            self._clear_view_markers(view_id_value)
            new_indices = []

            for p in custom_dicts:
                try:
                    xyz = XYZ(float(p['x']), float(p['y']), float(p['z']))
                    img_path = p.get('image_path')
                    if not img_path or not _validate_marker_image(img_path):
                        continue
                    data = InCanvasControlData(img_path, xyz)
                    ctrl_idx = tgm.AddControl(data, owner_view_id)
                    new_indices.append(ctrl_idx)
                    if p.get('clickable'):
                        payload = dict(p)
                        payload['session_id'] = (
                            p.get('session_id') or self._session_id)
                        payload['control_index'] = ctrl_idx
                        register_control_click(self._doc, ctrl_idx, payload)
                    try:
                        tgm.SetVisibility(ctrl_idx, True)
                    except Exception:
                        pass
                    tip = p.get('tooltip') or self._single_tooltip
                    try:
                        tgm.SetTooltip(ctrl_idx, tip)
                    except Exception:
                        pass
                except Exception:
                    pass

            xyz_list = _points_from_dicts(plain_dicts)
            if xyz_list:
                if corners is not None:
                    span = _view_span_from_corners(corners)
                else:
                    span = _view_span_from_view3d(self._doc, view_id_value)
                radius = self._style.span_factor * span
                clusters = cluster_points_model_space(xyz_list, radius)
                for cluster in clusters:
                    count = cluster['count']
                    xyz = cluster['xyz']
                    try:
                        sk = 'cluster' if count > 1 else 'dot'
                        img_path = get_marker_image_path(count, sk, self._style)
                        data = InCanvasControlData(img_path, xyz)
                        ctrl_idx = tgm.AddControl(data, owner_view_id)
                        new_indices.append(ctrl_idx)
                        try:
                            tgm.SetVisibility(ctrl_idx, True)
                        except Exception:
                            pass
                        try:
                            if count > 1:
                                tgm.SetTooltip(
                                    ctrl_idx,
                                    self._cluster_tooltip_template.format(
                                        count=count))
                            else:
                                tgm.SetTooltip(ctrl_idx, self._single_tooltip)
                        except Exception:
                            pass
                    except Exception:
                        pass

            self._control_indices[view_id_value] = new_indices
            self._last_view_refresh_ms[view_id_value] = now_ms
            if force and corners is not None:
                self._last_corners[view_id_value] = _corners_key(corners)
            self._maybe_refresh_active_view_for_session()
        finally:
            self._refresh_in_progress = False
