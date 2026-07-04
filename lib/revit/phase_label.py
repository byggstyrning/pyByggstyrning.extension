# -*- coding: utf-8 -*-
"""Zoom-anchored phase label for 3D views via TemporaryGraphicsManager.

Shows the active view's Phase name as a text badge pinned to the top-left
corner of the viewport. The InCanvasControl is anchored at a model-space
point, so the driver recomputes that point from UIView.GetZoomCorners()
whenever the user pans/zooms (Idling-driven, same pattern as
revit.view_markers). Temporary graphics only: nothing is written to the
model and the label never prints.
"""

import hashlib
import os
import sys
import time

import clr

clr.AddReference("System.Drawing")
from System.Drawing import (
    Bitmap, Graphics, Font, FontStyle, SolidBrush, Color, Pen, PointF,
)
from System.Drawing.Imaging import ImageFormat, PixelFormat
from System.Drawing.Drawing2D import SmoothingMode
from System.Drawing.Text import TextRenderingHint

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    InCanvasControlData,
    View3D,
)

from revit.compat import get_element_id_value
from revit.view_markers import (
    is_temporary_graphics_available,
    get_temporary_graphics_manager,
    ensure_temporary_graphics_handler,
    _normalize_image_path,
    _normalize_view_id,
)

_DRIVERS_SYS_KEY = '_pyBS_phase_label_drivers'
_THROTTLE_MS = 300
_CACHE_VERSION = 'v1'
_CACHE_SUBDIR = 'pyBS_phase_labels'
_MAX_LABEL_CHARS = 40
_MARGIN_PX = 14
_FONT_PT = 11.0
_PAD_X = 10
_PAD_Y = 6
# Revit clears exact RGB(0,128,128) as transparent; keep it off the palette.
_LABEL_BG = Color.FromArgb(255, 48, 50, 56)
_LABEL_BORDER = Color.FromArgb(255, 110, 110, 118)
_LABEL_TEXT = Color.FromArgb(255, 245, 245, 245)


def _cache_dir():
    localappdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    path = os.path.join(localappdata, 'pyBS', _CACHE_SUBDIR)
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except Exception:
            pass
    return path


def _label_font():
    for family in ('Segoe UI', 'Tahoma', 'Arial'):
        try:
            return Font(family, _FONT_PT, FontStyle.Bold)
        except Exception:
            continue
    return Font('Arial', _FONT_PT, FontStyle.Bold)


def _label_text_for_view(document, view):
    """Phase name of the view, or None when the view has no phase."""
    try:
        param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
        if param is None:
            return None
        phase_id = param.AsElementId()
        if phase_id is None or phase_id == ElementId.InvalidElementId:
            return None
        phase = document.GetElement(phase_id)
        if phase is None:
            return None
        name = phase.Name
        if not name:
            return None
        if len(name) > _MAX_LABEL_CHARS:
            name = name[:_MAX_LABEL_CHARS - 1] + u'…'
        return name
    except Exception:
        return None


def get_label_image_path(text):
    """Return absolute path to a cached BMP badge for this text.

    Also returns the (width, height) in pixels, needed to inset the anchor
    so the whole badge stays inside the viewport.
    """
    digest = hashlib.md5(
        (u'{}|{}|{}'.format(_CACHE_VERSION, _FONT_PT, text)).encode('utf-8')
    ).hexdigest()[:16]
    path = os.path.join(_cache_dir(), 'phase_{}.bmp'.format(digest))

    font = _label_font()
    probe = Bitmap(1, 1, PixelFormat.Format24bppRgb)
    g = Graphics.FromImage(probe)
    try:
        size = g.MeasureString(text, font)
        width = int(size.Width) + _PAD_X * 2
        height = int(size.Height) + _PAD_Y * 2
    finally:
        g.Dispose()
        probe.Dispose()

    if not os.path.isfile(path):
        bmp = Bitmap(width, height, PixelFormat.Format24bppRgb)
        g = Graphics.FromImage(bmp)
        try:
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit
            # Badge fills the whole bitmap: no chroma key, no edge fringes.
            bg = SolidBrush(_LABEL_BG)
            g.FillRectangle(bg, 0, 0, width, height)
            bg.Dispose()
            pen = Pen(_LABEL_BORDER, 1.0)
            g.DrawRectangle(pen, 0, 0, width - 1, height - 1)
            pen.Dispose()
            brush = SolidBrush(_LABEL_TEXT)
            g.DrawString(text, font, brush, PointF(float(_PAD_X), float(_PAD_Y)))
            brush.Dispose()
        finally:
            g.Dispose()
        bmp.Save(path, ImageFormat.Bmp)
        bmp.Dispose()
    font.Dispose()

    if not os.path.isfile(path) or os.path.getsize(path) < 32:
        raise IOError("Phase label bitmap missing: {}".format(path))
    return _normalize_image_path(path), (width, height)


def _view_basis(view):
    """Orthonormal screen basis (right, up, view direction) of a view."""
    return view.RightDirection, view.UpDirection, view.ViewDirection


def _top_left_anchor(view, uiview, label_px):
    """Model point where the badge center goes so it hugs the top-left corner.

    Decomposes the zoom-rectangle corners in the view's screen basis, takes
    (min right, max up) as the visible top-left, then insets by the margin
    plus half the badge size converted from pixels to model units via the
    UIView window width. InCanvasControl draws the image around its anchor,
    so a half-size inset keeps the badge fully inside the viewport.
    """
    corners = uiview.GetZoomCorners()
    if corners is None or len(corners) < 2:
        return None
    r, u, v = _view_basis(view)
    c1, c2 = corners[0], corners[1]
    r1, r2 = c1.DotProduct(r), c2.DotProduct(r)
    u1, u2 = c1.DotProduct(u), c2.DotProduct(u)
    depth = (c1.DotProduct(v) + c2.DotProduct(v)) / 2.0
    rmin, rmax = min(r1, r2), max(r1, r2)
    umax = max(u1, u2)

    rect = uiview.GetWindowRectangle()
    width_px = max(rect.Right - rect.Left, 1)
    units_per_px = (rmax - rmin) / float(width_px)
    label_w, label_h = label_px
    dx = (_MARGIN_PX + label_w / 2.0) * units_per_px
    dy = (_MARGIN_PX + label_h / 2.0) * units_per_px

    return (
        r.Multiply(rmin + dx)
        + u.Multiply(umax - dy)
        + v.Multiply(depth)
    )


def _corners_state_key(corners, rect, text):
    if corners is None or len(corners) < 2:
        return None
    a, b = corners[0], corners[1]
    return (
        round(a.X, 4), round(a.Y, 4), round(a.Z, 4),
        round(b.X, 4), round(b.Y, 4), round(b.Z, 4),
        rect.Right - rect.Left, rect.Bottom - rect.Top,
        text,
    )


def _get_uiview(uiapp, view_id):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    target = _normalize_view_id(get_element_id_value(view_id))
    for uiv in uidoc.GetOpenUIViews():
        try:
            if _normalize_view_id(get_element_id_value(uiv.ViewId)) == target:
                return uiv
        except Exception:
            continue
    return None


def collect_diagnostics(uiapp, document):
    """Step-by-step status report for troubleshooting a blank badge."""
    lines = []

    def add(key, fn):
        try:
            lines.append('{}: {}'.format(key, fn()))
        except Exception as ex:
            lines.append('{}: ERROR {}'.format(key, ex))

    add('engine', lambda: sys.version.replace('\n', ' '))
    add('tgm_available', is_temporary_graphics_available)
    add('tgm_manager', lambda: get_temporary_graphics_manager(document) is not None)

    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        lines.append('active_uidoc: None')
        return '\n'.join(lines)
    add('doc_match', lambda: uidoc.Document.Equals(document))
    view = uidoc.ActiveView
    if view is None:
        lines.append('active_view: None')
        return '\n'.join(lines)
    add('active_view', lambda: u'{} ({})'.format(view.Name, type(view).__name__))
    add('is_view3d', lambda: isinstance(view, View3D))
    add('is_template', lambda: view.IsTemplate)
    add('phase_param', lambda: view.get_Parameter(
        BuiltInParameter.VIEW_PHASE) is not None)
    text = _label_text_for_view(document, view)
    lines.append(u'phase_text: {}'.format(text))

    uiview = _get_uiview(uiapp, view.Id)
    lines.append('uiview_found: {}'.format(uiview is not None))
    if uiview is not None:
        add('window_rect', lambda: str(uiview.GetWindowRectangle()))
        add('zoom_corners', lambda: '; '.join(
            str(c) for c in uiview.GetZoomCorners()))

    if text:
        def _bitmap_info():
            path, px = get_label_image_path(text)
            return u'{} ({}x{}px, {} bytes)'.format(
                path, px[0], px[1], os.path.getsize(path))
        add('bitmap', _bitmap_info)
        if uiview is not None:
            add('anchor', lambda: str(_top_left_anchor(
                view, uiview, get_label_image_path(text)[1])))

    driver = find_phase_label_driver(document)
    lines.append('driver_running: {}'.format(driver is not None))
    if driver is not None:
        lines.append('control_index: {}'.format(driver._control_index))
        lines.append('control_view_id: {}'.format(driver._control_view_id))
        lines.append('last_error: {}'.format(driver._last_error))
    return '\n'.join(lines)


def find_phase_label_driver(document):
    drivers = getattr(sys, _DRIVERS_SYS_KEY, None)
    if not drivers:
        return None
    for driver in drivers:
        try:
            if driver._doc.Equals(document):
                return driver
        except Exception:
            continue
    return None


def stop_phase_label_driver(document):
    driver = find_phase_label_driver(document)
    if driver is not None:
        driver.stop()
        return True
    return False


def start_phase_label_driver(uiapp, document, logger=None):
    driver = find_phase_label_driver(document)
    if driver is not None:
        driver.refresh()
        return driver
    driver = PhaseLabelDriver(uiapp, document, logger=logger)
    driver.start()
    return driver


class PhaseLabelDriver(object):
    """Idling-driven phase badge pinned to the active 3D view's corner."""

    def __init__(self, uiapp, document, logger=None):
        self._uiapp = uiapp
        self._doc = document
        self._logger = logger
        self._control_index = None
        self._control_view_id = None
        self._state_key = None
        self._last_tick_ms = 0
        self._idling_handler = None
        self._view_activated_handler = None
        self._last_error = None

    def start(self):
        ensure_temporary_graphics_handler(self._doc, logger=self._logger)
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
        if not hasattr(sys, _DRIVERS_SYS_KEY):
            setattr(sys, _DRIVERS_SYS_KEY, [])
        getattr(sys, _DRIVERS_SYS_KEY).append(self)
        self.refresh()

    def stop(self):
        self._clear_control()
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
        try:
            getattr(sys, _DRIVERS_SYS_KEY).remove(self)
        except (AttributeError, ValueError):
            pass

    def refresh(self):
        self._state_key = None
        self._sync(force=True)

    def _clear_control(self):
        if self._control_index is None:
            return
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is not None:
            try:
                tgm.RemoveControl(self._control_index)
            except Exception:
                pass
        self._control_index = None
        self._control_view_id = None
        self._state_key = None

    def _on_view_activated(self, sender, args):
        try:
            self._sync(force=True)
        except Exception:
            pass

    def _on_idling(self, sender, args):
        now_ms = int(time.time() * 1000)
        if now_ms - self._last_tick_ms < _THROTTLE_MS:
            return
        self._last_tick_ms = now_ms
        try:
            self._sync()
        except Exception:
            pass

    def _sync(self, force=False):
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None or not uidoc.Document.Equals(self._doc):
            return
        view = uidoc.ActiveView
        if view is None or not isinstance(view, View3D) or view.IsTemplate:
            self._clear_control()
            return

        text = _label_text_for_view(self._doc, view)
        if not text:
            self._clear_control()
            return
        view_vid = _normalize_view_id(get_element_id_value(view.Id))

        uiview = _get_uiview(self._uiapp, view.Id)
        if uiview is None:
            return
        try:
            corners = uiview.GetZoomCorners()
            rect = uiview.GetWindowRectangle()
        except Exception:
            return
        key = _corners_state_key(corners, rect, text)
        if key is None:
            return
        if not force and key == self._state_key and \
                self._control_view_id == view_vid:
            return

        try:
            img_path, label_px = get_label_image_path(text)
        except Exception as ex:
            self._last_error = 'bitmap: {}'.format(ex)
            if self._logger is not None:
                self._logger.warning("Phase label bitmap failed: {}".format(ex))
            return
        anchor = _top_left_anchor(view, uiview, label_px)
        if anchor is None:
            return

        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            return
        ensure_temporary_graphics_handler(self._doc, logger=None)

        data = InCanvasControlData(img_path, anchor)
        updated = False
        if self._control_index is not None and \
                self._control_view_id == view_vid:
            try:
                tgm.UpdateControl(self._control_index, data)
                updated = True
            except Exception:
                updated = False
        if not updated:
            self._clear_control()
            try:
                self._control_index = tgm.AddControl(data, view.Id)
                self._control_view_id = view_vid
            except Exception as ex:
                self._last_error = 'AddControl: {}'.format(ex)
                self._control_index = None
                self._control_view_id = None
                return
        try:
            tgm.SetTooltip(
                self._control_index, "View phase: {}".format(text))
        except Exception:
            pass
        self._state_key = key
        try:
            uidoc.RefreshActiveView()
        except Exception:
            pass
