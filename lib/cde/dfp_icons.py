# -*- coding: utf-8 -*-
"""DFP icon bitmaps — Hub Lucide glyphs + category hues (DoorMarkersOverlay parity)."""

from __future__ import division

import clr
import os

clr.AddReference("System.Drawing")
from System.Drawing import Bitmap, Graphics, Color, SolidBrush, Pen
from System.Drawing.Imaging import ImageFormat, PixelFormat
from System.Drawing.Drawing2D import SmoothingMode, InterpolationMode, DashStyle

from cde.dfp_catalog import DFP_PSET_NAME, RAW_FUNCTIONS, dfp_property_name

# Hub DoorMarkersOverlay layout
_CELL_PX = 20
_GAP_PX = 1
_PAD_PX = 1
_MAX_COLS = 4
_CACHE_VERSION = "v6"
_MAX_MARKER_PX = 128
_MAX_OVERLAY_PX = 256
_MIN_MARKER_PX = 28
_MARKER_CELL_PX = 24

# Revit TGM chroma key (see view_markers)
_CHROMA = Color.FromArgb(0, 128, 128)
# Neutral placeholder for doors with no authored DFP keys yet
_CLEAR_CELL_BG = Color.FromArgb(72, 76, 88)

_CODE_BY_KEY = {}
_LABEL_BY_CODE = {}
for entry in RAW_FUNCTIONS:
    if len(entry) >= 3:
        code, label, _cat = entry[0], entry[1], entry[2]
    else:
        code, label = entry[0], entry[1]
    _CODE_BY_KEY["{}.{}".format(DFP_PSET_NAME, dfp_property_name(code))] = code
    _LABEL_BY_CODE[code] = label


def is_dfp_value_key(key):
    return key.startswith(DFP_PSET_NAME + ".")


def code_from_value_key(key):
    return _CODE_BY_KEY.get(key)


def has_authored_dfp_value(value):
    """Match Hub ``hasAuthoredValue`` — skip unset / false / empty."""
    if value is None or value is False or value == "":
        return False
    try:
        if isinstance(value, (int, long)) and value == 0:
            return False
    except NameError:
        if isinstance(value, int) and value == 0:
            return False
    return True


def _assets_dir():
    return os.path.join(os.path.dirname(__file__), "assets", "dfp_cells")


def _cell_bitmap_path(code):
    safe = str(code).replace(".", "_")
    path = os.path.join(_assets_dir(), "func_{}.bmp".format(safe))
    if os.path.isfile(path):
        return path
    fallback = os.path.join(_assets_dir(), "func_fallback.bmp")
    if os.path.isfile(fallback):
        return fallback
    return path


def _render_clear_cell_bitmap(out_path):
    """Muted cell + dashed circle — visible 'no functions assigned' glyph."""
    bmp = Bitmap(_CELL_PX, _CELL_PX, PixelFormat.Format24bppRgb)
    g = Graphics.FromImage(bmp)
    g.SmoothingMode = SmoothingMode.AntiAlias
    g.Clear(_CLEAR_CELL_BG)
    try:
        pen = Pen(Color.White, 1.5)
        pen.DashStyle = DashStyle.Dash
        inset = 3
        g.DrawEllipse(
            pen,
            inset,
            inset,
            _CELL_PX - (inset * 2) - 1,
            _CELL_PX - (inset * 2) - 1,
        )
        pen.Dispose()
    finally:
        g.Dispose()
    parent = os.path.dirname(out_path)
    if not os.path.isdir(parent):
        try:
            os.makedirs(parent)
        except Exception:
            pass
    bmp.Save(out_path, ImageFormat.Bmp)
    bmp.Dispose()


def _clear_cell_bitmap_path():
    """Asset path for unset-door placeholder; generated on first use if missing."""
    path = os.path.join(_assets_dir(), "func_clear.bmp")
    if os.path.isfile(path):
        try:
            probe = Bitmap(path)
            try:
                if probe.Width == _CELL_PX and probe.Height == _CELL_PX:
                    return path
            finally:
                probe.Dispose()
        except Exception:
            pass
    _render_clear_cell_bitmap(path)
    return path


def _cell_cache_dir():
    localappdata = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    path = os.path.join(localappdata, "pyBS", "dfp_marker_cells", _CACHE_VERSION)
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except Exception:
            pass
    return path


def _composite_cache_dir():
    localappdata = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    path = os.path.join(localappdata, "pyBS", "dfp_marker_composites", _CACHE_VERSION)
    if not os.path.isdir(path):
        try:
            os.makedirs(path)
        except Exception:
            pass
    return path


def _normalize_image_path(path):
    path = os.path.abspath(path)
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(512)
        if ctypes.windll.kernel32.GetLongPathNameW(path, buf, 512):
            return buf.value
    except Exception:
        pass
    return path


def dfp_code_sort_key(code):
    """Stable sort for 4-column grid column order (Hub legend order)."""
    parts = str(code).split(".")
    out = []
    for part in parts:
        try:
            out.append(int(part))
        except Exception:
            out.append(part)
    return out


def _grid_geometry(slot_count, max_px=None):
    """Pixel layout for a 4-column table grid inside a square BMP."""
    cap = max_px if max_px is not None else _MAX_MARKER_PX
    cols = _MAX_COLS
    rows = int((int(slot_count) + cols - 1) / cols)
    grid_w = _PAD_PX * 2 + cols * _CELL_PX + (cols + 1) * _GAP_PX
    grid_h = _PAD_PX * 2 + rows * _CELL_PX + (rows + 1) * _GAP_PX
    size = max(grid_w, grid_h, _MIN_MARKER_PX)
    if size > cap:
        size = cap
    ox = (size - grid_w) // 2
    oy = (size - grid_h) // 2
    return {
        "bmp_px": size,
        "cols": cols,
        "rows": rows,
        "slot_count": int(slot_count),
        "grid_w": grid_w,
        "grid_h": grid_h,
        "ox": ox,
        "oy": oy,
        "cell_px": _CELL_PX,
        "gap_px": _GAP_PX,
        "pad_px": _PAD_PX,
    }


def composite_layout_dict(slot_count, max_px=None):
    """Public layout metadata for hit-testing (matches draw geometry)."""
    return dict(_grid_geometry(slot_count, max_px=max_px))


def _composite_state_cache_key(sorted_keys, active_flags):
    parts = []
    for key, active in zip(sorted_keys, active_flags):
        code = code_from_value_key(key) or key
        safe = str(code).replace(".", "_")
        parts.append("{}_{}".format(safe, "1" if active else "0"))
    return "_".join(parts)


def _cached_cell_valid(path):
    try:
        if not os.path.isfile(path) or os.path.getsize(path) < 32:
            return False
        probe = Bitmap(path)
        try:
            return (
                probe.Width == probe.Height
                and probe.Width == _MARKER_CELL_PX)
        finally:
            probe.Dispose()
    except Exception:
        return False


def _render_cell_bitmap(src_path, out_path, active=True, hover=False):
    """Square TGM cell; dim when inactive, highlight border when hover."""
    size = _MARKER_CELL_PX
    bmp = Bitmap(size, size, PixelFormat.Format24bppRgb)
    g = Graphics.FromImage(bmp)
    g.SmoothingMode = SmoothingMode.HighQuality
    g.InterpolationMode = InterpolationMode.HighQualityBicubic
    g.Clear(_CHROMA)
    try:
        if os.path.isfile(src_path):
            cell = Bitmap(src_path)
            try:
                g.DrawImage(cell, 0, 0, size, size)
            finally:
                cell.Dispose()
        if not active:
            overlay = SolidBrush(Color.FromArgb(150, 0, 0, 0))
            g.FillRectangle(overlay, 0, 0, size, size)
        if hover:
            pen = Pen(Color.FromArgb(255, 255, 220, 60), 2)
            g.DrawRectangle(pen, 1, 1, size - 3, size - 3)
            pen.Dispose()
    finally:
        g.Dispose()
    bmp.Save(out_path, ImageFormat.Bmp)
    bmp.Dispose()


def get_cell_marker_path(code, active=True, hover=False):
    """Return square BMP for one overlay button (on/off/hover)."""
    safe = str(code).replace(".", "_")
    if hover:
        state = "hover_{}".format("on" if active else "off")
    else:
        state = "on" if active else "off"
    out_path = os.path.join(_cell_cache_dir(), "cell_{}_{}.bmp".format(safe, state))
    if _cached_cell_valid(out_path):
        return _normalize_image_path(out_path)
    src = _cell_bitmap_path(code)
    _render_cell_bitmap(src, out_path, active=active, hover=hover)
    return _normalize_image_path(out_path)


def _summary_cache_key(codes):
    return "summary_" + "_".join(str(c).replace(".", "_") for c in codes)


def get_dfp_summary_composite(active_codes):
    """Compact 4×Y composite of **active/true** functions only (collapsed view)."""
    codes = sorted(set(str(c) for c in active_codes if c), key=dfp_code_sort_key)
    layout = _grid_geometry(max(len(codes), 1))
    cache_dir = _composite_cache_dir()
    if not codes:
        out_path = os.path.join(cache_dir, "summary_empty.bmp")
        if not _cached_composite_valid(out_path, layout["bmp_px"]):
            size = layout["bmp_px"]
            bmp = Bitmap(size, size, PixelFormat.Format24bppRgb)
            g = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.HighQuality
            g.InterpolationMode = InterpolationMode.HighQualityBicubic
            g.Clear(_CHROMA)
            try:
                cx = (size - _CELL_PX) // 2
                cy = (size - _CELL_PX) // 2
                cell_path = _clear_cell_bitmap_path()
                cell = Bitmap(cell_path)
                try:
                    g.DrawImage(cell, cx, cy, _CELL_PX, _CELL_PX)
                finally:
                    cell.Dispose()
            finally:
                g.Dispose()
            bmp.Save(out_path, ImageFormat.Bmp)
            bmp.Dispose()
        return _normalize_image_path(out_path), layout, []

    cache_key = _summary_cache_key(codes)
    out_path = os.path.join(cache_dir, cache_key + ".bmp")
    layout = _grid_geometry(len(codes))
    if _cached_composite_valid(out_path, layout["bmp_px"]):
        return _normalize_image_path(out_path), layout, list(codes)

    cols = layout["cols"]
    ox = layout["ox"]
    oy = layout["oy"]
    size = layout["bmp_px"]
    bmp = Bitmap(size, size, PixelFormat.Format24bppRgb)
    g = Graphics.FromImage(bmp)
    g.SmoothingMode = SmoothingMode.HighQuality
    g.InterpolationMode = InterpolationMode.HighQualityBicubic
    g.Clear(_CHROMA)
    try:
        for idx, code in enumerate(codes):
            row = idx // cols
            col = idx % cols
            x = ox + _PAD_PX + _GAP_PX + col * (_CELL_PX + _GAP_PX)
            y = oy + _PAD_PX + _GAP_PX + row * (_CELL_PX + _GAP_PX)
            _draw_cell(g, code, x, y, True)
    finally:
        g.Dispose()
    bmp.Save(out_path, ImageFormat.Bmp)
    bmp.Dispose()
    return _normalize_image_path(out_path), layout, list(codes)


def composite_slot_screen_center(layout, slot_index, screen_cx, screen_cy):
    """Screen coords of one full-grid slot center (overlay alignment)."""
    bmp_px = float(layout.get("bmp_px", 0))
    ox = layout.get("ox", 0)
    oy = layout.get("oy", 0)
    cell = layout.get("cell_px", _CELL_PX)
    gap = layout.get("gap_px", _GAP_PX)
    pad = layout.get("pad_px", _PAD_PX)
    cols = layout.get("cols", _MAX_COLS)
    row = int(slot_index) // cols
    col = int(slot_index) % cols
    lx = ox + pad + gap + col * (cell + gap) + (cell / 2.0)
    ly = oy + pad + gap + row * (cell + gap) + (cell / 2.0)
    return screen_cx - (bmp_px / 2.0) + lx, screen_cy - (bmp_px / 2.0) + ly


def _cached_composite_valid(path, expected_px):
    try:
        if not os.path.isfile(path) or os.path.getsize(path) < 32:
            return False
        probe = Bitmap(path)
        try:
            return (
                probe.Width == probe.Height
                and probe.Width == expected_px)
        finally:
            probe.Dispose()
    except Exception:
        return False


def _draw_cell(g, code, x, y, active):
    cell_path = _cell_bitmap_path(code)
    if not os.path.isfile(cell_path):
        return
    cell = Bitmap(cell_path)
    try:
        g.DrawImage(cell, x, y, _CELL_PX, _CELL_PX)
    finally:
        cell.Dispose()
    if not active:
        overlay = SolidBrush(Color.FromArgb(150, 0, 0, 0))
        g.FillRectangle(overlay, x, y, _CELL_PX, _CELL_PX)


def get_dfp_composite_path(sorted_keys, active_flags):
    """Build/cache one square 4×Y composite for visible DFP columns on a door.

    ``sorted_keys`` — column keys in grid order.
    ``active_flags`` — bool per key (bright vs dimmed cell).
    Returns ``(path, layout_dict, slot_keys)``.
    """
    if not sorted_keys:
        raise ValueError("sorted_keys required")
    if len(active_flags) != len(sorted_keys):
        raise ValueError("active_flags length mismatch")

    layout = _grid_geometry(len(sorted_keys))
    cache_dir = _composite_cache_dir()
    cache_key = _composite_state_cache_key(sorted_keys, active_flags)
    out_path = os.path.join(cache_dir, cache_key + ".bmp")
    if _cached_composite_valid(out_path, layout["bmp_px"]):
        return _normalize_image_path(out_path), layout, list(sorted_keys)

    cols = layout["cols"]
    ox = layout["ox"]
    oy = layout["oy"]
    size = layout["bmp_px"]

    bmp = Bitmap(size, size, PixelFormat.Format24bppRgb)
    g = Graphics.FromImage(bmp)
    g.SmoothingMode = SmoothingMode.HighQuality
    g.InterpolationMode = InterpolationMode.HighQualityBicubic
    g.Clear(_CHROMA)
    try:
        for idx, key in enumerate(sorted_keys):
            code = code_from_value_key(key)
            if not code:
                continue
            row = idx // cols
            col = idx % cols
            x = ox + _PAD_PX + _GAP_PX + col * (_CELL_PX + _GAP_PX)
            y = oy + _PAD_PX + _GAP_PX + row * (_CELL_PX + _GAP_PX)
            _draw_cell(g, code, x, y, active_flags[idx])
    finally:
        g.Dispose()
    bmp.Save(out_path, ImageFormat.Bmp)
    bmp.Dispose()
    return _normalize_image_path(out_path), layout, list(sorted_keys)


def hit_test_composite_slot(pixel_x, pixel_y, layout):
    """Map local composite pixel coords to slot index, or None."""
    if layout is None:
        return None
    try:
        px = float(pixel_x)
        py = float(pixel_y)
    except Exception:
        return None
    bmp_px = layout.get("bmp_px", 0)
    if px < 0 or py < 0 or px >= bmp_px or py >= bmp_px:
        return None
    ox = layout.get("ox", 0)
    oy = layout.get("oy", 0)
    cell = layout.get("cell_px", _CELL_PX)
    gap = layout.get("gap_px", _GAP_PX)
    pad = layout.get("pad_px", _PAD_PX)
    cols = layout.get("cols", _MAX_COLS)
    slot_count = layout.get("slot_count", 0)
    lx = px - ox - pad - gap
    ly = py - oy - pad - gap
    if lx < 0 or ly < 0:
        return None
    step = cell + gap
    col = int(lx / step)
    row = int(ly / step)
    if col < 0 or col >= cols or row < 0:
        return None
    if lx - col * step >= cell or ly - row * step >= cell:
        return None
    idx = row * cols + col
    if idx < 0 or idx >= slot_count:
        return None
    return idx


def build_marker_tooltip(codes):
    parts = []
    for code in sorted(codes, key=dfp_code_sort_key):
        label = _LABEL_BY_CODE.get(code, code)
        parts.append(u"{} {}".format(code, label))
    return u"\n".join(parts)
