# -*- coding: utf-8 -*-
"""DFP door-function markers via TemporaryGraphicsManager (Hub icon parity)."""

import time

from Autodesk.Revit.DB import XYZ, InCanvasControlData

from revit.view_markers import (
    MarkerStyle,
    ViewMarkerDriver,
    is_temporary_graphics_available,
    ensure_temporary_graphics_handler,
    register_session,
    get_session,
    is_session_toggle_active,
    set_session_toggle_active,
    find_marker_driver as _find_marker_driver,
    refresh_session,
    clean_session,
    register_session_click_handler,
    unregister_session_click_handler,
    update_control_image,
    register_control_click,
    unregister_control_click,
    clear_control_clicks,
    cluster_points_model_space,
    get_marker_image_path,
    get_temporary_graphics_manager,
    _get_uiview_for_view,
    _view_span_from_corners,
    _view_span_from_view3d,
    _validate_marker_image,
    _owner_view_id_for_control,
    _normalize_view_id,
    _VIEW_REFRESH_MS,
)
from cde.dfp_markers import DFP_SESSION_ID
from cde.dfp_icons import (
    code_from_value_key,
    get_dfp_summary_composite,
    get_cell_marker_path,
    composite_slot_screen_center,
    hit_test_composite_slot,
    _LABEL_BY_CODE,
)

DFP_MARKER_STYLE = MarkerStyle(
    cache_subdir="pyBS_dfp_markers",
    cache_version="v5",
    marker_bmp_size=48,
    dot_diameter=14,
    span_factor=0.09,
    max_markers_per_view=500,
)

_DFP_CLICK_CB_KEY = "_pyBS_dfp_marker_click_cb"
_DFP_UIAPP_KEY = "_pyBS_dfp_marker_uiapp"
_DFP_SCHEDULE_WINDOW_KEY = "_pyBS_dfp_schedule_window"
_DFP_APPLY_BLOCK_KEY = "_pyBS_dfp_apply_block"


def set_dfp_apply_block(blocked):
    """Immediate UI-thread flag — block idling/TGM while CDE Apply runs."""
    import sys
    setattr(sys, _DFP_APPLY_BLOCK_KEY, bool(blocked))


def is_dfp_apply_blocked():
    import sys
    return bool(getattr(sys, _DFP_APPLY_BLOCK_KEY, False))


_HOVER_THROTTLE_MS = 60


class _ClickShim(object):
    """Minimal command_data stand-in (Revit may dispose the original)."""

    def __init__(self, index, document):
        self.Index = index
        self.Document = document


def register_dfp_session(document, view_points_map, toggle_active=None):
    register_session(
        document, DFP_SESSION_ID, view_points_map,
        toggle_active=toggle_active,
        marker_style=DFP_MARKER_STYLE)


def get_dfp_session(document):
    return get_session(document, DFP_SESSION_ID)


def is_dfp_toggle_active(document):
    return is_session_toggle_active(document, DFP_SESSION_ID)


def set_dfp_toggle_active(document, active):
    set_session_toggle_active(document, DFP_SESSION_ID, active)


def find_dfp_driver(document):
    return _find_marker_driver(document, DFP_SESSION_ID)


def hide_dfp_graphics_for_apply(document):
    """Hide all DFP TGM controls without RemoveControl (Apply quarantine)."""
    driver = find_dfp_driver(document)
    if driver is None:
        return 0
    return driver.hide_all_graphics_visibility()


def refresh_dfp_session(document):
    return refresh_session(document, DFP_SESSION_ID)


def soft_clean_dfp_session(document):
    """Disable DFP markers without TGM RemoveControl (Revit 2026 ArrowEditor)."""
    unregister_session_click_handler(DFP_SESSION_ID)
    driver = find_dfp_driver(document)
    hidden = 0
    if driver is not None:
        hidden = driver.soft_stop()
    set_dfp_toggle_active(document, False)
    return True


def clean_dfp_session(document):
    unregister_session_click_handler(DFP_SESSION_ID)
    import sys
    if hasattr(sys, _DFP_UIAPP_KEY):
        delattr(sys, _DFP_UIAPP_KEY)
    return clean_session(document, DFP_SESSION_ID)


def register_dfp_marker_click_callback(callback):
    import sys
    setattr(sys, _DFP_CLICK_CB_KEY, callback)
    register_session_click_handler(DFP_SESSION_ID, _handle_dfp_marker_click)


def register_dfp_schedule_window(window):
    """Keep Schedule window alive for marker click staging (pyRevit scope)."""
    import sys
    setattr(sys, _DFP_SCHEDULE_WINDOW_KEY, window)


def get_dfp_schedule_window():
    """Return the active Schedule window (survives pyRevit script scope disposal)."""
    import sys
    return getattr(sys, _DFP_SCHEDULE_WINDOW_KEY, None)


def _live_active_flags_for_door(global_id, sorted_keys):
    """Read pending-aware DFP flags from the Schedule table, if open."""
    win = get_dfp_schedule_window()
    if win is None or not sorted_keys:
        return None
    try:
        row = win._row_for_global_id(global_id)
        if row is None:
            return None
        return [win._dfp_cell_active(row, k) for k in sorted_keys]
    except Exception:
        return None


def unregister_dfp_marker_click_callback():
    import sys
    if hasattr(sys, _DFP_CLICK_CB_KEY):
        delattr(sys, _DFP_CLICK_CB_KEY)
    unregister_session_click_handler(DFP_SESSION_ID)


def _get_dfp_marker_click_callback():
    import sys
    return getattr(sys, _DFP_CLICK_CB_KEY, None)


def _store_dfp_uiapp(uiapp):
    import sys
    setattr(sys, _DFP_UIAPP_KEY, uiapp)


def _get_dfp_uiapp():
    import sys
    return getattr(sys, _DFP_UIAPP_KEY, None)


def _model_to_screen(uiview, pt):
    rect = uiview.GetWindowRectangle()
    corners = uiview.GetZoomCorners()
    if corners is None or len(corners) < 2:
        return None, None
    a = corners[0]
    b = corners[1]
    dx_model = b.X - a.X
    dy_model = b.Y - a.Y
    if abs(dx_model) < 1e-9:
        dx_model = 1e-9
    if abs(dy_model) < 1e-9:
        dy_model = 1e-9
    rel_x = (pt.X - a.X) / dx_model
    rel_y = (pt.Y - a.Y) / dy_model
    sx = rect.Left + rel_x * (rect.Right - rect.Left)
    sy = rect.Bottom + rel_y * (rect.Top - rect.Bottom)
    return sx, sy


def _screen_to_model(uiview, sx, sy, z):
    rect = uiview.GetWindowRectangle()
    corners = uiview.GetZoomCorners()
    if corners is None or len(corners) < 2:
        return None
    a = corners[0]
    b = corners[1]
    w = float(rect.Right - rect.Left)
    h = float(rect.Top - rect.Bottom)
    if abs(w) < 1.0:
        w = 1.0
    if abs(h) < 1.0:
        h = 1.0
    rel_x = (float(sx) - rect.Left) / w
    rel_y = (float(sy) - rect.Bottom) / h
    x = a.X + rel_x * (b.X - a.X)
    y = a.Y + rel_y * (b.Y - a.Y)
    return XYZ(x, y, z)


def _mouse_left_down():
    import clr
    clr.AddReference("System.Windows.Forms")
    from System.Windows.Forms import Control, MouseButtons
    return Control.MouseButtons == MouseButtons.Left


def _cursor_position():
    import clr
    clr.AddReference("System.Windows.Forms")
    from System.Windows.Forms import Cursor
    p = Cursor.Position
    return float(p.X), float(p.Y)


def _handle_dfp_marker_click(payload, command_data):
    if payload is None or not payload.get("dfp_cell"):
        return
    if not payload.get("clickable", True):
        return
    p = dict(payload)
    shim = _ClickShim(command_data.Index, command_data.Document)
    cb = _get_dfp_marker_click_callback()
    if cb is None:
        return
    try:
        cb(p, shim)
    except Exception:
        pass


def update_dfp_door_after_toggle(document, command_data, door_payload,
                                   sorted_keys, active_flags):
    """Refresh summary composite + overlay cells after staged toggle."""
    try:
        driver = find_dfp_driver(document)
        if driver is not None:
            driver.update_door_graphics(
                document, door_payload, sorted_keys, active_flags)
            return
        idx = command_data.Index
        active_codes = []
        for key, active in zip(sorted_keys, active_flags):
            if active:
                code = code_from_value_key(key)
                if code:
                    active_codes.append(code)
        img_path, _, _ = get_dfp_summary_composite(active_codes)
        comp_idx = door_payload.get("composite_ctrl_idx")
        if comp_idx is not None:
            x = float(door_payload.get("x", 0))
            y = float(door_payload.get("y", 0))
            z = float(door_payload.get("z", 0))
            update_control_image(document, comp_idx, img_path, XYZ(x, y, z))
    except Exception:
        pass


def start_or_get_dfp_driver(uiapp, document, view_points_map,
                            get_element_id_value, logger=None):
    _store_dfp_uiapp(uiapp)
    driver = find_dfp_driver(document)
    if driver is not None:
        if view_points_map:
            driver.set_sessions(view_points_map)
        return driver
    if not view_points_map:
        return None
    driver = DfpMarkerDriver(
        uiapp, document, view_points_map, get_element_id_value,
        logger=logger)
    driver.start()
    return driver


class DfpMarkerDriver(ViewMarkerDriver):
    """Summary composite per door; hover expands clickable cell overlays."""

    def __init__(self, uiapp, document, view_points_map,
                 get_element_id_value, logger=None):
        ViewMarkerDriver.__init__(
            self, uiapp, document, view_points_map, get_element_id_value,
            session_id=DFP_SESSION_ID,
            marker_style=DFP_MARKER_STYLE,
            sheet_entry_builder=None,
            sheet_tooltip="",
            single_tooltip="DFP",
            cluster_tooltip_template="{count} doors",
            logger=logger,
            refresh_active_view=False)
        _store_dfp_uiapp(uiapp)
        self._composite_by_ctrl = {}
        self._overlay_by_view = {}
        self._expanded_gid = {}
        self._hover_slot = {}
        self._last_hover_ms = 0
        self._mouse_down_by_view = {}
        self._last_synthetic_click_ms = 0
        self._interaction_paused = False

    def set_interaction_paused(self, paused):
        """Pause idling/hover/TGM work (e.g. during CDE Apply)."""
        self._interaction_paused = bool(paused)
        if paused:
            self._clear_all_dfp_graphics()

    def hide_all_graphics_visibility(self):
        """Hide composites, overlays, and cluster markers — no RemoveControl."""
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            return 0
        hidden = 0
        for idx in list(self._composite_by_ctrl.keys()):
            try:
                tgm.SetVisibility(int(idx), False)
                hidden += 1
            except Exception:
                pass
        for entries in list(self._overlay_by_view.values()):
            for entry in entries:
                try:
                    tgm.SetVisibility(int(entry["idx"]), False)
                    hidden += 1
                except Exception:
                    pass
        for indices in list(self._control_indices.values()):
            for idx in indices:
                try:
                    tgm.SetVisibility(int(idx), False)
                    hidden += 1
                except Exception:
                    pass
        for indices in list(self._sheet_control_indices.values()):
            for idx in indices:
                try:
                    tgm.SetVisibility(int(idx), False)
                    hidden += 1
                except Exception:
                    pass
        self._expanded_gid = {}
        self._hover_slot = {}
        self._mouse_down_by_view = {}
        return hidden

    def _clear_all_dfp_graphics(self):
        """Remove summary composites, hover overlays, and cluster markers."""
        tgm = get_temporary_graphics_manager(self._doc)
        for vid in list(self._overlay_by_view.keys()):
            self._clear_overlays(vid, tgm=tgm)
        composite_count = len(self._composite_by_ctrl)
        if tgm is not None:
            for idx in list(self._composite_by_ctrl.keys()):
                try:
                    unregister_control_click(self._doc, idx)
                except Exception:
                    pass
                try:
                    tgm.RemoveControl(idx)
                except Exception:
                    pass
        self._composite_by_ctrl = {}
        self._overlay_by_view = {}
        self._expanded_gid = {}
        self._hover_slot = {}
        self._mouse_down_by_view = {}
        ViewMarkerDriver._clear_all_markers(self)

    def soft_stop(self):
        """Hide graphics and unregister handlers — no RemoveControl."""
        self._enabled = False
        self._interaction_paused = True
        hidden = self.hide_all_graphics_visibility()
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
            getattr(sys, "_pyBS_view_marker_drivers").remove(self)
        except (AttributeError, ValueError):
            pass
        return hidden

    def stop(self):
        self._clear_all_dfp_graphics()
        ViewMarkerDriver.stop(self)

    def _reset_hover_state(self, view_id_value=None):
        if view_id_value is None:
            self._composite_by_ctrl = {}
            self._overlay_by_view = {}
            self._expanded_gid = {}
            self._hover_slot = {}
            return
        vid = _normalize_view_id(view_id_value)
        self._overlay_by_view.pop(vid, None)
        self._expanded_gid.pop(vid, None)
        self._hover_slot.pop(vid, None)
        drop = []
        for idx, payload in self._composite_by_ctrl.items():
            if _normalize_view_id(payload.get("view_id")) == vid:
                drop.append(idx)
        for idx in drop:
            self._composite_by_ctrl.pop(idx, None)

    def _set_summary_visible(self, door_payload, visible):
        comp_idx = door_payload.get("composite_ctrl_idx")
        if comp_idx is None:
            return
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            return
        try:
            tgm.SetVisibility(int(comp_idx), bool(visible))
        except Exception:
            pass

    def _clear_overlays(self, view_id_value, tgm=None):
        vid = _normalize_view_id(view_id_value)
        gid = self._expanded_gid.get(vid)
        if gid:
            door = self._payload_for_gid(gid)
            if door is not None:
                self._set_summary_visible(door, True)
        entries = self._overlay_by_view.pop(vid, [])
        if tgm is None:
            tgm = get_temporary_graphics_manager(self._doc)
        if tgm is not None:
            for entry in entries:
                try:
                    unregister_control_click(self._doc, entry["idx"])
                except Exception:
                    pass
                try:
                    tgm.RemoveControl(entry["idx"])
                except Exception:
                    pass
        self._expanded_gid.pop(vid, None)
        self._hover_slot.pop(vid, None)

    def _build_overlays(self, view_id_value, door_payload, uiview):
        vid = _normalize_view_id(view_id_value)
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None or uiview is None:
            return
        self._clear_overlays(vid, tgm=tgm)

        sorted_keys = door_payload.get("sorted_keys") or []
        active_flags = door_payload.get("active_flags") or []
        layout = door_payload.get("overlay_layout")
        if not sorted_keys or layout is None:
            return

        live_flags = _live_active_flags_for_door(
            door_payload.get("global_id"), sorted_keys)
        if live_flags is not None and len(live_flags) == len(sorted_keys):
            active_flags = live_flags
            door_payload = dict(door_payload)
            door_payload["active_flags"] = list(active_flags)

        anchor = XYZ(
            float(door_payload["x"]),
            float(door_payload["y"]),
            float(door_payload["z"]),
        )
        sx, sy = _model_to_screen(uiview, anchor)
        if sx is None:
            return

        owner_view_id = _owner_view_id_for_control(vid)
        entries = []
        for idx, key in enumerate(sorted_keys):
            code = code_from_value_key(key)
            if not code:
                continue
            active = False
            if idx < len(active_flags):
                active = bool(active_flags[idx])
            cell_sx, cell_sy = composite_slot_screen_center(
                layout, idx, sx, sy)
            xyz = _screen_to_model(uiview, cell_sx, cell_sy, anchor.Z)
            if xyz is None:
                continue
            try:
                img_path = get_cell_marker_path(code, active=active, hover=False)
                if not _validate_marker_image(img_path):
                    continue
                data = InCanvasControlData(img_path, xyz)
                ctrl_idx = tgm.AddControl(data, owner_view_id)
                try:
                    tgm.SetVisibility(ctrl_idx, True)
                except Exception:
                    pass
                cell_payload = {
                    "clickable": True,
                    "dfp_cell": True,
                    "session_id": DFP_SESSION_ID,
                    "global_id": door_payload.get("global_id"),
                    "param_key": key,
                    "code": code,
                    "slot_index": idx,
                    "sorted_keys": list(sorted_keys),
                    "active_flags": list(active_flags),
                    "door_payload": dict(door_payload),
                    "composite_ctrl_idx": door_payload.get("composite_ctrl_idx"),
                }
                register_control_click(self._doc, ctrl_idx, cell_payload)
                label = _LABEL_BY_CODE.get(code, code)
                tip = u"{} {} — click to toggle".format(code, label)
                try:
                    tgm.SetTooltip(ctrl_idx, tip)
                except Exception:
                    pass
                entries.append({
                    "idx": ctrl_idx,
                    "slot": idx,
                    "code": code,
                    "active": active,
                })
            except Exception:
                pass

        if entries:
            self._overlay_by_view[vid] = entries
            self._expanded_gid[vid] = door_payload.get("global_id")
            self._hover_slot[vid] = None
            comp_idx = door_payload.get("composite_ctrl_idx")
            if comp_idx is not None:
                self._composite_by_ctrl[comp_idx] = door_payload
            self._set_summary_visible(door_payload, False)

    def _update_cell_hover(self, view_id_value, uiview, slot):
        vid = _normalize_view_id(view_id_value)
        if self._hover_slot.get(vid) == slot:
            return
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            return
        entries = self._overlay_by_view.get(vid) or []
        door_payload = None
        for idx, p in self._composite_by_ctrl.items():
            if p.get("global_id") == self._expanded_gid.get(vid):
                door_payload = p
                break
        if door_payload is None:
            return
        sorted_keys = door_payload.get("sorted_keys") or []
        active_flags = door_payload.get("active_flags") or []

        for entry in entries:
            s = entry["slot"]
            code = entry["code"]
            active = entry["active"]
            hover = (slot is not None and s == slot)
            try:
                img_path = get_cell_marker_path(code, active=active, hover=hover)
                key = sorted_keys[s] if s < len(sorted_keys) else None
                if key is None:
                    continue
                layout = door_payload.get("overlay_layout")
                anchor = XYZ(
                    float(door_payload["x"]),
                    float(door_payload["y"]),
                    float(door_payload["z"]),
                )
                sx, sy = _model_to_screen(uiview, anchor)
                if sx is None:
                    continue
                cell_sx, cell_sy = composite_slot_screen_center(
                    layout, s, sx, sy)
                xyz = _screen_to_model(uiview, cell_sx, cell_sy, anchor.Z)
                if xyz is None:
                    continue
                update_control_image(self._doc, entry["idx"], img_path, xyz)
            except Exception:
                pass
        self._hover_slot[vid] = slot

    def update_door_graphics(self, document, door_payload, sorted_keys,
                             active_flags):
        """After table/marker toggle — refresh summary + overlays."""
        if self._interaction_paused:
            return
        gid = door_payload.get("global_id")
        active_codes = []
        for key, active in zip(sorted_keys, active_flags):
            if active:
                code = code_from_value_key(key)
                if code:
                    active_codes.append(code)
        img_path, _, _ = get_dfp_summary_composite(active_codes)
        door_payload = dict(door_payload)
        door_payload["active_flags"] = list(active_flags)
        door_payload["image_path"] = img_path

        comp_idx = door_payload.get("composite_ctrl_idx")
        if comp_idx is not None:
            xyz = XYZ(
                float(door_payload["x"]),
                float(door_payload["y"]),
                float(door_payload["z"]),
            )
            update_control_image(document, comp_idx, img_path, xyz)
            self._composite_by_ctrl[comp_idx] = door_payload

        for vid, exp_gid in list(self._expanded_gid.items()):
            if exp_gid == gid:
                uiview = _get_uiview_for_view(
                    self._uiapp, vid, self._get_element_id_value)
                self._build_overlays(vid, door_payload, uiview)
                break

    def _payload_for_gid(self, gid):
        for payload in self._composite_by_ctrl.values():
            if payload.get("global_id") == gid:
                return payload
        return None

    def _cursor_over_door(self, uiview, payload, expanded=False):
        cx, cy = _cursor_position()
        anchor = XYZ(
            float(payload["x"]), float(payload["y"]), float(payload["z"]))
        sx, sy = _model_to_screen(uiview, anchor)
        if sx is None:
            return False
        if expanded:
            layout = payload.get("overlay_layout") or {}
        else:
            layout = payload.get("summary_layout") or payload.get("overlay_layout") or {}
        bmp_px = float(layout.get("bmp_px", 28))
        half = bmp_px / 2.0
        return (sx - half <= cx <= sx + half and sy - half <= cy <= sy + half)

    def _find_hovered_composite(self, uiview, view_id_value):
        vid = _normalize_view_id(view_id_value)
        for payload in self._composite_by_ctrl.values():
            if _normalize_view_id(payload.get("view_id")) != vid:
                continue
            if self._cursor_over_door(uiview, payload, expanded=False):
                return payload
        return None

    def _find_hovered_overlay_slot(self, uiview, view_id_value, door_payload):
        vid = _normalize_view_id(view_id_value)
        entries = self._overlay_by_view.get(vid)
        if not entries:
            return None
        cx, cy = _cursor_position()
        anchor = XYZ(
            float(door_payload["x"]),
            float(door_payload["y"]),
            float(door_payload["z"]),
        )
        sx, sy = _model_to_screen(uiview, anchor)
        if sx is None:
            return None
        layout = door_payload.get("overlay_layout")
        if layout is None:
            return None
        bmp_px = float(layout.get("bmp_px", 0))
        lx = cx - sx + (bmp_px / 2.0)
        ly = cy - sy + (bmp_px / 2.0)
        return hit_test_composite_slot(lx, ly, layout)

    def _dispatch_overlay_slot_click(self, door_payload, slot_index):
        """Synthetic click when TGM OnClick does not fire (cursor hit-test path)."""
        sorted_keys = door_payload.get("sorted_keys") or []
        if slot_index is None or slot_index < 0 or slot_index >= len(sorted_keys):
            return
        key = sorted_keys[slot_index]
        code = code_from_value_key(key)
        cell_payload = {
            "clickable": True,
            "dfp_cell": True,
            "session_id": DFP_SESSION_ID,
            "global_id": door_payload.get("global_id"),
            "param_key": key,
            "code": code,
            "slot_index": slot_index,
            "sorted_keys": list(sorted_keys),
            "active_flags": list(door_payload.get("active_flags") or []),
            "door_payload": dict(door_payload),
            "composite_ctrl_idx": door_payload.get("composite_ctrl_idx"),
        }
        _handle_dfp_marker_click(cell_payload, _ClickShim(0, self._doc))

    def _apply_pending_graphics_refresh(self):
        """Refresh summary/overlay from staged table edits (idling thread, no ExternalEvent)."""
        win = get_dfp_schedule_window()
        if win is None or getattr(win, "_apply_in_flight", False):
            return
        if getattr(win, "_apply_revit_quarantine", False):
            return
        if getattr(win, "_dfp_markers_need_manual_reset", False):
            return
        gid = getattr(win, "_dfp_graphics_refresh_gid", None)
        if not gid:
            return
        win._dfp_graphics_refresh_gid = None
        door = self._payload_for_gid(gid)
        if door is None:
            return
        sorted_keys = door.get("sorted_keys") or []
        live_flags = _live_active_flags_for_door(gid, sorted_keys)
        if live_flags is None:
            return
        self.update_door_graphics(self._doc, door, sorted_keys, live_flags)

    def _poll_dfp_hover(self):
        if is_dfp_apply_blocked() or not self._enabled or self._interaction_paused:
            return
        self._apply_pending_graphics_refresh()
        now_ms = int(time.time() * 1000)
        if now_ms - self._last_hover_ms < _HOVER_THROTTLE_MS:
            return
        self._last_hover_ms = now_ms

        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None:
            return
        try:
            active_view = uidoc.ActiveView
            if active_view is None:
                return
            vid = _normalize_view_id(
                self._get_element_id_value(active_view.Id))
        except Exception:
            return
        if vid not in self._sessions:
            return
        uiview = _get_uiview_for_view(
            self._uiapp, vid, self._get_element_id_value)
        if uiview is None:
            return

        exp_gid = self._expanded_gid.get(vid)
        door = self._payload_for_gid(exp_gid) if exp_gid else None

        if door is not None and not self._cursor_over_door(uiview, door, expanded=True):
            self._clear_overlays(vid)
            door = None

        if door is None:
            door = self._find_hovered_composite(uiview, vid)
            if door is not None:
                self._build_overlays(vid, door, uiview)

        if door is not None:
            slot = self._find_hovered_overlay_slot(uiview, vid, door)
            self._update_cell_hover(vid, uiview, slot)
            mouse_down = _mouse_left_down()
            was_down = bool(self._mouse_down_by_view.get(vid, False))
            self._mouse_down_by_view[vid] = mouse_down
            if (
                slot is not None
                and mouse_down
                and not was_down
                and now_ms - self._last_synthetic_click_ms > 150
            ):
                self._last_synthetic_click_ms = now_ms
                self._dispatch_overlay_slot_click(door, slot)

    def _on_idling(self, sender, args):
        if is_dfp_apply_blocked() or self._interaction_paused:
            return
        ViewMarkerDriver._on_idling(self, sender, args)
        self._poll_dfp_hover()

    def _refresh_view_markers(self, view_id_value, corners=None, force=False):
        import time as _time

        view_id_value = _normalize_view_id(view_id_value)
        if self._refresh_in_progress:
            return
        now_ms = int(_time.time() * 1000)
        last_ms = self._last_view_refresh_ms.get(view_id_value, 0)
        if not force and now_ms - last_ms < _VIEW_REFRESH_MS:
            return

        point_dicts = self._sessions.get(view_id_value)
        if not point_dicts:
            self._clear_view_markers(view_id_value)
            self._reset_hover_state(view_id_value)
            return
        if len(point_dicts) > self._style.max_markers_per_view:
            point_dicts = point_dicts[: self._style.max_markers_per_view]

        uiview = _get_uiview_for_view(
            self._uiapp, view_id_value, self._get_element_id_value)
        if corners is None and uiview is not None:
            try:
                corners = uiview.GetZoomCorners()
            except Exception:
                corners = None

        ensure_temporary_graphics_handler(self._doc, logger=self._logger)
        tgm = get_temporary_graphics_manager(self._doc)
        if tgm is None:
            return

        exp_gid = self._expanded_gid.get(view_id_value)
        exp_door = None
        if exp_gid:
            payload = self._payload_for_gid(exp_gid)
            if payload is not None:
                exp_door = dict(payload)

        self._clear_overlays(view_id_value, tgm=tgm)
        self._reset_hover_state(view_id_value)
        clear_control_clicks(self._doc)
        self._refresh_in_progress = True
        try:
            owner_view_id = _owner_view_id_for_control(view_id_value)
            self._clear_view_markers(view_id_value)
            new_indices = []

            if corners is not None:
                span = _view_span_from_corners(corners)
            else:
                span = _view_span_from_view3d(self._doc, view_id_value)
            radius = self._style.span_factor * span

            indexed_anchors = []
            for i, p in enumerate(point_dicts):
                try:
                    pt = XYZ(float(p["x"]), float(p["y"]), float(p["z"]))
                except Exception:
                    continue
                indexed_anchors.append((i, pt))

            if not indexed_anchors:
                self._control_indices[view_id_value] = []
                return

            xyz_list = [pt for _, pt in indexed_anchors]
            clusters = cluster_points_model_space(xyz_list, radius)

            used_doors = set()
            for cluster in clusters:
                count = cluster["count"]
                xyz = cluster["xyz"]
                if count > 1:
                    try:
                        img_path = get_marker_image_path(
                            count, "cluster", self._style)
                        data = InCanvasControlData(img_path, xyz)
                        ctrl_idx = tgm.AddControl(data, owner_view_id)
                        new_indices.append(ctrl_idx)
                        try:
                            tgm.SetVisibility(ctrl_idx, True)
                        except Exception:
                            pass
                        try:
                            tgm.SetTooltip(
                                ctrl_idx,
                                self._cluster_tooltip_template.format(
                                    count=count))
                        except Exception:
                            pass
                    except Exception:
                        pass
                    continue

                best_i = None
                best_d = None
                for orig_i, pt in indexed_anchors:
                    if orig_i in used_doors:
                        continue
                    d = xyz.DistanceTo(pt)
                    if best_d is None or d < best_d:
                        best_d = d
                        best_i = orig_i
                if best_i is None or best_d is None or best_d > 0.05:
                    continue
                used_doors.add(best_i)
                p = dict(point_dicts[best_i])
                try:
                    img_path = p.get("image_path")
                    if not img_path or not _validate_marker_image(img_path):
                        continue
                    anchor = XYZ(float(p["x"]), float(p["y"]), float(p["z"]))
                    data = InCanvasControlData(img_path, anchor)
                    ctrl_idx = tgm.AddControl(data, owner_view_id)
                    new_indices.append(ctrl_idx)
                    p["view_id"] = view_id_value
                    p["composite_ctrl_idx"] = ctrl_idx
                    self._composite_by_ctrl[ctrl_idx] = p
                    try:
                        tgm.SetVisibility(ctrl_idx, True)
                    except Exception:
                        pass
                    tip = p.get("tooltip") or self._single_tooltip
                    try:
                        tgm.SetTooltip(ctrl_idx, tip)
                    except Exception:
                        pass
                except Exception:
                    pass

            self._control_indices[view_id_value] = new_indices
            self._last_view_refresh_ms[view_id_value] = now_ms
            if force and corners is not None:
                from revit.view_markers import _corners_key
                self._last_corners[view_id_value] = _corners_key(corners)
        finally:
            self._refresh_in_progress = False
            if exp_door is not None and exp_gid and uiview is not None:
                door = self._payload_for_gid(exp_gid)
                if door is None:
                    door = exp_door
                else:
                    door = dict(door)
                    door["sorted_keys"] = exp_door.get(
                        "sorted_keys", door.get("sorted_keys"))
                    door["active_flags"] = exp_door.get(
                        "active_flags", door.get("active_flags"))
                    door["overlay_layout"] = exp_door.get(
                        "overlay_layout", door.get("overlay_layout"))
                try:
                    self._build_overlays(view_id_value, door, uiview)
                except Exception:
                    pass
