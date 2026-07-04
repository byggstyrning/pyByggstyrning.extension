# -*- coding: utf-8 -*-
"""Phase label as a WPF overlay window pinned over the active 3D view.

Alternative to revit.phase_label (TemporaryGraphicsManager): a borderless,
topmost-within-Revit, click-through WPF window positioned at the top-left
corner of the active UIView's window rectangle. Because the anchor is in
screen pixels (UIView.GetWindowRectangle), the label stays glued to the
corner *during* pan/zoom/orbit — no snap-back. Pure UI: nothing is written
to the model and nothing prints. Works on any Revit version with UIView
(no TemporaryGraphicsManager requirement).
"""

import ctypes
import sys
import time

import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

from System import TimeSpan
from System.Windows import (
    Window, WindowStyle, ResizeMode, SizeToContent, Thickness,
    CornerRadius, PresentationSource, FontWeights,
)
from System.Windows.Controls import Border, TextBlock
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Media import SolidColorBrush, Color, Brushes, FontFamily
from System.Windows.Threading import DispatcherTimer

from Autodesk.Revit.DB import View3D

from revit.compat import get_element_id_value
from revit.phase_label import _label_text_for_view
from revit.view_markers import _normalize_view_id

_DRIVERS_SYS_KEY = '_pyBS_phase_hud_drivers'
_THROTTLE_MS = 300
_MARGIN_PX = 12
_MOVE_WATCH_MS = 100

_GWL_EXSTYLE = -20
_WS_EX_TRANSPARENT = 0x00000020
_WS_EX_TOOLWINDOW = 0x00000080
_WS_EX_NOACTIVATE = 0x08000000


def _apply_click_through(hwnd_int):
    """Make the overlay ignore mouse input and never steal focus."""
    try:
        user32 = ctypes.windll.user32
        try:
            get_fn = user32.GetWindowLongPtrW
            set_fn = user32.SetWindowLongPtrW
            long_type = ctypes.c_longlong
        except AttributeError:
            get_fn = user32.GetWindowLongW
            set_fn = user32.SetWindowLongW
            long_type = ctypes.c_long
        get_fn.restype = long_type
        get_fn.argtypes = [ctypes.c_void_p, ctypes.c_int]
        set_fn.restype = long_type
        set_fn.argtypes = [ctypes.c_void_p, ctypes.c_int, long_type]
        style = get_fn(hwnd_int, _GWL_EXSTYLE)
        set_fn(
            hwnd_int, _GWL_EXSTYLE,
            style | _WS_EX_TRANSPARENT | _WS_EX_TOOLWINDOW | _WS_EX_NOACTIVATE)
        return True
    except Exception:
        return False


class _RECT(ctypes.Structure):
    _fields_ = [
        ('left', ctypes.c_long),
        ('top', ctypes.c_long),
        ('right', ctypes.c_long),
        ('bottom', ctypes.c_long),
    ]


def _window_rect(hwnd_int):
    """Screen rect of a window via Win32 (safe outside Revit API context)."""
    if not hwnd_int:
        return None
    try:
        rect = _RECT()
        fn = ctypes.windll.user32.GetWindowRect
        fn.argtypes = [ctypes.c_void_p, ctypes.POINTER(_RECT)]
        fn.restype = ctypes.c_int
        if fn(hwnd_int, ctypes.byref(rect)):
            return (rect.left, rect.top, rect.right, rect.bottom)
    except Exception:
        pass
    return None


def _revit_main_window_handle(uiapp):
    try:
        hwnd = uiapp.MainWindowHandle
        if hwnd is not None:
            return hwnd
    except Exception:
        pass
    try:
        from System.Diagnostics import Process
        return Process.GetCurrentProcess().MainWindowHandle
    except Exception:
        return None


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


def find_phase_hud_driver(document):
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


def stop_phase_hud_driver(document):
    driver = find_phase_hud_driver(document)
    if driver is not None:
        driver.stop()
        return True
    return False


def start_phase_hud_driver(uiapp, document, logger=None):
    driver = find_phase_hud_driver(document)
    if driver is not None:
        driver.refresh()
        return driver
    driver = PhaseHudDriver(uiapp, document, logger=logger)
    driver.start()
    return driver


class PhaseHudDriver(object):
    """Idling-driven WPF badge pinned to the active 3D view's corner."""

    def __init__(self, uiapp, document, logger=None):
        self._uiapp = uiapp
        self._doc = document
        self._logger = logger
        self._window = None
        self._text_block = None
        self._visible = False
        self._state_key = None
        self._last_tick_ms = 0
        self._idling_handler = None
        self._view_activated_handler = None
        self._owner_hwnd_int = None
        self._last_owner_rect = None
        self._move_timer = None

    def _build_window(self):
        window = Window()
        window.WindowStyle = getattr(WindowStyle, 'None')
        window.ResizeMode = ResizeMode.NoResize
        window.AllowsTransparency = True
        window.Background = Brushes.Transparent
        window.SizeToContent = SizeToContent.WidthAndHeight
        window.ShowInTaskbar = False
        window.ShowActivated = False
        window.Focusable = False
        window.IsHitTestVisible = False
        window.Topmost = False

        text = TextBlock()
        text.Foreground = SolidColorBrush(Color.FromArgb(255, 245, 245, 245))
        text.FontFamily = FontFamily('Segoe UI')
        text.FontSize = 13.0
        text.FontWeight = FontWeights.Bold

        badge = Border()
        badge.Background = SolidColorBrush(Color.FromArgb(230, 48, 50, 56))
        badge.BorderBrush = SolidColorBrush(Color.FromArgb(255, 110, 110, 118))
        badge.BorderThickness = Thickness(1)
        badge.CornerRadius = CornerRadius(4)
        badge.Padding = Thickness(10, 5, 10, 5)
        badge.Child = text
        window.Content = badge

        helper = WindowInteropHelper(window)
        owner = _revit_main_window_handle(self._uiapp)
        if owner is not None:
            try:
                helper.Owner = owner
                self._owner_hwnd_int = owner.ToInt64()
            except Exception:
                pass
        hwnd = helper.EnsureHandle()
        try:
            _apply_click_through(hwnd.ToInt64())
        except Exception:
            pass

        self._window = window
        self._text_block = text

    def start(self):
        self._build_window()
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
        # Idling is silent while Windows runs the modal move/size loop on the
        # Revit window, but WM_TIMER still gets dispatched there — so a
        # DispatcherTimer can hide the badge the moment the window starts
        # moving; the first Idling tick after release shows it again.
        try:
            self._move_timer = DispatcherTimer()
            self._move_timer.Interval = TimeSpan.FromMilliseconds(
                _MOVE_WATCH_MS)
            self._move_timer.Tick += self._on_move_watch_tick
            self._move_timer.Start()
        except Exception:
            self._move_timer = None
        if not hasattr(sys, _DRIVERS_SYS_KEY):
            setattr(sys, _DRIVERS_SYS_KEY, [])
        getattr(sys, _DRIVERS_SYS_KEY).append(self)
        self.refresh()

    def stop(self):
        if self._move_timer is not None:
            try:
                self._move_timer.Stop()
                self._move_timer.Tick -= self._on_move_watch_tick
            except Exception:
                pass
            self._move_timer = None
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
        if self._window is not None:
            try:
                self._window.Close()
            except Exception:
                pass
            self._window = None
            self._text_block = None
        self._visible = False
        try:
            getattr(sys, _DRIVERS_SYS_KEY).remove(self)
        except (AttributeError, ValueError):
            pass

    def refresh(self):
        self._state_key = None
        self._sync(force=True)

    def _hide(self):
        if self._window is not None and self._visible:
            try:
                self._window.Hide()
            except Exception:
                pass
        self._visible = False
        self._state_key = None

    def _on_view_activated(self, sender, args):
        try:
            self._sync(force=True)
        except Exception:
            pass

    def _on_move_watch_tick(self, sender, args):
        """Hide while the Revit window is being moved or resized."""
        try:
            rect = _window_rect(self._owner_hwnd_int)
            if rect is None:
                return
            last = self._last_owner_rect
            self._last_owner_rect = rect
            if last is not None and rect != last:
                self._hide()
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

    def _move_to(self, rect):
        """Place the window at the view rect's top-left (device px -> DIP)."""
        left_px = rect.Left + _MARGIN_PX
        top_px = rect.Top + _MARGIN_PX
        scale_x = scale_y = 1.0
        try:
            source = PresentationSource.FromVisual(self._window)
            if source is not None:
                matrix = source.CompositionTarget.TransformFromDevice
                scale_x = matrix.M11
                scale_y = matrix.M22
        except Exception:
            pass
        self._window.Left = left_px * scale_x
        self._window.Top = top_px * scale_y

    def _sync(self, force=False):
        if self._window is None:
            return
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None or not uidoc.Document.Equals(self._doc):
            self._hide()
            return
        view = uidoc.ActiveView
        if view is None or not isinstance(view, View3D) or view.IsTemplate:
            self._hide()
            return

        text = _label_text_for_view(self._doc, view)
        if not text:
            self._hide()
            return

        uiview = _get_uiview(self._uiapp, view.Id)
        if uiview is None:
            self._hide()
            return
        try:
            rect = uiview.GetWindowRectangle()
        except Exception:
            self._hide()
            return

        key = (rect.Left, rect.Top, rect.Right, rect.Bottom, text)
        if not force and self._visible and key == self._state_key:
            return

        try:
            self._text_block.Text = text
            self._move_to(rect)
            if not self._visible:
                self._window.Show()
                self._visible = True
            self._state_key = key
        except Exception as ex:
            if self._logger is not None:
                try:
                    self._logger.warning(
                        "Phase HUD update failed: {}".format(ex))
                except Exception:
                    pass
