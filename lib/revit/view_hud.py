# -*- coding: utf-8 -*-
"""Reusable screen-anchored HUD badges for Revit views (WPF overlay).

A ViewHudDriver shows a small text badge pinned over the active view's
window rectangle. The anchor is in screen pixels (UIView.GetWindowRectangle)
so the badge stays put during pan/zoom/orbit. Machinery provided here:

- borderless, non-activating, owned WPF window (never steals Revit focus)
- muted palette following Revit's light/dark theme (UIThemeManager, 2024+)
- idle/hover opacity
- per-monitor DPI-aware positioning with configurable anchor
- deferred sizing: the badge is parked off-screen until WPF completes the
  first layout pass (SizeChanged), avoiding the mis-sized first frame that
  EnsureHandle + SizeToContent produces
- hidden while the Revit main window is moved/resized (DispatcherTimer
  still ticks inside the modal move loop; Idling does not fire there)
- optional click actions executed inside a Revit API context (ExternalEvent)

Build a new HUD with providers; None from text_provider hides the badge:

    from revit.view_hud import ViewHudDriver, find_hud_driver

    driver = ViewHudDriver(
        uiapp, doc, hud_id='my-hud',
        text_provider=lambda doc, view: 'text',   # or None to hide
        tooltip_provider=lambda text: 'hint',      # optional
        on_left_click=lambda uiapp: ...,           # optional, API context
        on_right_click=lambda uiapp: ...,          # optional, API context
    )
    driver.start()   # call from a command (needs API context)

Drivers are registered per (document, hud_id): find_hud_driver /
stop_hud_driver look them up, so toggle buttons stay stateless.
"""

import ctypes
import sys
import time

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

from System import TimeSpan
from System.Windows import (
    Window, WindowStyle, ResizeMode, SizeToContent, Thickness,
    CornerRadius, PresentationSource,
)
from System.Windows.Controls import Border, TextBlock
from System.Windows.Input import Cursors
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Media import SolidColorBrush, Color, Brushes, FontFamily
from System.Windows.Threading import DispatcherTimer

from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

from revit.compat import get_element_id_value

_DRIVERS_SYS_KEY = '_pyBS_view_hud_drivers'
_THROTTLE_MS = 300
_MOVE_WATCH_MS = 100
_OFFSCREEN = -32000

_GWL_EXSTYLE = -20
_WS_EX_TOOLWINDOW = 0x00000080
_WS_EX_NOACTIVATE = 0x08000000


class HudStyle(object):
    """Placement and look of a HUD badge; palettes are ARGB tuples."""

    def __init__(
        self,
        anchor='top-center',            # 'top-left' | 'top-center' | 'top-right'
        margin_px=12,
        idle_opacity=0.5,
        hover_opacity=1.0,
        font_family='Segoe UI',
        font_size=12.0,
        corner_radius=4.0,
        padding=(8, 3, 8, 3),
        light_palette=((215, 246, 246, 246),    # background
                       (60, 0, 0, 0),           # border
                       (255, 64, 64, 64)),      # text
        dark_palette=((215, 43, 43, 43),
                      (60, 255, 255, 255),
                      (255, 214, 214, 214)),
    ):
        self.anchor = anchor
        self.margin_px = margin_px
        self.idle_opacity = idle_opacity
        self.hover_opacity = hover_opacity
        self.font_family = font_family
        self.font_size = font_size
        self.corner_radius = corner_radius
        self.padding = padding
        self.light_palette = light_palette
        self.dark_palette = dark_palette


DEFAULT_HUD_STYLE = HudStyle()


def revit_theme_is_dark():
    """True when Revit runs its dark canvas/UI theme (2024+ API)."""
    try:
        from Autodesk.Revit.UI import UIThemeManager
        try:
            theme = UIThemeManager.CurrentCanvasTheme
        except Exception:
            theme = UIThemeManager.CurrentTheme
        return 'dark' in str(theme).lower()
    except Exception:
        return False


def _apply_overlay_styles(hwnd_int):
    """Keep the overlay out of the taskbar/alt-tab and never steal focus.

    The badge stays clickable (no WS_EX_TRANSPARENT); WS_EX_NOACTIVATE
    keeps keyboard focus in Revit throughout.
    """
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
            style | _WS_EX_TOOLWINDOW | _WS_EX_NOACTIVATE)
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


def window_rect(hwnd_int):
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


def revit_main_window_handle(uiapp):
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


def get_uiview(uiapp, view_id):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    target = get_element_id_value(view_id)
    for uiv in uidoc.GetOpenUIViews():
        try:
            if get_element_id_value(uiv.ViewId) == target:
                return uiv
        except Exception:
            continue
    return None


def find_hud_driver(document, hud_id):
    drivers = getattr(sys, _DRIVERS_SYS_KEY, None)
    if not drivers:
        return None
    for driver in drivers:
        try:
            if driver._hud_id == hud_id and driver._doc.Equals(document):
                return driver
        except Exception:
            continue
    return None


def stop_hud_driver(document, hud_id):
    driver = find_hud_driver(document, hud_id)
    if driver is not None:
        driver.stop()
        return True
    return False


class _HudActionEventHandler(IExternalEventHandler):
    """Runs a queued HUD click action inside a Revit API context."""

    def __init__(self):
        self.action = None
        self.driver = None

    def Execute(self, uiapp):
        try:
            action = self.action
            self.action = None
            if action is not None:
                action(uiapp)
            if self.driver is not None:
                self.driver.refresh()
        except Exception:
            pass

    def GetName(self):
        return "pyBS View HUD action"


class ViewHudDriver(object):
    """Idling-driven WPF badge pinned over the active view.

    text_provider(document, view) supplies the badge text; return None to
    hide. on_left_click / on_right_click are optional callables taking
    uiapp, executed inside a Revit API context (transactions allowed).
    """

    def __init__(self, uiapp, document, hud_id,
                 text_provider,
                 tooltip_provider=None,
                 on_left_click=None,
                 on_right_click=None,
                 style=None,
                 logger=None):
        self._uiapp = uiapp
        self._doc = document
        self._hud_id = hud_id
        self._text_provider = text_provider
        self._tooltip_provider = tooltip_provider
        self._on_left_click = on_left_click
        self._on_right_click = on_right_click
        self._style = style or DEFAULT_HUD_STYLE
        self._logger = logger
        self._window = None
        self._text_block = None
        self._badge = None
        self._theme_dark = None
        self._action_handler = None
        self._action_event = None
        self._visible = False
        self._state_key = None
        self._last_view_rect = None
        self._last_tick_ms = 0
        self._idling_handler = None
        self._view_activated_handler = None
        self._owner_hwnd_int = None
        self._last_owner_rect = None
        self._move_timer = None
        self._last_error = None

    # ------------------------------------------------------------------ ui

    def _build_window(self):
        style = self._style
        window = Window()
        window.WindowStyle = getattr(WindowStyle, 'None')
        window.ResizeMode = ResizeMode.NoResize
        window.AllowsTransparency = True
        window.Background = Brushes.Transparent
        window.SizeToContent = SizeToContent.WidthAndHeight
        window.ShowInTaskbar = False
        window.ShowActivated = False
        window.Focusable = False
        window.Topmost = False
        window.Opacity = style.idle_opacity
        # park off-screen: the first layout pass happens after Show(), so
        # any mis-sized first frame is invisible; SizeChanged then anchors
        # the settled window at the real position
        window.Left = _OFFSCREEN
        window.Top = _OFFSCREEN

        text = TextBlock()
        text.FontFamily = FontFamily(style.font_family)
        text.FontSize = style.font_size

        badge = Border()
        badge.BorderThickness = Thickness(1)
        badge.CornerRadius = CornerRadius(style.corner_radius)
        pad = style.padding
        badge.Padding = Thickness(pad[0], pad[1], pad[2], pad[3])
        badge.Child = text
        if self._on_left_click is not None or self._on_right_click is not None:
            badge.Cursor = Cursors.Hand
        badge.MouseLeftButtonDown += self._on_badge_left_click
        badge.MouseRightButtonDown += self._on_badge_right_click
        badge.MouseEnter += self._on_badge_mouse_enter
        badge.MouseLeave += self._on_badge_mouse_leave
        window.Content = badge

        # owner keeps the overlay above Revit only (not other apps); the
        # Win32 styles are applied once the HWND exists (SourceInitialized)
        # instead of forcing it early with EnsureHandle, which breaks
        # SizeToContent's first layout pass
        helper = WindowInteropHelper(window)
        owner = revit_main_window_handle(self._uiapp)
        if owner is not None:
            try:
                helper.Owner = owner
                self._owner_hwnd_int = owner.ToInt64()
            except Exception:
                pass
        window.SourceInitialized += self._on_source_initialized
        window.SizeChanged += self._on_window_size_changed

        self._window = window
        self._text_block = text
        self._badge = badge
        self._apply_theme()

    def _on_source_initialized(self, sender, args):
        try:
            hwnd = WindowInteropHelper(self._window).Handle
            _apply_overlay_styles(hwnd.ToInt64())
        except Exception:
            pass

    def _on_window_size_changed(self, sender, args):
        # fires when the first layout pass (and any text change) resizes
        # the window; re-anchor using the last known view rect
        if self._last_view_rect is not None:
            try:
                self._move_to(self._last_view_rect)
            except Exception:
                pass

    def _apply_theme(self):
        """Muted badge palette following Revit's light/dark theme."""
        dark = revit_theme_is_dark()
        if dark == self._theme_dark or self._badge is None:
            return
        self._theme_dark = dark
        palette = (self._style.dark_palette if dark
                   else self._style.light_palette)
        bg, border, fg = [Color.FromArgb(*argb) for argb in palette]
        self._badge.Background = SolidColorBrush(bg)
        self._badge.BorderBrush = SolidColorBrush(border)
        self._text_block.Foreground = SolidColorBrush(fg)

    # ----------------------------------------------------------- lifecycle

    def start(self):
        """Start tracking. Call from a command (needs API context)."""
        if self._on_left_click is not None or self._on_right_click is not None:
            try:
                self._action_handler = _HudActionEventHandler()
                self._action_handler.driver = self
                self._action_event = ExternalEvent.Create(self._action_handler)
            except Exception:
                self._action_handler = None
                self._action_event = None
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
        # Idling is silent while Windows runs the modal move/size loop on
        # the Revit window, but WM_TIMER still gets dispatched there — so a
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
            self._badge = None
        if self._action_handler is not None:
            self._action_handler.driver = None
            self._action_handler = None
        if self._action_event is not None:
            try:
                self._action_event.Dispose()
            except Exception:
                pass
            self._action_event = None
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

    # -------------------------------------------------------------- events

    def _on_view_activated(self, sender, args):
        try:
            self._sync(force=True)
        except Exception:
            pass

    def _on_badge_left_click(self, sender, args):
        self._raise_action(self._on_left_click)

    def _on_badge_right_click(self, sender, args):
        self._raise_action(self._on_right_click)

    def _raise_action(self, action):
        if (action is None or self._action_event is None
                or self._action_handler is None):
            return
        self._action_handler.action = action
        try:
            self._action_event.Raise()
        except Exception:
            pass

    def _on_badge_mouse_enter(self, sender, args):
        try:
            self._window.Opacity = self._style.hover_opacity
        except Exception:
            pass

    def _on_badge_mouse_leave(self, sender, args):
        try:
            self._window.Opacity = self._style.idle_opacity
        except Exception:
            pass

    def _on_move_watch_tick(self, sender, args):
        """Hide while the Revit window is being moved or resized."""
        try:
            rect = window_rect(self._owner_hwnd_int)
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

    # ---------------------------------------------------------- placement

    def _move_to(self, rect):
        """Anchor the window over the view rect (device px -> DIP)."""
        scale_x = scale_y = 1.0
        try:
            source = PresentationSource.FromVisual(self._window)
            if source is not None:
                matrix = source.CompositionTarget.TransformFromDevice
                scale_x = matrix.M11
                scale_y = matrix.M22
        except Exception:
            pass
        style = self._style
        margin = style.margin_px
        width_dip = self._window.ActualWidth or 0.0
        anchor = style.anchor
        if anchor == 'top-left':
            left_dip = (rect.Left + margin) * scale_x
        elif anchor == 'top-right':
            left_dip = (rect.Right - margin) * scale_x - width_dip
        else:  # top-center
            center_px = rect.Left + (rect.Right - rect.Left) / 2.0
            left_dip = center_px * scale_x - width_dip / 2.0
        self._window.Left = left_dip
        self._window.Top = (rect.Top + margin) * scale_y

    def _sync(self, force=False):
        if self._window is None:
            return
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None or not uidoc.Document.Equals(self._doc):
            self._hide()
            return
        view = uidoc.ActiveView
        if view is None or view.IsTemplate:
            self._hide()
            return

        try:
            text = self._text_provider(self._doc, view)
        except Exception as ex:
            self._last_error = 'text_provider: {}'.format(ex)
            text = None
        if not text:
            self._hide()
            return

        uiview = get_uiview(self._uiapp, view.Id)
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
            if self._tooltip_provider is not None and self._badge is not None:
                try:
                    self._badge.ToolTip = self._tooltip_provider(text)
                except Exception:
                    pass
            self._apply_theme()
            self._last_view_rect = rect
            # settle layout so ActualWidth reflects the new text before
            # anchoring; before the first Show this is a no-op and the
            # SizeChanged handler does the anchoring instead
            try:
                self._window.UpdateLayout()
            except Exception:
                pass
            self._move_to(rect)
            if not self._visible:
                self._window.Show()
                self._visible = True
                self._move_to(rect)
            self._state_key = key
        except Exception as ex:
            self._last_error = 'show: {}'.format(ex)
            if self._logger is not None:
                try:
                    self._logger.warning(
                        "View HUD update failed: {}".format(ex))
                except Exception:
                    pass
