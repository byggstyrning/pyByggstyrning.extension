# -*- coding: utf-8 -*-
"""Phase label as a WPF overlay window pinned over the active view.

Works in any graphical view that has a Phase parameter (3D, plan, section,
elevation, ...).

Alternative to revit.phase_label (TemporaryGraphicsManager): a borderless,
topmost-within-Revit, non-activating WPF window centered over the top edge
of the active UIView's window rectangle. Because the anchor is in screen
pixels (UIView.GetWindowRectangle), the label stays glued in place *during*
pan/zoom/orbit — no snap-back. The muted palette follows Revit's light/dark
theme and the badge idles at 50% opacity, turning solid on hover. The badge is clickable:
left-click switches the view to the next project phase, right-click to the
previous one (via ExternalEvent + transaction). Beyond that parameter
change, nothing is written to the model and nothing prints. Works on any
Revit version with UIView (no TemporaryGraphicsManager requirement).
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

from Autodesk.Revit.DB import BuiltInParameter, Transaction
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

from revit.compat import get_element_id_value
from revit.phase_label import _label_text_for_view
from revit.view_markers import _normalize_view_id

_DRIVERS_SYS_KEY = '_pyBS_phase_hud_drivers'
_THROTTLE_MS = 300
_MARGIN_PX = 12
_MOVE_WATCH_MS = 100
_IDLE_OPACITY = 0.5
_HOVER_OPACITY = 1.0

_GWL_EXSTYLE = -20
_WS_EX_TRANSPARENT = 0x00000020
_WS_EX_TOOLWINDOW = 0x00000080
_WS_EX_NOACTIVATE = 0x08000000


def _apply_overlay_styles(hwnd_int):
    """Keep the overlay out of the taskbar/alt-tab and never steal focus.

    The badge stays clickable (no WS_EX_TRANSPARENT): clicks cycle the view
    phase, and WS_EX_NOACTIVATE keeps keyboard focus in Revit throughout.
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


def _revit_theme_is_dark():
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


def _ordered_phase_ids(document):
    """Project phases in chronological order (past -> future)."""
    ids = []
    try:
        for phase in document.Phases:
            ids.append(phase.Id)
    except Exception:
        pass
    return ids


def _shift_view_phase(uiapp, step):
    """Set the active view's Phase to the next/previous project phase.

    Must run inside a Revit API context (transaction). Returns the new
    phase name, or None when the phase cannot be changed.
    """
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return None
    doc = uidoc.Document
    view = uidoc.ActiveView
    if view is None:
        return None
    param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
    if param is None or param.IsReadOnly:
        return None
    phase_ids = _ordered_phase_ids(doc)
    if len(phase_ids) < 2:
        return None
    id_values = [_normalize_view_id(get_element_id_value(p))
                 for p in phase_ids]
    try:
        idx = id_values.index(
            _normalize_view_id(get_element_id_value(param.AsElementId())))
    except ValueError:
        idx = 0
    new_id = phase_ids[(idx + step) % len(phase_ids)]
    t = Transaction(doc, 'Switch view phase')
    t.Start()
    try:
        param.Set(new_id)
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass
        return None
    try:
        return doc.GetElement(new_id).Name
    except Exception:
        return None


class _PhaseCycleEventHandler(IExternalEventHandler):
    """Runs the phase switch inside a Revit API context."""

    def __init__(self):
        self.step = 1
        self.driver = None

    def Execute(self, uiapp):
        try:
            _shift_view_phase(uiapp, self.step)
            if self.driver is not None:
                self.driver.refresh()
        except Exception:
            pass

    def GetName(self):
        return "pyBS Phase HUD phase switch"


def collect_diagnostics(uiapp, document):
    """Step-by-step status report for troubleshooting a blank HUD."""
    lines = []

    def add(key, fn):
        try:
            lines.append('{}: {}'.format(key, fn()))
        except Exception as ex:
            lines.append('{}: ERROR {}'.format(key, ex))

    add('engine', lambda: sys.version.replace('\n', ' '))

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
    add('is_template', lambda: view.IsTemplate)
    add('phase_param', lambda: view.get_Parameter(
        BuiltInParameter.VIEW_PHASE) is not None)
    text = _label_text_for_view(document, view)
    lines.append(u'phase_text: {}'.format(text))

    uiview = _get_uiview(uiapp, view.Id)
    lines.append('uiview_found: {}'.format(uiview is not None))
    if uiview is not None:
        add('window_rect', lambda: str(uiview.GetWindowRectangle()))

    driver = find_phase_hud_driver(document)
    lines.append('driver_running: {}'.format(driver is not None))
    if driver is not None:
        lines.append('window_built: {}'.format(driver._window is not None))
        lines.append('window_visible_flag: {}'.format(driver._visible))
        if driver._window is not None:
            add('window_pos', lambda: '{}, {} ({}x{})'.format(
                driver._window.Left, driver._window.Top,
                driver._window.ActualWidth, driver._window.ActualHeight))
            add('window_is_visible', lambda: driver._window.IsVisible)
            add('window_text', lambda: driver._text_block.Text)
        lines.append('owner_hwnd: {}'.format(driver._owner_hwnd_int))
        add('owner_rect', lambda: str(_window_rect(driver._owner_hwnd_int)))
        add('theme_dark', _revit_theme_is_dark)
        lines.append('last_error: {}'.format(driver._last_error))
    return '\n'.join(lines)


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
    """Idling-driven WPF badge pinned to the active view's corner."""

    def __init__(self, uiapp, document, logger=None):
        self._uiapp = uiapp
        self._doc = document
        self._logger = logger
        self._window = None
        self._text_block = None
        self._badge = None
        self._theme_dark = None
        self._cycle_handler = None
        self._cycle_event = None
        self._visible = False
        self._state_key = None
        self._last_tick_ms = 0
        self._idling_handler = None
        self._view_activated_handler = None
        self._owner_hwnd_int = None
        self._last_owner_rect = None
        self._move_timer = None
        self._last_error = None

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
        window.Topmost = False
        window.Opacity = _IDLE_OPACITY

        text = TextBlock()
        text.FontFamily = FontFamily('Segoe UI')
        text.FontSize = 12.0

        badge = Border()
        badge.BorderThickness = Thickness(1)
        badge.CornerRadius = CornerRadius(4)
        badge.Padding = Thickness(8, 3, 8, 3)
        badge.Cursor = Cursors.Hand
        badge.Child = text
        badge.MouseLeftButtonDown += self._on_badge_left_click
        badge.MouseRightButtonDown += self._on_badge_right_click
        badge.MouseEnter += self._on_badge_mouse_enter
        badge.MouseLeave += self._on_badge_mouse_leave
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
            _apply_overlay_styles(hwnd.ToInt64())
        except Exception:
            pass

        self._window = window
        self._text_block = text
        self._badge = badge
        self._apply_theme()

    def _apply_theme(self):
        """Muted badge palette following Revit's light/dark theme."""
        dark = _revit_theme_is_dark()
        if dark == self._theme_dark or self._badge is None:
            return
        self._theme_dark = dark
        if dark:
            bg = Color.FromArgb(215, 43, 43, 43)
            border = Color.FromArgb(60, 255, 255, 255)
            fg = Color.FromArgb(255, 214, 214, 214)
        else:
            bg = Color.FromArgb(215, 246, 246, 246)
            border = Color.FromArgb(60, 0, 0, 0)
            fg = Color.FromArgb(255, 64, 64, 64)
        self._badge.Background = SolidColorBrush(bg)
        self._badge.BorderBrush = SolidColorBrush(border)
        self._text_block.Foreground = SolidColorBrush(fg)

    def start(self):
        # ExternalEvent.Create requires an API context; start() runs inside
        # the command that toggled the button, so this is the safe spot.
        try:
            self._cycle_handler = _PhaseCycleEventHandler()
            self._cycle_handler.driver = self
            self._cycle_event = ExternalEvent.Create(self._cycle_handler)
        except Exception:
            self._cycle_handler = None
            self._cycle_event = None
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
            self._badge = None
        if self._cycle_handler is not None:
            self._cycle_handler.driver = None
            self._cycle_handler = None
        if self._cycle_event is not None:
            try:
                self._cycle_event.Dispose()
            except Exception:
                pass
            self._cycle_event = None
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

    def _on_badge_left_click(self, sender, args):
        self._raise_phase_cycle(1)

    def _on_badge_right_click(self, sender, args):
        self._raise_phase_cycle(-1)

    def _on_badge_mouse_enter(self, sender, args):
        try:
            self._window.Opacity = _HOVER_OPACITY
        except Exception:
            pass

    def _on_badge_mouse_leave(self, sender, args):
        try:
            self._window.Opacity = _IDLE_OPACITY
        except Exception:
            pass

    def _raise_phase_cycle(self, step):
        if self._cycle_event is None or self._cycle_handler is None:
            return
        self._cycle_handler.step = step
        try:
            self._cycle_event.Raise()
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
        """Center the window over the view rect's top edge (device px -> DIP)."""
        scale_x = scale_y = 1.0
        try:
            source = PresentationSource.FromVisual(self._window)
            if source is not None:
                matrix = source.CompositionTarget.TransformFromDevice
                scale_x = matrix.M11
                scale_y = matrix.M22
        except Exception:
            pass
        center_px = rect.Left + (rect.Right - rect.Left) / 2.0
        width_dip = self._window.ActualWidth or 0.0
        self._window.Left = center_px * scale_x - width_dip / 2.0
        self._window.Top = (rect.Top + _MARGIN_PX) * scale_y

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
            if self._badge is not None:
                self._badge.ToolTip = (
                    u"Phase: {}\nClick: next phase. "
                    u"Right-click: previous phase.".format(text))
            self._apply_theme()
            # settle layout so ActualWidth reflects the new text before
            # horizontal centering
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
                        "Phase HUD update failed: {}".format(ex))
                except Exception:
                    pass
