# -*- coding: utf-8 -*-
"""Reusable in-view context switchers for Revit (WPF overlay HUD bar).

One ViewHudHost per document owns a horizontal bar of switcher *items*
pinned over the active view's window rectangle. The anchor is in screen
pixels (UIView.GetWindowRectangle) so the bar stays put during
pan/zoom/orbit. Multiple switchers coexist side by side, horizontally
separated; each shows/hides independently and the bar re-centers itself.

Host machinery (items get all of this for free):

- borderless, non-activating, owned WPF window (never steals Revit focus)
- muted palette following Revit's light/dark theme (UIThemeManager, 2024+)
- per-item idle/hover opacity
- per-monitor DPI-aware positioning with configurable anchor
- deferred sizing: the bar is parked off-screen until WPF completes the
  first layout pass (SizeChanged), avoiding the mis-sized first frame that
  EnsureHandle + SizeToContent produces
- hidden while the Revit main window is moved/resized (DispatcherTimer
  still ticks inside the modal move loop; Idling does not fire there)
- an ExternalEvent executor so item click actions run inside a Revit API
  context (transactions allowed)

Adding a simple switcher (a text badge, optionally clickable):

    from revit.view_hud import TextSwitcher, get_hud_host, find_hud_item

    item = find_hud_item(doc, 'my-switcher')
    if item is None:
        host = get_hud_host(uiapp, doc)          # created on demand
        host.add_item(TextSwitcher(
            'my-switcher',
            text_provider=lambda doc, view: 'text',   # None hides the item
            tooltip_provider=lambda text: 'hint',     # optional
            on_left_click=lambda uiapp: ...,          # optional, API context
            on_right_click=lambda uiapp: ...,         # optional, API context
        ))                                            # call from a command
    else:
        item.stop()                                   # toggle off

Adding new switcher *kinds* (planned: single/multi-select dropdowns for
temporary isolate, parameter colorization, ...): subclass HudItem —
build() returns any WPF element (e.g. a badge opening a Popup with a
ListBox or CheckBoxes), sync() returns an immutable value-equality
snapshot of its state (str/tuple; lists are defensively copied) or None
to hide, and run_in_api_context() executes model changes safely. Notes
for popup authors: a WPF Popup is its own top-level HWND — close or
reposition it in the on_bar_hidden / on_bar_moved / on_removed hooks;
the host window is non-activating (WS_EX_NOACTIVATE) and not focusable,
so light-dismiss (StaysOpen=False) may not see focus changes — prefer
StaysOpen=True with explicit close; and set hover_dim = False to keep
the badge solid while its popup is open.

Hosts and items are registered per document in an AppDomain slot (so the
registry survives pyRevit engine reloads), letting stateless toggle
buttons find them again: get_hud_host / find_hud_item / remove_hud_item.
The host starts its event plumbing when the first item is added and tears
down when the last one is removed or its document closes. add_item and
item click actions need a Revit API context (ExternalEvent.Create), so
add items from a command, not from arbitrary UI callbacks.
"""

import ctypes
import time

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

from System import TimeSpan, AppDomain
from System.Collections import ArrayList
from System.Windows import (
    Window, WindowStyle, ResizeMode, SizeToContent, Thickness,
    CornerRadius, PresentationSource, Visibility,
)
from System.Windows.Controls import Border, TextBlock, StackPanel, Orientation
from System.Windows.Input import Cursors
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Media import SolidColorBrush, Color, Brushes, FontFamily
from System.Windows.Threading import DispatcherTimer

from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

from revit.compat import get_element_id_value

# AppDomain-scoped so the registry survives pyRevit engine reloads: a new
# engine still finds (and can stop or reuse) hosts started by the old one
# instead of stacking a duplicate bar.
_HOSTS_DOMAIN_KEY = 'pyBS.view_hud.hosts'
_THROTTLE_MS = 300
_MOVE_WATCH_MS = 100
_OFFSCREEN = -32000

_GWL_EXSTYLE = -20
_WS_EX_TOOLWINDOW = 0x00000080
_WS_EX_NOACTIVATE = 0x08000000


class HudStyle(object):
    """Placement and look of the HUD bar; palettes are ARGB tuples."""

    def __init__(
        self,
        anchor='top-center',            # 'top-left' | 'top-center' | 'top-right'
        margin_px=12,
        item_spacing=6.0,               # horizontal gap between switchers (DIP)
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
        self.item_spacing = item_spacing
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

    Items stay clickable (no WS_EX_TRANSPARENT); WS_EX_NOACTIVATE keeps
    keyboard focus in Revit throughout.
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


# --------------------------------------------------------------- registry

def _hosts_registry():
    """Host list stored on the AppDomain (shared across pyRevit engines)."""
    domain = AppDomain.CurrentDomain
    hosts = domain.GetData(_HOSTS_DOMAIN_KEY)
    if hosts is None:
        hosts = ArrayList()
        domain.SetData(_HOSTS_DOMAIN_KEY, hosts)
    return hosts


def _doc_is_valid(document):
    try:
        return bool(document.IsValidObject)
    except Exception:
        return False


def find_hud_host(document):
    """Find the live host for a document, pruning hosts of closed docs."""
    hosts = _hosts_registry()
    for host in list(hosts):
        try:
            host_doc = host._doc
            if host_doc is None or not _doc_is_valid(host_doc):
                try:
                    host.stop()
                except Exception:
                    try:
                        hosts.Remove(host)
                    except Exception:
                        pass
                continue
            if host_doc.Equals(document):
                return host
        except Exception:
            continue
    return None


def get_hud_host(uiapp, document, style=None, logger=None):
    """Return the document's HUD host, creating (not starting) it if needed."""
    host = find_hud_host(document)
    if host is not None:
        return host
    host = ViewHudHost(uiapp, document, style=style, logger=logger)
    hosts = _hosts_registry()
    if not hosts.Contains(host):
        hosts.Add(host)
    return host


def find_hud_item(document, item_id):
    host = find_hud_host(document)
    if host is None:
        return None
    return host.find_item(item_id)


def remove_hud_item(document, item_id):
    """Remove one switcher; the host tears down when the last one goes."""
    host = find_hud_host(document)
    if host is None:
        return False
    return host.remove_item(item_id)


class _HudActionEventHandler(IExternalEventHandler):
    """Runs queued HUD click actions inside a Revit API context."""

    def __init__(self):
        self.pending = []
        self.host = None

    def Execute(self, uiapp):
        try:
            actions = self.pending
            self.pending = []
            for action in actions:
                try:
                    action(uiapp)
                except Exception:
                    pass
            if self.host is not None:
                self.host.refresh()
        except Exception:
            pass

    def GetName(self):
        return "pyBS View HUD action"


# ------------------------------------------------------------------ items

class HudItem(object):
    """One switcher cell in a ViewHudHost. Subclass to add new kinds.

    Contract:
    - build(style) -> FrameworkElement: create the cell UI once; called
      when the item is added to a host.
    - sync(document, view) -> state or None: refresh the cell from the
      model; return None to hide this cell (set your root's Visibility
      yourself). The returned state feeds the host's change detection
      (compared by ==), so include everything that affects rendering and
      return an immutable snapshot (str/tuple), never a live mutable
      object — the host stores the reference (lists are copied to tuples
      defensively).
    - apply_theme(dark, palette): recolor for Revit's light/dark theme;
      palette is (background, border, text) ARGB tuples from HudStyle.

    The host wires idle/hover opacity on the item root automatically; set
    hover_dim = False to manage opacity yourself (e.g. a dropdown that must
    stay solid while its popup is open). Lifecycle hooks (all optional):
    on_bar_hidden / on_bar_moved / on_removed — items owning extra top-level
    UI (a WPF Popup is its own HWND that neither hides nor moves with the
    bar) must close or reposition it in these.
    """

    hover_dim = True

    def __init__(self, item_id):
        self.item_id = item_id
        self.host = None
        self.root = None

    def build(self, style):
        raise NotImplementedError

    def sync(self, document, view):
        raise NotImplementedError

    def apply_theme(self, dark, palette):
        pass

    def on_bar_hidden(self):
        """Called when the host bar hides (view/doc switch, window move)."""
        pass

    def on_bar_moved(self):
        """Called after the host bar is repositioned."""
        pass

    def on_removed(self):
        """Called when the item leaves its host (removal or teardown)."""
        pass

    def refresh(self):
        """Force the host to re-sync all items."""
        if self.host is not None:
            self.host.refresh()

    def stop(self):
        """Remove this item from its host (tears the host down if last)."""
        if self.host is not None:
            self.host.remove_item(self.item_id)

    def run_in_api_context(self, action):
        """Queue action(uiapp) to run inside a Revit API context."""
        if self.host is not None:
            self.host.run_action(action)


class TextSwitcher(HudItem):
    """Text badge switcher: shows provider text, optional click actions.

    text_provider(document, view) supplies the text; return None to hide.
    on_left_click / on_right_click are optional callables taking uiapp,
    executed inside a Revit API context (transactions allowed).
    """

    def __init__(self, item_id, text_provider,
                 tooltip_provider=None,
                 on_left_click=None,
                 on_right_click=None):
        HudItem.__init__(self, item_id)
        self._text_provider = text_provider
        self._tooltip_provider = tooltip_provider
        self._on_left_click = on_left_click
        self._on_right_click = on_right_click
        self._style = None
        self._text_block = None
        self._last_text = None

    def build(self, style):
        self._style = style
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
        badge.MouseLeftButtonDown += self._on_mouse_left
        badge.MouseRightButtonDown += self._on_mouse_right

        self._text_block = text
        self._last_text = None
        self.root = badge
        return badge

    def sync(self, document, view):
        text = self._text_provider(document, view)
        if not text:
            self.root.Visibility = Visibility.Collapsed
            return None
        if text != self._last_text:
            self._text_block.Text = text
            self._last_text = text
            if self._tooltip_provider is not None:
                try:
                    self.root.ToolTip = self._tooltip_provider(text)
                except Exception:
                    pass
        self.root.Visibility = Visibility.Visible
        return text

    def apply_theme(self, dark, palette):
        bg, border, fg = [Color.FromArgb(*argb) for argb in palette]
        self.root.Background = SolidColorBrush(bg)
        self.root.BorderBrush = SolidColorBrush(border)
        self._text_block.Foreground = SolidColorBrush(fg)

    def _on_mouse_left(self, sender, args):
        if self._on_left_click is not None:
            self.run_in_api_context(self._on_left_click)

    def _on_mouse_right(self, sender, args):
        if self._on_right_click is not None:
            self.run_in_api_context(self._on_right_click)


# ------------------------------------------------------------------- host

class ViewHudHost(object):
    """Per-document HUD bar hosting the registered switcher items.

    Owns the overlay window and all plumbing (Idling sync, theme, DPI
    anchoring, hide-on-window-move, API-context click executor). Starts
    when the first item is added, tears down when the last one is removed.
    """

    def __init__(self, uiapp, document, style=None, logger=None):
        self._uiapp = uiapp
        self._doc = document
        self._style = style or DEFAULT_HUD_STYLE
        self._logger = logger
        self._items = []
        self._started = False
        self._window = None
        self._panel = None
        self._theme_dark = None
        self._action_handler = None
        self._action_event = None
        self._visible = False
        self._state_key = None
        self._last_view_rect = None
        self._last_tick_ms = 0
        self._idling_handler = None
        self._view_activated_handler = None
        self._doc_closing_handler = None
        self._owner_hwnd_int = None
        self._last_owner_rect = None
        self._move_timer = None
        self._last_error = None

    # ------------------------------------------------------------- items

    def find_item(self, item_id):
        for item in self._items:
            if item.item_id == item_id:
                return item
        return None

    def item_ids(self):
        return [item.item_id for item in self._items]

    def add_item(self, item, index=None):
        """Register a switcher (replacing any with the same id) and show it.

        Call from a command: the first add starts the host, which needs a
        Revit API context (ExternalEvent.Create).
        """
        # replacing must not tear the host down when the old item was the
        # last one — a full stop would unregister the host mid-add
        self.remove_item(item.item_id, _teardown_if_empty=False)
        if not self._started:
            self._start()
        item.host = self
        root = item.build(self._style)
        half = self._style.item_spacing / 2.0
        root.Margin = Thickness(half, 0, half, 0)
        if getattr(item, 'hover_dim', True):
            root.Opacity = self._style.idle_opacity
            root.MouseEnter += self._on_item_mouse_enter
            root.MouseLeave += self._on_item_mouse_leave
        if index is None or index >= len(self._items):
            self._items.append(item)
            self._panel.Children.Add(root)
        else:
            self._items.insert(index, item)
            self._panel.Children.Insert(index, root)
        try:
            item.apply_theme(
                self._theme_dark,
                self._style.dark_palette if self._theme_dark
                else self._style.light_palette)
        except Exception:
            pass
        self.refresh()
        return item

    def remove_item(self, item_id, _teardown_if_empty=True):
        item = self.find_item(item_id)
        if item is None:
            return False
        self._items.remove(item)
        self._detach_item_ui(item)
        try:
            item.on_removed()
        except Exception:
            pass
        item.host = None
        if not self._items and _teardown_if_empty:
            self._stop()
        else:
            self.refresh()
        return True

    def _detach_item_ui(self, item):
        if item.root is not None:
            if getattr(item, 'hover_dim', True):
                try:
                    item.root.MouseEnter -= self._on_item_mouse_enter
                    item.root.MouseLeave -= self._on_item_mouse_leave
                except Exception:
                    pass
            if self._panel is not None:
                try:
                    self._panel.Children.Remove(item.root)
                except Exception:
                    pass

    def _on_item_mouse_enter(self, sender, args):
        try:
            sender.Opacity = self._style.hover_opacity
        except Exception:
            pass

    def _on_item_mouse_leave(self, sender, args):
        try:
            sender.Opacity = self._style.idle_opacity
        except Exception:
            pass

    def run_action(self, action):
        """Queue action(uiapp) to run inside a Revit API context."""
        if action is None or self._action_event is None \
                or self._action_handler is None:
            return
        self._action_handler.pending.append(action)
        try:
            self._action_event.Raise()
        except Exception:
            pass

    # ---------------------------------------------------------------- ui

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
        # park off-screen: the first layout pass happens after Show(), so
        # any mis-sized first frame is invisible; SizeChanged then anchors
        # the settled window at the real position
        window.Left = _OFFSCREEN
        window.Top = _OFFSCREEN

        panel = StackPanel()
        panel.Orientation = Orientation.Horizontal
        window.Content = panel

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
        self._panel = panel

    def _on_source_initialized(self, sender, args):
        try:
            hwnd = WindowInteropHelper(self._window).Handle
            _apply_overlay_styles(hwnd.ToInt64())
        except Exception:
            pass

    def _on_window_size_changed(self, sender, args):
        # fires when the first layout pass (and any item/text change)
        # resizes the bar; re-anchor using the last known view rect
        if self._last_view_rect is not None:
            try:
                self._move_to(self._last_view_rect)
            except Exception:
                pass

    def _apply_theme(self):
        """Recolor all items when Revit's light/dark theme flips."""
        dark = revit_theme_is_dark()
        if dark == self._theme_dark:
            return
        self._theme_dark = dark
        palette = (self._style.dark_palette if dark
                   else self._style.light_palette)
        for item in self._items:
            try:
                item.apply_theme(dark, palette)
            except Exception:
                pass

    # --------------------------------------------------------- lifecycle

    def _start(self):
        """Start tracking. Runs inside the command that adds the first item."""
        # a host can be restarted after a full teardown (e.g. via a stale
        # reference) — make sure it is findable in the registry again
        hosts = _hosts_registry()
        if not hosts.Contains(self):
            hosts.Add(self)
        try:
            self._action_handler = _HudActionEventHandler()
            self._action_handler.host = self
            self._action_event = ExternalEvent.Create(self._action_handler)
        except Exception:
            self._action_handler = None
            self._action_event = None
        self._build_window()
        self._theme_dark = None
        self._apply_theme()
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
        # DispatcherTimer can hide the bar the moment the window starts
        # moving; the first Idling tick after release shows it again.
        try:
            self._move_timer = DispatcherTimer()
            self._move_timer.Interval = TimeSpan.FromMilliseconds(
                _MOVE_WATCH_MS)
            self._move_timer.Tick += self._on_move_watch_tick
            self._move_timer.Start()
        except Exception:
            self._move_timer = None
        # tear down with the document — otherwise the host leaks its
        # Idling handler and timer for the rest of the session
        self._doc_closing_handler = self._on_document_closing
        try:
            self._uiapp.Application.DocumentClosing += \
                self._doc_closing_handler
        except Exception:
            self._doc_closing_handler = None
        self._started = True

    def _stop(self):
        if self._doc_closing_handler is not None:
            try:
                self._uiapp.Application.DocumentClosing -= \
                    self._doc_closing_handler
            except Exception:
                pass
            self._doc_closing_handler = None
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
            self._panel = None
        if self._action_handler is not None:
            self._action_handler.host = None
            self._action_handler = None
        if self._action_event is not None:
            try:
                self._action_event.Dispose()
            except Exception:
                pass
            self._action_event = None
        for item in self._items:
            try:
                item.on_removed()
            except Exception:
                pass
            item.host = None
        self._items = []
        self._visible = False
        self._started = False
        try:
            _hosts_registry().Remove(self)
        except Exception:
            pass

    def stop(self):
        """Tear down the whole bar regardless of registered items."""
        self._stop()

    def refresh(self):
        self._state_key = None
        self._sync(force=True)

    def _hide(self):
        was_visible = self._visible
        if self._window is not None and self._visible:
            try:
                self._window.Hide()
            except Exception:
                pass
        self._visible = False
        self._state_key = None
        if was_visible:
            for item in self._items:
                try:
                    item.on_bar_hidden()
                except Exception:
                    pass

    # ------------------------------------------------------------ events

    def _on_view_activated(self, sender, args):
        try:
            self._sync(force=True)
        except Exception:
            pass

    def _on_document_closing(self, sender, args):
        try:
            if args.Document.Equals(self._doc):
                self.stop()
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

    # -------------------------------------------------------- placement

    def _move_to(self, rect):
        """Anchor the bar over the view rect (device px -> DIP)."""
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
        for item in self._items:
            try:
                item.on_bar_moved()
            except Exception:
                pass

    def _sync(self, force=False):
        if self._window is None:
            return
        if self._doc is None or not _doc_is_valid(self._doc):
            # document closed under us (DocumentClosing missed or blocked):
            # full teardown, not just hide, or the handlers leak all session
            self.stop()
            return
        uidoc = self._uiapp.ActiveUIDocument
        if uidoc is None or not uidoc.Document.Equals(self._doc):
            self._hide()
            return
        view = uidoc.ActiveView
        if view is None or view.IsTemplate:
            self._hide()
            return

        item_states = []
        any_visible = False
        for item in self._items:
            try:
                state = item.sync(self._doc, view)
            except Exception as ex:
                self._last_error = 'sync[{}]: {}'.format(item.item_id, ex)
                state = None
                try:
                    if item.root is not None:
                        item.root.Visibility = Visibility.Collapsed
                except Exception:
                    pass
            if isinstance(state, list):
                # snapshot: a live list stored by reference would alias the
                # item's own state and defeat change detection
                state = tuple(state)
            item_states.append((item.item_id, state))
            if state is not None:
                any_visible = True
        if not any_visible:
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

        # theme is part of the key so a light/dark flip on an otherwise
        # unchanged view still recolors the bar
        key = (rect.Left, rect.Top, rect.Right, rect.Bottom,
               revit_theme_is_dark(), tuple(item_states))
        if not force and self._visible and key == self._state_key:
            return

        try:
            self._apply_theme()
            self._last_view_rect = rect
            # settle layout so ActualWidth reflects the new content before
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
