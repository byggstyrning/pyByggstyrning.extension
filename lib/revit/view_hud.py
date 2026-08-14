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
- the whole bar fades out/in on view switch so chips do not collapse
  one-by-one
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

For a single-select dropdown badge use the built-in DropdownSwitcher:

        host.add_item(DropdownSwitcher(
            'my-picker',
            options_provider=lambda doc, view: (key1, 'Current', [
                (key1, 'Current'), (key2, 'Other')]),  # None hides the item
            on_select=lambda uiapp, key: ...,          # API context
            tooltip_provider=lambda label: 'hint',     # optional
        ))

Pass on_left_click / on_right_click to split the badge: the label cycles
on click, the caret still opens the dropdown (see the phase switcher).

Adding further switcher *kinds* (e.g. multi-select / checkbox popups for
temporary isolate, parameter colorization, ...): subclass HudItem —
build() returns any WPF element, sync() returns an immutable value-
equality snapshot of its state (str/tuple; lists are defensively copied)
or None to hide, and run_in_api_context() executes model changes safely.
Notes for popup authors (DropdownSwitcher already handles these): a WPF
Popup is its own top-level HWND — close it in the on_bar_hidden /
on_bar_moved / on_removed hooks; StaysOpen=False gives click-outside
dismissal; and set hover_dim = False to keep the badge solid (self-manage
opacity) while its popup is open.

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
    CornerRadius, PresentationSource, Visibility, Duration,
)
from System.Windows.Controls import (
    Border, TextBlock, StackPanel, Orientation, ListBox, ListBoxItem,
)
from System.Windows.Controls.Primitives import Popup, PlacementMode
from System.Windows.Input import Cursors
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Media import SolidColorBrush, Color, Brushes, FontFamily
from System.Windows.Media.Animation import DoubleAnimation, FillBehavior
from System.Windows.Threading import DispatcherTimer

from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

from revit.compat import get_element_id_value

# AppDomain-scoped so the registry survives pyRevit engine reloads: a new
# engine still finds (and can stop or reuse) hosts started by the old one
# instead of stacking a duplicate bar.
_HOSTS_DOMAIN_KEY = 'pyBS.view_hud.hosts'
_THROTTLE_MS = 80
_MOVE_WATCH_MS = 100
_OFFSCREEN = -32000
_FADE_IN_MS = 70

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
            # After an extension reload the AppDomain still holds the old
            # class instance. Drop it so callers get a current ViewHudHost.
            if type(host) is not ViewHudHost:
                try:
                    host.stop()
                except Exception:
                    try:
                        hosts.Remove(host)
                    except Exception:
                        pass
                continue
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
        self._flash_timer = None

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
        timer = self._flash_timer
        if timer is None:
            return
        try:
            timer.Stop()
        except Exception:
            pass
        self._flash_timer = None

    def close_popup(self):
        """Close any popup this item owns (no-op unless the item has one)."""
        pass

    def highlight_briefly(self, duration_ms=2200):
        """Flash this cell gold (e.g. point at the locking scope box)."""
        root = self.root
        if root is None:
            return
        try:
            timer = self._flash_timer
            if timer is not None:
                try:
                    timer.Stop()
                except Exception:
                    pass
            old_opacity = 1.0
            try:
                old_opacity = float(root.Opacity)
            except Exception:
                pass
            gold = Color.FromArgb(255, 255, 200, 0)
            ink = Color.FromArgb(255, 30, 30, 30)
            root.Opacity = 1.0
            root.Background = SolidColorBrush(gold)
            root.BorderBrush = SolidColorBrush(gold)
            # Keep BorderThickness unchanged. Growing/shrinking the chip
            # fires SizeChanged → _move_to → on_bar_moved → popup close.
            for attr in ('_label', '_caret', '_text_block'):
                child = getattr(self, attr, None)
                if child is not None:
                    try:
                        child.Foreground = SolidColorBrush(ink)
                    except Exception:
                        pass
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromMilliseconds(duration_ms)
            item = self

            def _restore(sender, args):
                try:
                    timer.Stop()
                except Exception:
                    pass
                item._flash_timer = None
                popup_open = False
                try:
                    popup = getattr(item, '_popup', None)
                    popup_open = popup is not None and bool(popup.IsOpen)
                except Exception:
                    popup_open = False
                try:
                    if not popup_open:
                        root.Opacity = old_opacity
                except Exception:
                    pass
                host = item.host
                if host is None or host._style is None:
                    return
                try:
                    palette = (
                        host._style.dark_palette if host._theme_dark
                        else host._style.light_palette)
                    item.apply_theme(host._theme_dark, palette)
                except Exception:
                    pass
                if popup_open:
                    try:
                        item._update_opacity()
                    except Exception:
                        pass

            timer.Tick += _restore
            self._flash_timer = timer
            timer.Start()
        except Exception:
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


class DropdownSwitcher(HudItem):
    """Badge that opens a single-select dropdown popup.

    options_provider(document, view) -> (current_key, current_label,
    options) or None to hide the cell. ``options`` is a list of
    (key, label) pairs; ``key`` is any hashable id (matched by value, so
    display-label collisions never mis-select). ``current_key`` marks the
    active option, ``current_label`` is shown on the badge. on_select(
    uiapp, key) runs inside a Revit API context when a different item is
    picked.

    The badge shows the current label plus a caret. Clicking it opens a
    themed ListBox popup; picking an item closes the popup and (if the key
    changed) queues on_select. Clicking elsewhere dismisses the popup
    (StaysOpen=False); the popup is also closed when the bar hides, moves,
    or the item is removed. Opacity is self-managed so the badge stays
    solid while its popup is open.

    Optional on_left_click / on_right_click split the badge: the label
    cycles on click, the caret still opens the dropdown. Use
    caret_tooltip_provider for a distinct hover on the caret.
    """

    hover_dim = False

    def __init__(self, item_id, options_provider, on_select,
                 tooltip_provider=None,
                 on_left_click=None,
                 on_right_click=None,
                 caret_tooltip_provider=None):
        HudItem.__init__(self, item_id)
        self._options_provider = options_provider
        self._on_select = on_select
        self._tooltip_provider = tooltip_provider
        self._caret_tooltip_provider = caret_tooltip_provider
        self._on_left_click = on_left_click
        self._on_right_click = on_right_click
        self._split_clicks = (
            on_left_click is not None or on_right_click is not None)
        self._style = None
        self._label = None
        self._popup = None
        self._listbox = None
        self._current_key = None
        self._options = []
        self._last_state = None
        self._is_hovered = False
        self._suppress_selection = False
        self._popup_closed_ms = 0
        self._caret_hit = None
        self._split_rule = None

    def build(self, style):
        self._style = style
        label = TextBlock()
        label.FontFamily = FontFamily(style.font_family)
        label.FontSize = style.font_size
        caret = TextBlock()
        caret.FontFamily = FontFamily(style.font_family)
        caret.FontSize = style.font_size
        caret.Text = u'▾'

        caret_hit = Border()
        caret_hit.Background = Brushes.Transparent
        caret_hit.Padding = Thickness(6, 0, 0, 0)
        caret_hit.Child = caret
        caret_hit.Cursor = Cursors.Hand

        split_rule = Border()
        split_rule.Width = 1
        split_rule.Margin = Thickness(6, 2, 2, 2)
        split_rule.Visibility = (
            Visibility.Visible if self._split_clicks
            else Visibility.Collapsed)

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Children.Add(label)
        row.Children.Add(split_rule)
        row.Children.Add(caret_hit)

        badge = Border()
        badge.BorderThickness = Thickness(1)
        badge.CornerRadius = CornerRadius(style.corner_radius)
        pad = style.padding
        badge.Padding = Thickness(pad[0], pad[1], pad[2], pad[3])
        badge.Child = row
        badge.Cursor = Cursors.Hand
        badge.Opacity = style.idle_opacity
        # open on button-UP, not down: setting Popup.IsOpen while the mouse
        # button is held puts a StaysOpen=False popup into native drag-select
        # mode (it closes on button-up unless you drag into an item)
        if self._split_clicks:
            label.Cursor = Cursors.Hand
            label.MouseLeftButtonDown += self._on_label_left
            label.MouseRightButtonDown += self._on_label_right
            caret_hit.MouseLeftButtonUp += self._on_badge_click
        else:
            badge.MouseLeftButtonUp += self._on_badge_click
        badge.MouseEnter += self._on_mouse_enter
        badge.MouseLeave += self._on_mouse_leave

        listbox = ListBox()
        listbox.FontFamily = FontFamily(style.font_family)
        listbox.FontSize = style.font_size
        listbox.BorderThickness = Thickness(0)
        listbox.SelectionChanged += self._on_selection_changed

        popup_border = Border()
        popup_border.BorderThickness = Thickness(1)
        popup_border.CornerRadius = CornerRadius(style.corner_radius)
        popup_border.Child = listbox

        popup = Popup()
        popup.StaysOpen = False
        popup.AllowsTransparency = True
        popup.PlacementTarget = badge
        popup.Placement = PlacementMode.Bottom
        popup.Child = popup_border
        popup.Opened += self._on_popup_opened
        popup.Closed += self._on_popup_closed

        self._label = label
        self._caret = caret
        self._caret_hit = caret_hit
        self._split_rule = split_rule
        self._listbox = listbox
        self._popup = popup
        self._popup_border = popup_border
        self._last_state = None
        self.root = badge
        return badge

    def sync(self, document, view):
        result = self._options_provider(document, view)
        if not result:
            self.root.Visibility = Visibility.Collapsed
            self._close_popup()
            return None
        current_key, current_label, options = result
        self._current_key = current_key
        self._options = list(options or [])
        self._label.Text = current_label
        many = len(self._options) > 1
        caret_vis = Visibility.Visible if many else Visibility.Collapsed
        self._caret.Visibility = caret_vis
        if self._caret_hit is not None:
            self._caret_hit.Visibility = caret_vis
        if self._split_rule is not None:
            self._split_rule.Visibility = (
                Visibility.Visible if many and self._split_clicks
                else Visibility.Collapsed)
        self._apply_tooltips(current_label)
        self.root.Visibility = Visibility.Visible
        # value-equality snapshot for host change detection
        return (current_key, current_label, tuple(self._options))

    def apply_theme(self, dark, palette):
        bg, border, fg = [Color.FromArgb(*argb) for argb in palette]
        bg_brush = SolidColorBrush(bg)
        border_brush = SolidColorBrush(border)
        fg_brush = SolidColorBrush(fg)
        self.root.Background = bg_brush
        self.root.BorderBrush = border_brush
        self._label.Foreground = fg_brush
        self._caret.Foreground = fg_brush
        if self._split_rule is not None:
            self._split_rule.Background = border_brush
        if self._popup_border is not None:
            self._popup_border.Background = bg_brush
            self._popup_border.BorderBrush = border_brush
        if self._listbox is not None:
            self._listbox.Background = bg_brush
            self._listbox.Foreground = fg_brush

    def _apply_tooltips(self, current_label):
        try:
            if self._split_clicks:
                if self._tooltip_provider is not None:
                    self._label.ToolTip = self._tooltip_provider(current_label)
                if self._caret_tooltip_provider is not None:
                    tip = self._caret_tooltip_provider(current_label)
                    self._caret_hit.ToolTip = tip
                    self._caret.ToolTip = tip
                self.root.ToolTip = None
            elif self._tooltip_provider is not None:
                self.root.ToolTip = self._tooltip_provider(current_label)
        except Exception:
            pass

    def _on_label_left(self, sender, args):
        if self._on_left_click is not None:
            args.Handled = True
            self.run_in_api_context(self._on_left_click)

    def _on_label_right(self, sender, args):
        if self._on_right_click is not None:
            args.Handled = True
            self.run_in_api_context(self._on_right_click)

    # -------------------------------------------------------------- popup

    def _populate_listbox(self):
        self._suppress_selection = True
        try:
            self._listbox.Items.Clear()
            selected = None
            for key, lbl in self._options:
                item = ListBoxItem()
                item.Content = lbl
                item.Tag = key
                self._listbox.Items.Add(item)
                if key == self._current_key:
                    selected = item
            self._listbox.SelectedItem = selected
        finally:
            self._suppress_selection = False

    def _on_badge_click(self, sender, args):
        if self._popup is None:
            return
        if self._popup.IsOpen:
            self._close_popup()
            return
        # if this same click just dismissed the popup (StaysOpen=False
        # closes it on the button-down that precedes this up), don't reopen
        # — the badge acts as a close toggle
        if int(time.time() * 1000) - self._popup_closed_ms < 300:
            return
        if len(self._options) <= 1:
            return
        # one popup open at a time across the bar
        if self.host is not None:
            self.host.close_item_popups(except_item=self)
        self._populate_listbox()
        try:
            self._popup.IsOpen = True
        except Exception:
            pass

    def _on_selection_changed(self, sender, args):
        if self._suppress_selection:
            return
        item = self._listbox.SelectedItem
        self._close_popup()
        if item is None:
            return
        key = item.Tag
        if key == self._current_key:
            return
        # optimistic: reflect the pick now so a fast reopen before the
        # ExternalEvent runs shows the right selection and re-picking the
        # same item is a no-op
        self._current_key = key
        on_select = self._on_select
        self.run_in_api_context(lambda ua: on_select(ua, key))

    def close_popup(self):
        self._close_popup()

    def _close_popup(self):
        if self._popup is not None and self._popup.IsOpen:
            try:
                self._popup.IsOpen = False
            except Exception:
                pass

    def _update_opacity(self):
        try:
            solid = self._is_hovered or (
                self._popup is not None and self._popup.IsOpen)
            self.root.Opacity = (
                self._style.hover_opacity if solid
                else self._style.idle_opacity)
        except Exception:
            pass

    def _on_mouse_enter(self, sender, args):
        self._is_hovered = True
        self._update_opacity()

    def _on_mouse_leave(self, sender, args):
        self._is_hovered = False
        self._update_opacity()

    def _on_popup_opened(self, sender, args):
        self._update_opacity()

    def _on_popup_closed(self, sender, args):
        self._popup_closed_ms = int(time.time() * 1000)
        self._update_opacity()

    # ----------------------------------------------------- lifecycle hooks

    def on_bar_hidden(self):
        self._close_popup()

    def on_bar_moved(self):
        self._close_popup()

    def on_removed(self):
        HudItem.on_removed(self)
        self._close_popup()


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
        self._last_view_id = None
        self._fading = False
        self._fade_resync = False
        self._fade_anim = None
        self._fade_gen = 0
        self._paused = False

    # ------------------------------------------------------------- items

    def find_item(self, item_id):
        for item in self._items:
            if item.item_id == item_id:
                return item
        return None

    def item_ids(self):
        return [item.item_id for item in self._items]

    def add_item(self, item, index=None, refresh=True):
        """Register a switcher (replacing any with the same id) and show it.

        Call from a command: the first add starts the host, which needs a
        Revit API context (ExternalEvent.Create). Pass refresh=False when
        adding several items, then call refresh() once.
        """
        # replacing must not tear the host down when the old item was the
        # last one — a full stop would unregister the host mid-add
        self.remove_item(item.item_id, _teardown_if_empty=False, refresh=False)
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
        if refresh:
            self.refresh()
        return item

    def remove_item(self, item_id, _teardown_if_empty=True, refresh=True):
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
        elif refresh:
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

    def close_item_popups(self, except_item=None):
        """Close every item's popup except one (single-open-popup bar)."""
        for item in self._items:
            if item is not except_item:
                try:
                    item.close_popup()
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

    def _apply_theme(self, dark=None):
        """Recolor all items when Revit's light/dark theme flips."""
        if dark is None:
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

    def _hook_live_events(self):
        """Idling / ViewActivated / move-watch. Cheap to pause/resume."""
        if self._idling_handler is None:
            self._idling_handler = self._on_idling
            try:
                self._uiapp.Idling += self._idling_handler
            except Exception:
                self._idling_handler = None
        if self._view_activated_handler is None:
            self._view_activated_handler = self._on_view_activated
            try:
                self._uiapp.ViewActivated += self._view_activated_handler
            except Exception:
                self._view_activated_handler = None
        if self._move_timer is None:
            try:
                self._move_timer = DispatcherTimer()
                self._move_timer.Interval = TimeSpan.FromMilliseconds(
                    _MOVE_WATCH_MS)
                self._move_timer.Tick += self._on_move_watch_tick
                self._move_timer.Start()
            except Exception:
                self._move_timer = None

    def _unhook_live_events(self):
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

    def is_paused(self):
        return bool(self._paused)

    def pause(self):
        """Toggle off: hide the bar, stop live tracking, keep the WPF tree."""
        if not self._started or self._paused:
            self._hide()
            self._paused = True
            return
        self._paused = True
        self._cancel_fade()
        self._unhook_live_events()
        self._hide()

    def resume(self):
        """Toggle on: rehook events and show the existing bar."""
        if not self._started:
            self._start()
        self._paused = False
        self._last_view_id = None
        self._last_tick_ms = 0
        self._hook_live_events()
        self._bake_opacity(1.0)
        self.refresh()

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
        self._hook_live_events()
        # tear down with the document — otherwise the host leaks its
        # Idling handler and timer for the rest of the session
        self._doc_closing_handler = self._on_document_closing
        try:
            self._uiapp.Application.DocumentClosing += \
                self._doc_closing_handler
        except Exception:
            self._doc_closing_handler = None
        self._paused = False
        self._started = True

    def _stop(self):
        self._paused = False
        self._cancel_fade()
        if self._doc_closing_handler is not None:
            try:
                self._uiapp.Application.DocumentClosing -= \
                    self._doc_closing_handler
            except Exception:
                pass
            self._doc_closing_handler = None
        self._unhook_live_events()
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
        if not self._fading:
            self._cancel_fade()
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

    # -------------------------------------------------------------- fade

    def _current_view_id(self):
        try:
            uidoc = self._uiapp.ActiveUIDocument
            if uidoc is None:
                return None
            view = uidoc.ActiveView
            if view is None:
                return None
            return get_element_id_value(view.Id)
        except Exception:
            return None

    def _bake_opacity(self, value):
        """Drop any opacity animation and pin the current value."""
        window = self._window
        if window is None:
            return
        try:
            window.BeginAnimation(Window.OpacityProperty, None)
        except Exception:
            pass
        try:
            window.Opacity = float(value)
        except Exception:
            pass
        self._fade_anim = None

    def _cancel_fade(self):
        self._fade_gen += 1
        self._fading = False
        self._fade_resync = False
        self._bake_opacity(1.0)

    def _animate_opacity(self, target, duration_ms, on_done=None):
        window = self._window
        if window is None:
            self._fading = False
            if on_done is not None:
                on_done()
            return
        try:
            anim = DoubleAnimation()
            anim.To = float(target)
            anim.Duration = Duration(TimeSpan.FromMilliseconds(duration_ms))
            anim.FillBehavior = FillBehavior.HoldEnd
            gen = self._fade_gen

            def _done(sender, args):
                try:
                    anim.Completed -= _done
                except Exception:
                    pass
                if gen != self._fade_gen:
                    return
                if on_done is not None:
                    on_done()

            if on_done is not None:
                anim.Completed += _done
            window.BeginAnimation(Window.OpacityProperty, anim)
            self._fade_anim = anim
        except Exception:
            self._bake_opacity(target)
            if on_done is not None:
                on_done()

    def _finish_fade(self):
        self._bake_opacity(1.0)
        self._fading = False
        if self._fade_resync:
            self._fade_resync = False
            if self._current_view_id() != self._last_view_id:
                self._sync(force=True)

    def _crossfade(self):
        """Fade the whole bar out, rebuild chips, fade in.

        Item.sync() collapses cells as it goes; doing that while the bar
        is visible makes chips vanish one-by-one. Fade first so the old
        bar leaves as a unit.
        """
        self._fading = True
        self._fade_resync = False
        self.close_item_popups()
        for item in self._items:
            try:
                item.on_bar_hidden()
            except Exception:
                pass

        def _after_out():
            if self._window is None:
                self._fading = False
                return
            try:
                self._last_view_id = self._current_view_id()
                self._relayout(force=True)
            except Exception:
                pass
            if self._window is None:
                self._fading = False
                return
            if not self._visible:
                self._bake_opacity(1.0)
                self._fading = False
                if self._fade_resync:
                    self._fade_resync = False
                    self._sync(force=True)
                return
            self._animate_opacity(1.0, _FADE_IN_MS, self._finish_fade)

        # Hide immediately, then fade in. A 50ms fade-out animation's
        # Completed callback arrived at 85–170ms in practice.
        self._bake_opacity(0.0)
        _after_out()

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
                self._cancel_fade()
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
        new_top = (rect.Top + margin) * scale_y
        old_left = self._window.Left
        old_top = self._window.Top
        self._window.Left = left_dip
        self._window.Top = new_top
        moved = (abs(float(old_left) - left_dip) > 0.5
                 or abs(float(old_top) - new_top) > 0.5)
        if not moved:
            return
        for item in self._items:
            try:
                item.on_bar_moved()
            except Exception:
                pass

    def _sync(self, force=False):
        if self._window is None or self._paused:
            return
        if self._fading:
            self._fade_resync = True
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

        view_id = self._current_view_id()
        if (self._visible
                and self._last_view_id is not None
                and view_id != self._last_view_id):
            self._crossfade()
            return
        self._last_view_id = view_id
        self._relayout(force)

    def _relayout(self, force=False):
        """Rebuild chip visibility and pin the bar. Caller handles fade."""
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

        dark = revit_theme_is_dark()
        key = (rect.Left, rect.Top, rect.Right, rect.Bottom,
               dark, tuple(item_states))
        if not force and self._visible and key == self._state_key:
            return

        try:
            self._apply_theme(dark)
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
