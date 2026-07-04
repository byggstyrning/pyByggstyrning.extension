# Phase HUD (prototype)

Toggle button that shows the active view's **Phase** name as a WPF overlay
badge pinned to the top-left corner of the viewport in any view with a Phase parameter (3D, plan, section, elevation, ...). This is
variant 2 from the survey in `Phase Label.pushbutton/README.md`, built for
side-by-side evaluation against the TemporaryGraphicsManager variant.

## How it works

- `lib/revit/phase_label_wpf.py` creates one borderless WPF window per
  document: no chrome, transparent background, owned by the Revit main
  window (so it never floats above other applications), and made
  click-through / non-activating via `WS_EX_TRANSPARENT | WS_EX_NOACTIVATE`.
- The window is positioned at `UIView.GetWindowRectangle()` top-left plus a
  margin, converting device pixels to WPF units through the window's
  `TransformFromDevice` matrix (per-monitor DPI).
- An `Idling` handler (throttled) plus `ViewActivated` keep text and
  position in sync: view switches, phase changes, Revit window move/resize.
- While the Revit window is being dragged or resized, `Idling` is silent
  (Windows runs a modal move/size loop), so a `DispatcherTimer` — whose
  `WM_TIMER` messages still get dispatched inside that loop — watches the
  main window rect via Win32 `GetWindowRect` and hides the badge as soon as
  the rect starts changing. The first `Idling` tick after release re-syncs
  and shows it at the new position.
- Because the anchor is in **screen pixels**, the badge does not move at all
  during pan/zoom/orbit — the key difference from the Phase Label button,
  whose model-anchored badge snaps back only after navigation ends.

## Known limitations (evaluation notes)

- During a Revit window move/resize the badge disappears (by design) and
  reappears ~0.3 s after release. Re-tiling views inside the window does
  not hide it — the badge just jumps to the new corner on the next idle
  tick, since the main window rect is unchanged and `UIView` coordinates
  cannot be read outside a Revit API context.
- The overlay can sit on top of Revit dialogs that open over the viewport
  corner (it is click-through, so it never blocks input).
- Does not print or export, and screen captures of the Revit window may or
  may not include it depending on the capture method.
- Crossing monitors with different DPI can be one tick late to rescale.
- Both phase buttons ON at once will stack two badges in the same corner —
  toggle one at a time when comparing.
