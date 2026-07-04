# Phase HUD (prototype)

Toggle button that shows the active view's **Phase** name as a WPF overlay
badge pinned to the top-left corner of the viewport in any view with a Phase parameter (3D, plan, section, elevation, ...). Chosen after side-by-side evaluation
against the TemporaryGraphicsManager variant, whose implementation is kept
in `lib/revit/phase_label.py` (see Alternatives surveyed below).

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

## Alternatives surveyed (2026-07)

All known Revit API variants for an in-viewport label, with why they were
or weren't chosen:

1. **TemporaryGraphicsManager / InCanvasControl** *(evaluated first; implementation kept in `lib/revit/phase_label.py`, ribbon button retired)* —
   Revit 2022+, bitmap badge, no model impact, no print. Model-anchored, so
   the badge snaps back to the corner only after pan/zoom ends.
2. **WPF overlay window** *(this button, `lib/revit/phase_label_wpf.py`)* — a borderless, click-through WPF window
   owned by the Revit main window, positioned over the viewport using
   `UIView.GetWindowRectangle()`. Pixel-anchored: stays in the corner *during*
   pan/zoom/orbit, full text rendering, no model impact, works pre-2022.
   Costs: must track window move/resize/DPI/multi-monitor and view switches;
   floats above dialogs if ownership is wrong; never prints. Strongest
   candidate if the TGM snap-back feels bad in evaluation.
3. **DirectContext3D server** — Revit 2017+ per-frame drawing, so tracking
   is smooth, but text must be tessellated into triangles (FormattedText /
   GraphicsPath), and corner-anchoring in ortho 3D still needs zoom state
   cached from UI events. Highest complexity by far; no print.
4. **Locked 3D view + TextNote/annotation** (`View3D.SaveOrientationAndLock`)
   — persistent and printable, but modifies the model, forces the orientation
   lock, and stays at a model point rather than the viewport corner.
5. **ModelText element** — a real 3D element; prints; but lives in model
   space (not screen-anchored) and pollutes the model.
6. **3D view background image** (`ViewDisplayBackground.CreateImage`,
   Revit 2014+) — render the phase text into an image and set it as the
   view background (flags/offset/scale relative to the view boundary).
   Per-view, persistent, prints; but requires a transaction, replaces any
   existing background, and draws behind model geometry.
7. **AVF legend hack** — Analysis Visualization Framework shows a legend with
   the display-style title in the view; some add-ins abuse it for status
   text. Requires dummy analysis data + transaction; legend placement is not
   controllable. Not recommended.
8. **Non-graphics fallbacks** — keep the phase in the view *name* (visible in
   the view tab/title bar, updated by a DocumentChanged updater), or show it
   in a dockable pane. Zero canvas footprint, zero risk.

The temporary-graphics approach was chosen first because it is zero-impact
on the model and reuses the proven marker driver pattern. If the snap-back
during navigation bothers users, try variant 2 (WPF overlay). If the label
must print, use variant 4 or 6.
