# Phase Label (prototype)

Toggle button that pins the active view's **Phase** name as a badge in the
top-left corner of the viewport in 3D views.

## How it works

- Uses `TemporaryGraphicsManager` (Revit 2022+) `InCanvasControl`s — the same
  zoom-aware infrastructure as the clash Markers button
  (`lib/revit/view_markers.py`). Nothing is written to the model and the
  badge never prints or exports.
- An `InCanvasControl` is anchored at a fixed model-space point, so
  `lib/revit/phase_label.py` recomputes the anchor from
  `UIView.GetZoomCorners()` on `Idling` (throttled) and `ViewActivated`:
  the visible top-left corner is derived in the view's
  Right/Up/ViewDirection basis, then inset by the badge's pixel size
  converted to model units via `UIView.GetWindowRectangle()`.
- The badge bitmap (phase name text) is rendered with System.Drawing and
  cached per phase name under `%LOCALAPPDATA%\pyBS\pyBS_phase_labels`.
- The badge redraws when the phase parameter, zoom, window size, or active
  view changes. Views without a Phase parameter show no badge.

## Known limitations (evaluation notes)

- During a pan/zoom drag the badge moves with the model; it snaps back to
  the corner when Revit idles (~0.3 s after the gesture ends). This is
  inherent to model-anchored temporary graphics.
- The badge appears only in the active 3D view while the toggle is on.
- Revit 2022+ only (`TemporaryGraphicsManager`).

## Alternatives surveyed (2026-07)

All known Revit API variants for an in-viewport label, with why they were
or weren't chosen:

1. **TemporaryGraphicsManager / InCanvasControl** *(implemented here)* —
   Revit 2022+, bitmap badge, no model impact, no print. Model-anchored, so
   the badge snaps back to the corner only after pan/zoom ends.
2. **WPF overlay window** *(implemented as the Phase HUD button,
   `lib/revit/phase_label_wpf.py`)* — a borderless, click-through WPF window
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
