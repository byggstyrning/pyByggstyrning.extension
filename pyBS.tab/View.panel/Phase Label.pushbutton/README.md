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

## Alternative considered: graphics locked inside the view

A persistent variant would lock the 3D view (`View3D.SaveOrientationAndLock`)
and place a `TextNote` in the view — Revit allows annotations in locked 3D
views. That label would survive sessions, print, and export, but it:

- modifies the model (annotation element + view lock),
- forces the view orientation to stay locked (users can't orbit without
  unlocking and losing the annotation),
- does not follow pan/zoom to stay in the corner.

The temporary-graphics approach was chosen for evaluation because it is
zero-impact on the model and reuses the proven marker driver pattern. If a
printable label is needed later, the locked-view TextNote can be added as a
separate command.
