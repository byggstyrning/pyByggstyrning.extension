# View HUD

Ribbon panel (`View HUD.panel`) with **Toggle** (on/off overlay). Overlay
is a WPF bar pinned over the viewport. A **⚙** badge on the right of the
bar opens pyRevit `CommandSwitchWindow` chip toggles (same pattern as
pyRevit Pick).

## Ribbon

| Button | Role |
| --- | --- |
| Toggle | Show / hide the HUD. Shift+Click: refresh + diagnostics. |

Chip on/off and auto-start live in pyRevit user config section `ViewHUD`
(`lib/revit/view_hud_config.py`). Gear on the bar applies them live.
Auto-start hooks `DocumentOpened` from `startup.py`.

## Switchers

| Badge | Shows | Click | Hidden when |
| --- | --- | --- | --- |
| Phase | The view's Phase | **Name:** next phase (right-click: previous). **Caret:** dropdown of all phases | view has no Phase parameter |
| Workset | Active workset | dropdown of user worksets | model is not workshared |
| Phase filter | View phase filter | dropdown of project phase filters | view has no Phase Filter parameter |
| Detail level | Coarse / Medium / Fine | dropdown | view has no detail level |
| Visual style | Wireframe, Hidden Line, Shaded, … | dropdown | view has no display style |
| Scope box | Assigned scope box, or None | dropdown of project scope boxes | view has no Scope Box parameter |
| Section box | On / Off | toggle 3D section box | view is not a 3D view |

Phase and Workset are on by default. The others are opt-in via the ⚙
badge (CommandSwitchWindow toggles).

The ⚙ badge is always last on the bar. Click it: **Top left** / **Top
middle** / **Top right** first, then chip toggles, then **Save**. Current
position is highlighted.

Phase hover is split: the name tooltip says it **cycles**; the caret
tooltip says **Pick a phase**.

The workset (and other pickers) use `DropdownSwitcher`. Phase uses the
same primitive with `on_left_click` / `on_right_click` so the label
cycles and only the caret opens the list.

The bundle is `lib/revit/context_switchers.py`
(`start_view_hud` / `stop_view_hud`). Phase logic is
`lib/revit/phase_label_wpf.py`. New switchers subclass `HudItem` /
`DropdownSwitcher` and register in `_FACTORIES`.

An active-design-option badge was considered and left out: the Revit API
has no setter for the active design option (only
`DesignOption.GetActiveDesignOptionId`), and the only known way to switch
it is a fragile out-of-process UI-automation hack on the status-bar
combobox (Jeremy Tammik's *DesignOptionModifier*) — not worth wiring
behind a HUD click.

## How it works

- Built on the reusable in-view switcher framework `lib/revit/view_hud.py`
  (`ViewHudHost` + `HudItem`/`TextSwitcher`/`DropdownSwitcher`): one
  borderless WPF window per document hosting a horizontal bar of
  switchers — no chrome, transparent background, owned by the Revit main
  window (so it never floats above other applications), non-activating via
  `WS_EX_NOACTIVATE`, so Revit keeps keyboard focus. Register via
  `get_hud_host`/`add_item`. `lib/revit/phase_label_wpf.py` is the
  phase-specific adapter.
- The window is parked off-screen until WPF completes its first layout
  pass (`SizeChanged`), then anchored — avoids the mis-sized first frame
  that `EnsureHandle` + `SizeToContent` produces.
- Clickable badges raise an `ExternalEvent` whose handler runs inside a
  Revit API context: the phase switcher sets the view's Phase parameter
  in a transaction, the workset switcher calls
  `WorksetTable.SetActiveWorksetId` (with a transactional retry). The
  badge text updates right after.
- Placement (top-left / top-center / top-right) comes from Settings.
  Device pixels convert to WPF units through `TransformFromDevice`
  (per-monitor DPI).
- Styling is muted: palette follows Revit's light/dark theme
  (`UIThemeManager`, Revit 2024+; light fallback on older versions), and
  badges idle at the configured opacity, turning solid while hovered.
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
  corner. Clicks on the badges go to the switchers, so that
  strip of the canvas is not click-able for model work.
- Does not print or export, and screen captures of the Revit window may or
  may not include it depending on the capture method.
- Crossing monitors with different DPI can be one tick late to rescale.
- Dropdown popups can be visually offset if the Revit view window straddles
  the seam between two monitors with different DPI scaling (a WPF `Popup`
  placement limitation); they render correctly when the view is within one
  monitor.

## Alternatives surveyed (2026-07)

All known Revit API variants for an in-viewport label, with why they were
or weren't chosen:

1. **TemporaryGraphicsManager / InCanvasControl** *(evaluated first; implementation kept in `lib/revit/phase_label.py`, ribbon button retired)* —
   Revit 2022+, bitmap badge, no model impact, no print. Model-anchored, so
   the badge snaps back to the corner only after pan/zoom ends.
2. **WPF overlay window** *(this panel, `lib/revit/phase_label_wpf.py`)* — a borderless, non-activating WPF window (badges receive clicks)
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
