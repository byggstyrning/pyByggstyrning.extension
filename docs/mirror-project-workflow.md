# Mirror Project workflow

Guided pyRevit wrapper around Revit’s native **Mirror Project** command. Version 1 captures a reusable baseline, disables eligible locked constraints without restoring them, launches Mirror Project, then compares the model after the native operation.

Reload the pyByggstyrning extension after any change under `lib/` or this document. The pushbutton script itself reloads on the next run.

## v1 scope

| In scope | Out of scope (deferred) |
| --- | --- |
| Preflight snapshot | **Rotate Project North** launcher |
| Baseline JSON capture | **Rotate True North** launcher |
| Constraint classification + unlock (no restore) | Programmatic mirror axis / angle |
| Deferred native **Mirror Project** launch | `ElementTransformUtils` whole-model mirror |
| Compare with baseline + review JSON / optional CSV | Auto sign-off / design-intent certification |
| Source-model isolated 3D review views | |

Use Revit’s built-in **Rotate Project North** and **Rotate True North** manually if needed. Compare may still infer `rotate_z_*` transforms when geometry matches a rotation, but this plugin does not post those commands in v1.

## Compare improvements from the v2 field report

- Transform inference uses model-point inlier consensus. Fixed coordinate
  controls (for example a Survey Point that native Mirror Project leaves in
  place) no longer outweigh correctly mirrored walls and families.
- Reports retain occurrence totals but compact repeated orientation,
  From/To-room, residual-constraint, and risk-cohort signals into actionable
  groups with example UniqueIds.
- Exact From/To swaps and mirror-induced orientation toggles are informational;
  changed adjacency pairs remain review findings.
- Facing vectors accept the reflected direction or a flip-state-backed reversed
  reflected direction. Wall curves compare both endpoints without assuming
  Revit preserves start/end order.
- A family absent from the post-operation `hosted_families` risk cohort is not
  called deleted. The cohort is state-dependent; category counts and stable
  door/window cohorts remain the deletion evidence.
- Prepare records eligible constraints that were already unlocked after the
  selected baseline was captured. This preserves manifest lineage across a
  repeated preparation run instead of reporting those changes as unrecorded.
- Duplicate category totals (for example `windows` and `Windows`) are emitted
  once.
- CSV exports include `occurrences` and `examples`.

## Deferred follow-up plan

1. Infer an arbitrary vertical mirror plane (offset and angle), not only
   origin-centred `mirror_x` / `mirror_y`.
2. Secondary-match recreated elements by category, type, host, and transformed
   location when UniqueId changes.
3. Store group member IDs so reports show exact gained/lost members.
4. Separate generated/internal Revit count drift from user-model categories.
5. Revise snapshot schema: per-collector timing, consistent nullability, area
   units, normalized floats, compact payloads, and optional redaction of local
   path/username.

## Recommended sacrificial-copy procedure

1. Sync, relinquish, and make a permanent backup of the production model.
2. Create an audited detached copy named so it cannot be mistaken for production, for example `PROJECT_MIRROR_TEST_01.rvt`.
3. Open that copy in the same Revit build used by the team.
4. On the **pyBS > Coordination** panel, run **Mirror Project**.

## Command menu

| Action | What it does |
| --- | --- |
| **Run Preflight** | Collects identity, warnings, counts, coordinates, links, rooms/spaces, wall attachments, constraints, groups, family orientation, and MEP unused connectors. Prints a Markdown summary. Does not modify the model. |
| **Capture Baseline** | Runs the same collectors and writes a versioned JSON file you choose. Keep this file next to the detached copy so the original and the mirrored file can share it. |
| **Prepare + Launch Mirror Project** | Requires a baseline. Classifies `OST_Constraints` dimensions, writes an immutable mutation manifest, unlocks API-eligible locked constraints in one transaction, then schedules native `Mirror Project` on the next Revit idle tick (so pyRevit dialogs do not skip the axis step). Revit prompts for axis, then direction. |
| **Compare With Baseline** | Recaptures the current model, infers a named transform (mirror X/Y or 90° Z rotations), compares UniqueIds, and prints blocking/review/info findings. Optional review-package JSON and CSV exports. |
| **Create Review Views From JSON** | Run in the non-mirrored source model. Select the review-package JSON (not the baseline). Resolves failed elements by UniqueId and creates persistent isolated 3D views grouped as Blocking, Missing, Geometry, Relationships, and Review. |

Review views use a padded section box and split groups above 150 elements into
numbered views. Names identify both issue and recommended response, for example
`MP Review - Missing Windows - Restore or Approve Deletion 01` and
`MP Review - Moved Walls - Reposition or Approve Movement`. The output report
also prints a full suggested action for each view. An additional
`MP Review - ALL Missing and Review Elements - Select Individually` view
isolates every graphical target together for element-by-element inspection.
Unresolved and non-graphical IDs are listed separately.
Count-only findings remain in the package but cannot create an element view.

## Constraint policy (v1)

- Eligible: ungrouped, unlabeled `Dimension` constraints with `IsLocked` and a settable segment count (`0` or `1`).
- Excluded and reported: grouped, multi-segment, labeled, non-dimension, missing, or not editable.
- Eligible locks are disabled **before** Mirror Project and **are not restored**. Original states remain in the baseline JSON and `*_manifest.json` for audit or later manual recreation.
- If you cancel Revit’s Mirror Project dialog, unlocked constraints stay unlocked.

## What the API cannot do

- Supply the Mirror Project axis or rotation angle. `PostCommand` only launches the built-in UI.
- Capture Revit’s modal Mirror Project error dialog as structured data. Export those errors from Revit after the command, then run **Compare With Baseline**.
- Mirror or rotate the whole model with `ElementTransformUtils`. That path is out of scope because it does not match native Mirror Project behaviour.
- Detach/reattach walls, repair MEP, or certify door handing / egress / slope design intent. Those remain review findings.

## Compare behaviour

- Same-document UniqueIds are the match key (Mirror Project transforms in place).
- Stable sample points (Project Base Point, Survey Point, wall/door locations) infer `identity`, `mirror_x`, `mirror_y`, or `rotate_z_90/180/270`.
- Locations and facing vectors are compared **after** that inferred transform, so a correct geometric mirror is not reported as movement.
- Manifest-recorded unlocks are expected. Unrecorded unlocks and residual eligible locks are flagged.
- Blocking examples: door/window/room/sheet count drift, unused MEP connectors increasing, schema mismatch, manifest unlock that is still locked.
- Review examples: new warning groups, attachment host changes, handedness flags, From/To room changes, residual excluded locks, True North angle change.

## Acceptance

The tool never auto-signs off design intent. Treat the mirrored model as production-ready only after a model manager reviews blocking and review findings, representative sheets, and (where used) IFC/federated placement.
