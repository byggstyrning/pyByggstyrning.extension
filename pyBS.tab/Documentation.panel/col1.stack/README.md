# 3D View References

Places a work plane based family instance at the location and extent of a view, so that
sections, elevations, detail views and plan callouts can be seen in 3D.

Each instance remembers which view it belongs to (extensible storage on the instance, keyed on
the view's UniqueId). Running either tool again updates the existing instance: the size and
name are refreshed in place, and if the view has moved the instance is replaced. A view never
gets two references. References created before this was added carry no such link and are left
alone; delete them by hand.

The placement logic is shared by both tools and lives in `lib/revit/view_references.py`.

### Load Family

Loads the `3D View Reference` family into the current project with a single click. If the
project already has an older version of the family, it is upgraded and the references already
placed keep their values. Run it once in projects that got the family before the sheet number
text existed.

### Create References

Creates or updates references for many views at once.

* Lists sections, elevations, detail views and plan callouts (floor, ceiling, structural and
  area plans), filtered by category
* Shows view type, scale, sheet placement and whether a reference is already placed
* Search box: every word typed must occur somewhere in the row (view name, category, scale,
  sheet, sheet parameter); not case sensitive
* Filters: on sheet / not on sheet, and reference placed / not placed
* Sheet parameter column: pick any parameter found on the project's sheets and it is shown
  for each view's sheet, searchable like the other columns. The choice is remembered.
* Select/deselect all acts on the rows shown; ticks survive searching and filtering. Create
  works on the ticked rows that are shown.
* Highlight several rows (Shift or Ctrl click) and tick one of their checkboxes: all highlighted
  rows get the same state
* **Show view depth** (off by default): off gives a thin plate on the cut plane; on gives a box
  from the cut plane to the far clip (sections, elevations, details) or to the view depth
  plane (floor plans). Views without a far limit, and ceiling plans, stay plates.
* **Depth (mm)**: type a depth and every reference is drawn that deep, measured from the cut
  plane in the direction the view looks (upwards for a ceiling plan). A typed value wins over
  the checkbox; empty means the 10 mm plate, or the view depth if ticked. The value is
  remembered.
* Reports how many references were created, updated and skipped, with the reason for each skip
* Can isolate the references it just created or updated in the active view

### Add to View

Creates or updates the reference for the active view, as a plate on the cut plane, and
selects it.

### How the reference is placed

* The work plane is built from the view's own right and up directions, so the instance is
  oriented like the view whatever direction the view faces.
* It is centred on the crop box, on the cut plane. For plans the cut plane height comes from
  the view range (level elevation + cut plane offset).
* `View Width` and `View Height` come from the crop box, `View Name` from the view, and
  `View Depth` as described above. The family type `Standard Reference` is used when present,
  otherwise the first type of the family.
* The number of the sheet the view is placed on goes in `Sheet Number`, shown below the view
  name; views that are not on a sheet show `-`. With a family that has no `Sheet Number`
  parameter (Revit 2024/2025, or a project where Load Family has not been run since) it becomes
  a second line of the `View Name` text instead. Run the tool again after placing views on
  sheets or renumbering sheets to refresh the text.
* The family's "View Name" text faces the viewer's side of the cut plane.

### The family files

* `Load Family.pushbutton/3D View Reference.rfa` is saved in Revit 2024 format so that every
  supported Revit version can load it. Saving it from a newer Revit upgrades it and locks out
  the older versions, so edit it in Revit 2024.
* `Load Family.pushbutton/2026/3D View Reference.rfa` is a Revit 2026 copy with the
  `Sheet Number` text. It is generated from the file above by `build_2026_family.py`
  (`pyrevit run`, see the script); rebuild it when the 2024 file changes. Load Family uses the
  newest file the running Revit can open, so a folder for a later version can be added the same
  way.

The `Sheet Number` text is a nested label family, not a second model text. The family keeps its
`View Name` text at the frame's left edge by grouping it with an invisible model line that is
locked to the left reference plane, and the API can neither dimension a model text nor put model
lines in a group. A nested family instance has references, so the build script locks its centre
reference to the same reference plane.
