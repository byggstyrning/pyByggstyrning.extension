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

Loads the `3D View Reference` family into the current project with a single click.

### Create References

Creates or updates references for many views at once.

* Lists sections, elevations, detail views and plan callouts (floor, ceiling, structural and
  area plans), filtered by category
* Shows view type, scale, sheet placement and whether a reference is already placed
* Select/deselect all, and the selection survives changing the category filter
* **Show view depth** (off by default): off gives a thin plate on the cut plane; on gives a box
  from the cut plane to the far clip (sections, elevations, details) or to the view depth
  plane (floor plans). Views without a far limit, and ceiling plans, stay plates.
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
* The family's "View Name" text faces the viewer's side of the cut plane.
