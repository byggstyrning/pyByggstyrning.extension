# 3D View References

Places a work plane based family instance at the location and extent of a view, so that
sections, elevations, detail views and plan callouts can be seen in 3D.

Each instance remembers which view it belongs to (extensible storage on the instance, keyed on
the view's UniqueId). Running either tool again updates the existing instance: the size and
name are refreshed in place, and if the view has moved the instance is replaced. A view never
gets two references. References placed by hand carry no such link. Their size and position are left
alone. Running either tool still sets Export to IFC to Yes and Export to IFC As
to IfcVirtualElement on every instance of the family, including those.

The placement logic is shared by both tools and lives in `lib/revit/view_references.py`.

### Load Family

Loads the `3D View Reference` family into the current project with a single click. If the
project already has an older version of the family, it is upgraded and the references already
placed keep their values. Run it once in projects that got the family before the sheet number
text existed.

### Create References

Creates or updates references for many views at once.

* Lists sections, elevations, detail views and plan callouts (floor, ceiling, structural and
  area plans). The category dropdown shows all of them, or just one
* Shows view type, scale, sheet placement and whether a reference is already placed
* Search box: every word typed must occur somewhere in the row (view name, category, scale,
  sheet, sheet parameter); not case sensitive
* Filters: view category, on sheet / not on sheet, and reference placed / not placed.
  The category dropdown has a search field
* Max area (m²): hides views whose crop box (width x height) is larger, e.g. 6,25 to keep
  details and leave out building sections that share a sheet or a sheet parameter value with
  them. The view type cannot tell those apart. The Area column shows the value and sorts by
  number. Nothing is deleted from the model; the limit is remembered.
* Stacked sorting: click column headers in the order they should apply, e.g. View Scale then
  View Name sorts by scale and by name within each scale. A second click on a header flips its
  direction, a third takes it out of the sort. The order is shown bottom left. View Scale sorts
  by number (1:50 before 1:100).
* Sheet parameter column: pick any parameter found on the project's sheets and it is shown
  for each view's sheet, searchable like the other columns. The choice is remembered.
* Select/deselect all acts on the rows shown; ticks survive searching and filtering. Create
  works on the ticked rows that are shown.
* Highlight several rows (Shift or Ctrl click) and tick one of their checkboxes: all highlighted
  rows get the same state
* Right-click a row: **Go to view** or **Go to sheet**. Revit is blocked while the window is
  open, so the window closes and the view opens. Start the tool again and the window comes
  back as it was. A view can only be on one sheet, so there is never more than one sheet to go
  to.
* The window remembers itself per model: ticks, search, filters, the category dropdown, view depth
  and sort are saved whenever it closes (Cancel, X, Go to view, after creating) and restored at
  the next start. The first time in a model nothing is ticked.
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
* `View Width` and `View Height` come from the crop box, and `View Depth` as described
  above. The family type `Standard Reference` is used when present, otherwise the first
  type of the family.
* The view name row of the window builds the text written to `View Name`, left to right.
  The default part is the view's own name. Press + to add a view parameter, the sheet
  number, another sheet parameter, or a project information parameter. The box between two
  parts is the text placed there; leave it empty to join them directly. Hover a part and
  click the × on its top right corner to remove it. A new part is inserted in front of the
  view name. The formula is saved in the model. An empty result falls back to the view's
  own name.
* The sheet number row builds the text written to `Sheet Number` the same way. Its default
  part is the sheet's own number, and + adds a sheet or project parameter in front of it.
  Views that are not on a sheet, and empty results, show `-`. With a family that has no
  `Sheet Number` parameter the sheet text becomes a second line of the `View Name` text
  instead. Run the tool again after placing views on sheets, renumbering sheets, or
  changing either formula.
* The family's "View Name" text faces the viewer's side of the cut plane.
* Each instance is exported to IFC as `IfcVirtualElement`. That is set on the
  instance (Export to IFC = Yes, Export to IFC As = IfcVirtualElement), including
  instances placed by hand, whenever references are created or updated.

### The family file

`Load Family.pushbutton/3D View Reference.rfa` is the only family. It is saved in Revit 2025
format, so edit it in Revit 2025. Revit 2026 loads the same file. Saving it from a newer Revit
upgrades it and locks out Revit 2025.
