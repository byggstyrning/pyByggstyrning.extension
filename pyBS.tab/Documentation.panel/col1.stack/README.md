# 3D View References

A pyRevit extension that places workplane-based family instances at the location and extent of views (sections, elevations, callouts, floor plans, etc.), allowing you to create visual references to your 2D views in 3D space.

### Load Family

Loads the 3D View Reference family into your current Revit project with a single click.

Features:

* One-click loading of the required family
* No prompts or additional steps required
* Automatically searches for the family file in the appropriate directories

### Create References

Creates multiple 3D reference markers for selected views in a batch operation.

* Shows an interface to select multiple views at once
* Filters views by category (Sections, Callouts)
* Displays view properties including scale and sheet placement
* Creates workplane-based families representing each selected view
* Sets size parameters based on view dimensions
* Includes an option to isolate created elements after completion
* Supports selecting/deselecting all views with one click

### Add to View

Creates a 3D reference marker for the currently active view.

* Single-click operation for the current view
* Automatically calculates the view's dimensions
* Places the reference at the appropriate location in 3D space
* Sets view name and dimensions on the reference marker
* Selects the newly created element
