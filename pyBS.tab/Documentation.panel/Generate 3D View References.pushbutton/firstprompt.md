# Overall goal

Make a pyrevit script, that can place an instance of a workplane based family on the location and extent of selected section/elevation/callouts/view.
@script.py

use Claude 3.7 sonnet max in Cursor.

# User interface requirement

The UX should start by presenting a xaml window, where there are a few rows:

1. tool information
2. view categories, checkboxes for view categories in Revit. Sections, Elevations, Call-outs etc. All should be selected by default.
3. list box / datagrid with views, checkboxes for selecting views.
   the header should have a checkbox for selecting all checklists in view. all checkboxes should be checked by default.
   the listbox/datagrid should be able to be sortable asc/desc by clicking the headers.

present these columns and data about the views:
3a: Select (checkbox)
3b: View Name
3c: View Category
3d: View Scale
3e: Sheet Reference

4. actions bar, "Create 3D View References", "Isolate X created elements" should be disabled by default and enabled after created elements

Functional requirements:

* The view size / extent of the view should be calculated by the bounding box of the view crop / window since views can be odd shapes.
* The workplane based family, should be places in the same place and orientation as the view.
* The workplane based family has two instance based parameters called "View Height" and View Width", use those instance parameters to set the parameters coming from the view size.
* When the elements have been created by clicking "Create 3D View References", i want a button that can isolate the newly created elements by clicking "Isolate X created elements".
* the default family is named "3D View Reference" with Type called "Standard Reference" and should for now be hard coded in the family.

# Other requirements:

* The user interface should be XAML based.
* The codebase should follow best practices from the learnrevitapi course:
  @learnrevitapi.txt
* the software architecture should be code behind for simplicity

# Rules

* Make sure to follow the rules, remember to format strings for ironpython 2.7.
  @pyrevitdev.mdc
