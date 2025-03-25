# Generate 3D View References

A pyRevit tool that places workplane-based family instances at the location and extent of selected views (sections, elevations, callouts, floor plans, etc.).

## Features

- **View Selection UI**: XAML-based window with categorized view selection
- **Workplane Family Placement**: Places families at the exact position and orientation of views
- **Dynamic Parameter Adjustment**: Sets family parameters based on view dimensions 
- **Element Isolation**: Ability to isolate created elements for easy visualization
- **3D View Reference Family**: Uses Revit's standard 3D View Reference family

## Functional Requirements

### User Interface

- **Tool Information**: Clear description and instructions at the top of the window
- **View Category Selection**: Checkboxes for view categories (Sections, Elevations, Callouts, Floor Plans, Ceiling Plans, 3D Views)
  - All view categories checked by default
  - Filtering views dynamically when categories are selected/deselected

### View Selection Grid

- **Data Display**: Interactive grid showing available views with these columns:
  - Select (checkbox)
  - View Name
  - View Category
  - View Scale
  - Sheet Reference (if placed on sheet)
- **Selection Features**:
  - "Select All" checkbox to toggle all view selections
  - Individual view selection via checkboxes
  - All views selected by default
  - Sortable columns (ascending/descending) by clicking headers

### Family Placement

- **View Size Calculation**: Uses the bounding box of the view crop/window to determine dimensions
- **Family Requirements**:
  - Requires the "3D View Reference" family
  - Family must be workplane-based with instance parameters "View Width" and "View Height"
- **Instance Placement**:
  - Places family at the same location and orientation as the view
  - Sets "View Height" and "View Width" parameters to match view dimensions

### After Creation

- **Feedback**: Shows success message with number of created elements
- **Isolation**: Button to isolate created elements in current view (enabled after creation)
  - Button text updates to show number of elements ("Isolate X created elements")

### Logging

- **Debug Information**: Comprehensive logging to track the tool's operation
  - Uses `logger.debug()` for detailed operation information
  - Uses `logger.info()` for general process steps
  - Uses `logger.warning()` for non-critical issues (e.g., missing bounding box)
  - Uses `logger.error()` for exceptions and failures
- **Log Information Includes**:
  - View selection and filtering details
  - Family detection and activation
  - Bounding box calculations
  - Parameter setting
  - Element creation success/failure
  - Exception details when errors occur

## Technical Requirements

- **XAML-Based UI**: Uses WPF for the user interface
- **Code-Behind Architecture**: For simplicity and maintainability
- **Revit API Integration**: Proper use of Revit's View and Family APIs
- **IronPython 2.7 Compatibility**: String formatting and other operations compatible with IronPython
- **Transaction Handling**: Proper use of transactions for Revit operations
- **Error Handling**: 
  - Comprehensive exception handling at every level
  - Defensive programming with property existence checks
  - Default fallback values for safety
  - Detailed error logging with actionable messages
- **Memory Management**: Proper scope management for UI elements

## Usage

1. Ensure the "3D View Reference" family is loaded into your project
2. Run the tool from the pyRevit tab > Documentation panel
3. If the required family is not found, you'll be prompted to load it
4. Select desired view categories and specific views
5. Click "Create 3D View References" 
6. Use "Isolate X created elements" to highlight the created references in the current view

## Notes

- The tool requires the "3D View Reference" family to be loaded in your project
- **Enhanced Robustness**:
  - Uses multi-level error handling to prevent crashes from property access issues
  - Includes fallback mechanisms for critical operations
  - Handles special cases like StructuralType namespace issues
  - Provides detailed debug logging to identify and resolve issues
- The family must contain instance parameters named "View Width" and "View Height"
- Family detection happens at tool startup with clear feedback if not found
- Views without a defined bounding box will use a default size
- The tool provides detailed logging to help troubleshoot any issues during operation

# Overall goal

Make a pyrevit script, that can place an instance of a workplane based family on the location and extent of selected section/elevation/callouts/view.
@script.py

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
