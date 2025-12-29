# PyRevit User Selection Best Practices

This document outlines best practices for handling user selections in PyRevit scripts and extensions.

## Table of Contents

1. [Getting Current Selection](#getting-current-selection)
2. [Interactive Element Picking](#interactive-element-picking)
3. [Selection Validation](#selection-validation)
4. [Common Patterns](#common-patterns)
5. [Best Practices](#best-practices)
6. [Examples](#examples)

---

## Getting Current Selection

### Using `revit.get_selection()`

The most common way to get the current user selection in PyRevit is using `revit.get_selection()`:

```python
from pyrevit import revit

# Get current selection
selection = revit.get_selection()

# Check if selection exists and has elements
if selection and selection.element_ids:
    # Process selected elements
    for element_id in selection.element_ids:
        element = revit.doc.GetElement(element_id)
        # Do something with element
else:
    # No selection - handle accordingly
    print("No elements selected")
```

### Selection Object Properties

The selection object returned by `revit.get_selection()` has the following useful properties:

- `selection.element_ids` - List of ElementId objects
- `selection.elements` - List of Element objects (lazy-loaded)
- `selection.first_element` - First element in selection (or None)
- `selection.count` - Number of selected elements

---

## Interactive Element Picking

### Using `pick_element()`

For scripts that need to prompt the user to select an element interactively, use `pyrevit.revit.selection.pick_element()`:

```python
from pyrevit import revit
from pyrevit.revit.selection import pick_element

# Pick a single element
selected_element = pick_element()

if selected_element:
    # Element was selected
    print("Selected: {}".format(selected_element.Name))
else:
    # User cancelled
    print("Selection cancelled")
```

### Pick Element with Filter

You can filter which elements can be selected using a filter function:

```python
from pyrevit import revit
from pyrevit.revit.selection import pick_element

# Filter function - only allow walls
def wall_filter(element):
    return element.Category.Name == "Walls"

# Pick element with filter
selected_wall = pick_element(filter_func=wall_filter)

if selected_wall:
    print("Selected wall: {}".format(selected_wall.Name))
```

### Other Pick Functions

PyRevit provides several pick functions for different scenarios:

- `pick_element()` - Pick a single element
- `pick_elements()` - Pick multiple elements
- `pick_edges()` - Pick edges
- `pick_faces()` - Pick faces
- `pick_points()` - Pick points

**Reference:** [PyRevit Selection Documentation](https://docs.pyrevitlabs.io/reference/pyrevit/revit/selection/#pyrevit.revit.selection.pick_element)

---

## Selection Validation

### Always Validate Selection

Never assume a selection exists. Always check before processing:

```python
from pyrevit import revit

selection = revit.get_selection()

# Good: Check selection exists and has elements
if not selection or not selection.element_ids:
    forms.alert("Please select at least one element.", title="No Selection")
    script.exit()

# Process selection
for element_id in selection.element_ids:
    element = revit.doc.GetElement(element_id)
    # Process element...
```

### Validate Element Types

Check that selected elements are of the expected type:

```python
from pyrevit import revit
from pyrevit import forms

selection = revit.get_selection()

if not selection or not selection.element_ids:
    forms.alert("Please select elements.", title="No Selection")
    script.exit()

# Filter to only walls
walls = []
for element_id in selection.element_ids:
    element = revit.doc.GetElement(element_id)
    if element.Category.Name == "Walls":
        walls.append(element)

if not walls:
    forms.alert("No walls found in selection.", title="Invalid Selection")
    script.exit()
```

---

## Common Patterns

### Pattern 1: Use Selection or Pick

A common pattern is to use existing selection if available, otherwise prompt for selection:

```python
from pyrevit import revit
from pyrevit.revit.selection import pick_element

# Try to get current selection first
selection = revit.get_selection()

if selection and selection.element_ids:
    # Use existing selection
    element = selection.first_element
else:
    # No selection - prompt user
    element = pick_element()
    if not element:
        script.exit()  # User cancelled

# Process element
print("Processing: {}".format(element.Name))
```

### Pattern 2: Select by Parameter Value

If no selection exists, select elements by parameter value:

```python
from pyrevit import revit
from pyrevit import forms

selection = revit.get_selection()

if not selection or not selection.element_ids:
    # No selection: select elements by MMI value
    doc = revit.doc
    collector = FilteredElementCollector(doc, doc.ActiveView.Id)
    elements = collector.WhereElementIsNotElementType().ToElements()
    
    target_elements = []
    for element in elements:
        param = element.LookupParameter("MMI")
        if param and param.AsString() == "200":
            target_elements.append(element.Id)
    
    if target_elements:
        # Set selection
        uidoc = revit.uidoc
        uidoc.Selection.SetElementIds(List[ElementId](target_elements))
        selection = revit.get_selection()
    else:
        forms.alert("No elements found with MMI value 200.")
        script.exit()

# Process selection
for element_id in selection.element_ids:
    element = revit.doc.GetElement(element_id)
    # Process element...
```

### Pattern 3: Set Selection Programmatically

You can set the selection programmatically:

```python
from pyrevit import revit
from System.Collections.Generic import List
from Autodesk.Revit.DB import ElementId

uidoc = revit.uidoc
doc = revit.doc

# Collect elements to select
element_ids = []
collector = FilteredElementCollector(doc, doc.ActiveView.Id)
for element in collector.WhereElementIsNotElementType().ToElements():
    if element.Category.Name == "Walls":
        element_ids.append(element.Id)

# Set selection
if element_ids:
    uidoc.Selection.SetElementIds(List[ElementId](element_ids))
    
    # Optional: Zoom to selection
    uidoc.ShowElements(List[ElementId](element_ids))
    uidoc.RefreshActiveView()
else:
    print("No walls found to select")
```

---

## Best Practices

### 1. Always Provide User Feedback

Inform users when selection is required or when selection is invalid:

```python
from pyrevit import forms

if not selection or not selection.element_ids:
    forms.alert(
        "Please select at least one element before running this tool.",
        title="Selection Required"
    )
    script.exit()
```

### 2. Use Context-Specific Selection

Use `__context__` in your script metadata to control when the script is available:

```python
__context__ = 'Selection'  # Only available when elements are selected
```

Available contexts:
- `'Selection'` - Requires selection
- `'zerodoc'` - Available in zero-document state
- `'active'` - Requires active document
- `'not_zerodoc'` - Not available in zero-document state

### 3. Handle Cancellation Gracefully

When using pick functions, always handle cancellation:

```python
from pyrevit.revit.selection import pick_element

element = pick_element()
if not element:
    # User cancelled - exit gracefully
    script.exit()
```

### 4. Filter Selection Early

Filter invalid elements early to provide better user feedback:

```python
selection = revit.get_selection()

if not selection or not selection.element_ids:
    forms.alert("Please select elements.", title="No Selection")
    script.exit()

# Filter to valid elements
valid_elements = []
invalid_reasons = []

for element_id in selection.element_ids:
    element = revit.doc.GetElement(element_id)
    
    # Check if element is valid for this operation
    if element.Category.Name == "Walls":
        valid_elements.append(element)
    else:
        invalid_reasons.append("{} is not a wall".format(element.Name))

# Report invalid elements
if invalid_reasons:
    forms.alert(
        "Some elements were skipped:\n" + "\n".join(invalid_reasons),
        title="Invalid Elements"
    )

if not valid_elements:
    forms.alert("No valid elements found in selection.")
    script.exit()
```

### 5. Use Transactions for Modifications

Always wrap modifications in transactions:

```python
from pyrevit import revit

selection = revit.get_selection()

if not selection or not selection.element_ids:
    script.exit()

with revit.Transaction("Modify Elements"):
    for element_id in selection.element_ids:
        element = revit.doc.GetElement(element_id)
        param = element.LookupParameter("MMI")
        if param and not param.IsReadOnly:
            param.Set("200")
```

### 6. Consider Performance for Large Selections

For large selections, consider processing in batches or providing progress feedback:

```python
from pyrevit import revit
from pyrevit import forms

selection = revit.get_selection()

if not selection or not selection.element_ids:
    script.exit()

total = len(selection.element_ids)
processed = 0

with revit.Transaction("Process Elements"):
    for element_id in selection.element_ids:
        element = revit.doc.GetElement(element_id)
        # Process element...
        processed += 1
        
        # Optional: Update progress
        if processed % 10 == 0:
            print("Processed {}/{} elements...".format(processed, total))

forms.alert("Processed {} elements.".format(processed))
```

---

## Examples

### Example 1: Simple Selection Processing

```python
"""Set MMI value on selected elements."""

__title__ = "Set MMI Value"
__author__ = "Your Name"

from pyrevit import revit, forms
from pyrevit.revit import Transaction

selection = revit.get_selection()

if not selection or not selection.element_ids:
    forms.alert("Please select elements first.", title="No Selection")
    script.exit()

# Get MMI value from user
mmi_value = forms.ask_for_one_item(
    items=["200", "225", "250", "300", "400"],
    default="200",
    prompt="Select MMI value:",
    title="MMI Value"
)

if not mmi_value:
    script.exit()

# Set MMI parameter
with revit.Transaction("Set MMI Value"):
    success_count = 0
    for element_id in selection.element_ids:
        element = revit.doc.GetElement(element_id)
        param = element.LookupParameter("MMI")
        if param and not param.IsReadOnly:
            param.Set(mmi_value)
            success_count += 1

forms.alert("Set MMI value {} on {} elements.".format(mmi_value, success_count))
```

### Example 2: Pick Element with Validation

```python
"""Analyze selected or picked element."""

__title__ = "Analyze Element"
__author__ = "Your Name"

from pyrevit import revit, forms
from pyrevit.revit.selection import pick_element

# Try existing selection first
selection = revit.get_selection()
element = None

if selection and selection.element_ids:
    element = selection.first_element
else:
    # Pick element interactively
    element = pick_element()
    if not element:
        script.exit()

# Validate element
if not element.Category:
    forms.alert("Selected element has no category.")
    script.exit()

# Display element information
info = [
    "Name: {}".format(element.Name),
    "Category: {}".format(element.Category.Name),
    "Element ID: {}".format(element.Id.IntegerValue),
    "Type: {}".format(element.GetType().Name)
]

forms.alert("\n".join(info), title="Element Information")
```

### Example 3: Filtered Selection

```python
"""Process only walls from selection."""

__title__ = "Process Walls"
__author__ = "Your Name"

from pyrevit import revit, forms

selection = revit.get_selection()

if not selection or not selection.element_ids:
    forms.alert("Please select elements.", title="No Selection")
    script.exit()

# Filter to walls only
walls = []
other_elements = []

for element_id in selection.element_ids:
    element = revit.doc.GetElement(element_id)
    if element.Category and element.Category.Name == "Walls":
        walls.append(element)
    else:
        other_elements.append(element.Name if element.Name else "Unnamed")

# Report filtered results
if other_elements:
    forms.alert(
        "Skipped {} non-wall elements.".format(len(other_elements)),
        title="Filtered Selection"
    )

if not walls:
    forms.alert("No walls found in selection.")
    script.exit()

# Process walls
with revit.Transaction("Process Walls"):
    for wall in walls:
        # Process wall...
        pass

forms.alert("Processed {} walls.".format(len(walls)))
```

---

## Additional Resources

- [PyRevit Selection API Documentation](https://docs.pyrevitlabs.io/reference/pyrevit/revit/selection/)
- [PyRevit Forms Documentation](https://docs.pyrevitlabs.io/reference/pyrevit/forms/)
- [Revit API Selection Documentation](https://www.revitapidocs.com/)

---

**Last Updated:** December 2024  
**References:**
- [PyRevit Selection Documentation](https://docs.pyrevitlabs.io/reference/pyrevit/revit/selection/#pyrevit.revit.selection.pick_element)
- PyRevit Community Best Practices
