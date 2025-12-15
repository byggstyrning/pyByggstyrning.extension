# 3D Zone Parameter Mapping - Research & Implementation Guide

## Overview

This document outlines the research and implementation strategy for mapping parameters from spatial elements (Rooms, Spaces, Areas, Mass/Generic Model) to contained elements in Revit using the Revit API.

## Core Concepts

Revit does **not** store a native "contained elements" list on Rooms, Areas, or Spaces. We must:

1. Determine **spatial containment** relationships
2. Write data **downward** (Room/Space/Area → Element parameter)

There are three conceptually different containers:
- **Rooms** (`Autodesk.Revit.DB.Architecture.Room`)
- **Areas** (`Autodesk.Revit.DB.Area`)
- **Spaces** (`Autodesk.Revit.DB.Mechanical.Space`)
- **Zones** (HVAC Zones → `MechanicalZone`, Space-based → `Space` (MEP))
- **Mass/Generic Model Elements** (Custom spatial containers)

## Containment Strategies (Ranked by Performance)

### ✅ Fastest (Preferred)

**Point-in-room/space test using element Location**

- Works for: most model elements (walls, doors, equipment, furniture, etc.)
- Get a representative point for the element (usually midpoint)
- Ask the room/space: *is this point inside you?*
- Avoids geometry calculation completely
- **API**: `Room.IsPointInRoom(XYZ)`, `Space.IsPointInSpace(XYZ)`

### 🟡 Medium Performance

**BoundingBox intersection**

- Build room solid or bounding box
- Use `BoundingBoxIntersectsFilter`
- Good fallback, but coarse
- Less accurate than point-in-room but faster than full geometry

### 🔴 Slow (Avoid Unless Necessary)

**Full geometry intersection**

- `SpatialElementGeometryCalculator`
- `ElementIntersectsSolidFilter`
- Accurate but expensive
- Use only for Mass/Generic Model elements where point-in-room doesn't work

## Key Revit API Calls

### Collecting Elements

```python
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

# Collect rooms
rooms = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_Rooms)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Collect spaces
spaces = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_MEPSpaces)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Collect target elements
elements = FilteredElementCollector(doc)\
    .WhereElementIsNotElementType()\
    .ToElements()
```

**Important**: Cache collectors **once** and reuse them.

### Rooms

```python
from Autodesk.Revit.DB.Architecture import Room

# Room properties
room.Area
room.Level
room.LookupParameter("Name")
room.LookupParameter("Number")

# Point-in-room test (FASTEST)
room.IsPointInRoom(XYZ)
```

### Spaces

```python
from Autodesk.Revit.DB.Mechanical import Space

# Space properties
space.Area
space.Level
space.Number
space.Zone

# Point-in-space test (FASTEST)
space.IsPointInSpace(XYZ)
```

### Areas

**Areas do not support IsPointInRoom** ❗

You must use geometry or boundaries:

```python
from Autodesk.Revit.DB import SpatialElementBoundaryOptions

area.GetBoundarySegments(SpatialElementBoundaryOptions())
```

This is usually slow → try to avoid Areas if possible, or use as fallback only.

### Element Location (Critical!)

```python
loc = element.Location

# Common cases:
if isinstance(loc, LocationPoint):
    pt = loc.Point
elif isinstance(loc, LocationCurve):
    pt = loc.Curve.Evaluate(0.5, True)  # Midpoint
elif isinstance(loc, Location):
    # Try to get bounding box center
    bbox = element.get_BoundingBox(None)
    if bbox:
        pt = (bbox.Min + bbox.Max) / 2.0
```

### Geometry (Only if Needed)

```python
from Autodesk.Revit.DB import SpatialElementGeometryCalculator

calc = SpatialElementGeometryCalculator(doc)
results = calc.CalculateSpatialElementGeometry(room)
solid = results.GetGeometry()
```

### Writing Parameters

```python
param = element.LookupParameter("RoomName")
if param and not param.IsReadOnly:
    param.Set("Room 101")
```

## High-Performance Architecture

### ✅ Strategy: Element-Driven, Not Room-Driven

❌ **Don't loop**: room → get elements inside  
✅ **Do**: element → find its container

This scales MUCH better.

### Pseudocode Structure

```python
# 1. Collect rooms once
rooms = FilteredElementCollector(doc)\
    .OfClass(Room)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Optional: group rooms by LevelId
from collections import defaultdict
rooms_by_level = defaultdict(list)
for r in rooms:
    rooms_by_level[r.LevelId].append(r)

# 2. Collect elements once
elements = FilteredElementCollector(doc)\
    .WhereElementIsNotElementType()\
    .ToElements()

# 3. For each element
for el in elements:
    pt = get_representative_point(el)
    if not pt:
        continue
    
    # Only check rooms on same level
    for room in rooms_by_level.get(el.LevelId, []):
        if room.IsPointInRoom(pt):
            write_params(el, room)
            break
```

✅ This avoids:
- Geometry extraction
- Element filters per room
- N² scaling

## Containment Method Detection

Based on source categories, automatically detect the fastest containment method:

```python
def detect_containment_strategy(source_categories):
    """Detect optimal containment strategy based on source categories."""
    if BuiltInCategory.OST_Rooms in source_categories:
        return "room"  # Use IsPointInRoom (fastest)
    elif BuiltInCategory.OST_MEPSpaces in source_categories:
        return "space"  # Use IsPointInSpace (fastest)
    elif BuiltInCategory.OST_Areas in source_categories:
        return "area"  # Use boundary segments (slower)
    elif BuiltInCategory.OST_Mass in source_categories or \
         BuiltInCategory.OST_GenericModel in source_categories:
        return "element"  # Use geometry intersection (slowest)
    else:
        return None  # Unknown/unsupported
```

## Performance Killers (Avoid These)

❌ `SpatialElementGeometryCalculator` inside loops  
❌ New `FilteredElementCollector` per room  
❌ Geometry extraction on every element  
❌ Not grouping by Level  
❌ Writing parameters repeatedly without checking value change  
❌ Not caching spatial element collections

## Order of Operations

When multiple configurations exist, execution order matters:

1. **Lower order numbers execute first**
2. Configurations can depend on previous configurations
3. Example:
   - Order 1: Write Room Name to elements
   - Order 2: Write Room Number to elements
   - Order 3: Write Room Area to elements

This allows cascading parameter writes where later configs can use values from earlier configs.

## Practical Tips for pyRevit Speed

1. Wrap writes in **one transaction**:
   ```python
   from pyrevit import revit
   with revit.Transaction("Assign Zone Data"):
       # ... write operations
   ```

2. **Disable UI updates** (pyRevit does this well by default)

3. **Group by Level**: Only check spatial elements on the same level as target elements

4. **Cache collectors**: Don't recreate FilteredElementCollector in loops

5. **Early exit**: Once you find a container, break out of the loop

## API Call Summary

| Task              | API                                  | Performance |
| ----------------- | ------------------------------------ | ----------- |
| Collect elements  | `FilteredElementCollector`           | Fast        |
| Rooms             | `Architecture.Room`                  | Fast        |
| Containment check | `Room.IsPointInRoom()`               | Fastest     |
| Containment check | `Space.IsPointInSpace()`             | Fastest     |
| Containment check | `Area.GetBoundarySegments()`        | Slow        |
| Containment check | Geometry intersection                | Slowest     |
| Element position  | `LocationPoint / LocationCurve`      | Fast        |
| Write params      | `Parameter.Set()`                    | Fast        |
| Speed             | Level grouping + element-driven loop | Optimal    |

## Troubleshooting

### Elements Not Being Found

- Check that elements have valid locations
- Verify elements are on the same level as spatial elements
- Ensure spatial elements are not placed/unplaced

### Performance Issues

- Verify you're using element-driven loops, not room-driven
- Check that you're grouping by LevelId
- Ensure you're using `IsPointInRoom()` / `IsPointInSpace()` instead of geometry
- Cache collectors outside loops

### Parameters Not Writing

- Check parameter exists: `element.LookupParameter("ParamName")`
- Verify parameter is not read-only: `param.IsReadOnly`
- Check parameter storage type matches: `param.StorageType`
- Ensure transaction is active

### Configuration Order Issues

- Verify configurations are sorted by `order` field
- Check that dependencies between configs are correct
- Ensure earlier configs complete before later ones execute




