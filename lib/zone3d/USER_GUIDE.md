# 3D Zone Tool - User Guide

## Overview

The 3D Zone tool automatically copies parameters from **source elements** (spatial zones like Rooms, Spaces, Areas, or 3D Zone families) to **target elements** (any elements contained within those zones). This is useful for propagating zone information (like fire ratings or zone numbers or zone names) to all elements within a zone.

## How It Works

The tool works in two steps:

1. **Containment Detection**: For each target element, the tool finds which source zone contains it
2. **Parameter Copying**: Parameters are copied from the containing zone to the target element

## Containment Strategies

The tool automatically selects the best containment strategy based on your source categories. Each strategy uses different methods to determine containment:

### 1. Room Strategy (Fastest) ⚡

**When Used**: Source category is **Rooms**

**How Containment Works**:
- Uses Revit's built-in `Room.IsPointInRoom()` method
- Checks multiple test points along the target element (especially useful for walls and doors)
- **Phase-aware**: Considers element and room phases - latest applicable phase wins
- Optimized by checking rooms on the same level first

**Things to Consider**:
- ✅ **Fastest method** - Rooms have optimized containment detection in Revit
- ✅ **Phase-aware** - Handles phased projects correctly
- ✅ **Most reliable** - Uses Revit's native room containment
- ⚠️ Rooms must be **placed** (not unplaced) - rooms with Area = 0 are skipped
- ⚠️ Rooms always contain themselves - Rooms are automatically excluded from target elements to prevent incorrect parameter writes
- ⚠️ For elements spanning multiple rooms, the first room found wins (deterministic by ElementId)

**Best For**: 
- Architectural projects with well-defined rooms
- Projects using Rooms for zone management
- Phased projects where room assignments change over time

---

### 2. Space Strategy (Fastest) ⚡

**When Used**: Source category is **Spaces** (MEP Spaces)

**How Containment Works**:
- Uses Revit's built-in `Space.IsPointInSpace()` method
- Checks multiple test points along the target element
- Optimized by checking spaces on the same level first

**Things to Consider**:
- ✅ **Very fast** - Spaces have optimized containment detection in Revit
- ✅ **MEP-focused** - Designed for mechanical, electrical, and plumbing zones
- ⚠️ Spaces must be **placed** (not unplaced)
- ⚠️ Spaces always contain themselves - Spaces are automatically excluded from target elements
- ⚠️ For elements spanning multiple spaces, the first space found wins

**Best For**:
- MEP projects using Spaces for zone management
- Projects where Spaces define HVAC zones

---

### 3. Area Strategy (Moderate Speed) 🐢

**When Used**: Source category is **Areas**

**How Containment Works**:
- Areas are 2D planar elements, so the tool creates a 3D solid by:
  1. Extracting the Area's boundary segments
  2. Creating a CurveLoop from the segments
  3. Extruding vertically from the Area's level to the level above (or default 10 feet)
  4. Checking if target element points are inside the 3D solid
- Uses optimized point-in-solid checking with `SolidCurveIntersectionOptions`
- Pre-computes all Area geometries in batch for better performance

**Things to Consider**:
- ⚠️ **Slower than Rooms/Spaces** - Requires geometry calculation
- ⚠️ Areas must have valid **boundary segments** - Areas without boundaries are skipped
- ⚠️ Areas must have a **Level** assigned - Areas without levels use a default height
- ⚠️ **Height calculation**: Uses level-to-level height, or default 10 feet if no level above exists
- ⚠️ **Boundary issues**: If Area boundaries are open or disconnected, solid creation fails
- ⚠️ Areas with complex or invalid boundaries may fail containment checks

**Best For**:
- Projects using Areas for zone planning
- Gross area calculations
- When Rooms/Spaces aren't suitable

**Troubleshooting**:
- If Areas aren't being detected, check that:
  - Areas have valid closed boundaries
  - Areas are assigned to a Level
  - Area boundaries don't have gaps or disconnected segments

---

### 4. Element Strategy (Slowest) 🐌

**When Used**: Source categories are **Mass** or **Generic Model** (including 3D Zone families)

**How Containment Works**:
- Extracts the element's geometry and finds the first solid with volume > 0
- Uses optimized point-in-solid checking with `SolidCurveIntersectionOptions`
- For **3D Zone families**: Filters Generic Models by family name containing "3DZone"
- Uses spatial indexing for performance (50-foot grid cells)
- Pre-computes all geometries in batch for better performance

**Things to Consider**:
- ⚠️ **Slowest method** - Requires geometry extraction and solid calculations
- ⚠️ Elements must have **valid geometry** with solids (not just surfaces or lines)
- ⚠️ Elements must have **volume > 0** - Elements without volume are skipped
- ⚠️ **3D Zone families**: Must have family name containing "3DZone" (case-sensitive)
- ⚠️ For overlapping zones, the element with the **lowest ElementId** wins (deterministic)
- ⚠️ **Performance**: Large numbers of source elements can be slow - consider using Rooms/Spaces instead

**Best For**:
- Custom 3D Zone families created from Rooms
- Mass elements used as zones
- When you need custom-shaped zones that don't match Rooms/Spaces

**Performance Tips**:
- Use spatial indexing (automatic) - reduces checks from O(n) to O(1) per cell
- Pre-compute geometries (automatic) - calculates all geometries once upfront
- Consider using Rooms/Spaces instead if possible - much faster

---

## Containment Detection Details

### Multiple Test Points

For better detection of linear elements (walls, doors, windows), the tool checks multiple points:
- **Point-based elements**: Single point (LocationPoint or bounding box center)
- **Linear elements**: Multiple points along the curve (25%, 50%, 75%) plus perpendicular offsets
- **Area elements**: Bounding box center

This ensures that walls and doors are correctly detected even if their insertion point is outside the zone.

### Phase-Aware Containment (Rooms Only)

For Room-based strategies, the tool considers phases:
- Elements exist in phases from `CreatedPhaseId` to `DemolishedPhaseId`
- Rooms have a `Phase` parameter
- The tool finds the **latest phase** where both element and room exist
- If an element moves between phases, the latest applicable phase wins

**Example**: 
- Element created in Phase 1, exists in Phase 2
- Room 1 exists in Phase 1
- Room 2 exists in Phase 2
- Element will get parameters from Room 2 (latest phase)

### Deterministic Selection

When multiple zones contain an element, the tool selects deterministically:
- **Rooms/Spaces/Areas**: First zone found (by ElementId order)
- **Elements**: Lowest ElementId wins

This ensures consistent results across runs.

---

## Configuration

### Creating a Configuration

1. Click **Config** button
2. Select **Add New Configuration**
3. Choose **Source Categories** (Rooms, Spaces, Areas, Mass, Generic Model, or 3D Zone)
4. Select **Source Parameters** to copy from
5. Select **Target Parameters** to copy to (must match source count)
6. Optionally select **Target Filter Categories** to limit which elements get updated
7. **Optional**: Enable **Use Linked Document** to search for source elements in a linked Revit model
8. Set **Execution Order** (configurations run in order)

### Source Categories

- **Rooms**: Fastest, phase-aware, best for architectural projects
- **Spaces**: Fast, best for MEP projects
- **Areas**: Moderate speed, requires valid boundaries
- **Mass**: Slow, requires solid geometry
- **Generic Model**: Slow, requires solid geometry
- **3D Zone**: Special filter for Generic Models with family name containing "3DZone"

### Linked Document Support

The tool can search for source elements in **linked Revit documents**, enabling cross-model parameter propagation.

**When to Use**:
- **MEP models** copying Room data from **Architectural models**
- **Site models** copying zone information from **Building models**
- **Discipline-specific models** using zones from **coordination models**
- Any scenario where source zones exist in a different model than target elements

**How It Works**:
1. Enable **Use Linked Document** checkbox in configuration
2. Select the linked document from the dropdown (shows all linked models)
3. The tool automatically:
   - Transforms coordinates between host and linked document coordinate systems
   - Handles phase mismatches between documents (checks all phases if names don't match)
   - Selects the room with lowest ElementId when multiple rooms contain an element

**Important Notes**:
- ⚠️ **Linked document must be loaded** - The linked model must be present in the project
- ⚠️ **Coordinate transformation** - Automatically handled, but linked models with complex transformations may require verification
- ⚠️ **Phase handling** - If phase names don't match between documents, the tool checks all phases in the linked document
- ⚠️ **Performance** - Slightly slower than same-document searches due to coordinate transformation
- ✅ **Validation** - The dropdown validates that the selected linked document exists

**Best Practices**:
- Ensure linked models are up-to-date before running
- Verify coordinate systems are compatible (Revit handles this automatically)
- Test with a small set of elements first to verify containment detection

### Target Filter Categories

Optional - limits which elements receive parameter updates. If not specified, all elements are checked.

**Example**: 
- Source: Rooms
- Target Filter: Walls, Doors, Windows
- Result: Only walls, doors, and windows get room parameters

---

## Execution

### Manual Execution

1. Click **Write** button
2. Review the list of enabled configurations
3. Confirm execution
4. Tool processes each configuration in order
5. Results show elements updated and parameters copied


## Performance Considerations

### Strategy Speed (Fastest to Slowest)

1. **Rooms** ⚡ - Fastest, uses native Revit API
2. **Spaces** ⚡ - Fast, uses native Revit API  
3. **Areas** 🐢 - Moderate, requires geometry calculation
4. **Elements** 🐌 - Slowest, requires full geometry extraction

### Optimization Features

The tool includes several optimizations:

- **Spatial Indexing** (Element strategy): 50-foot grid cells for fast lookup
- **Pre-computed Geometries**: Calculates all geometries once upfront
- **Bounding Box Pre-filtering**: Fast rejection before expensive checks
- **Level-based Grouping**: Checks same-level zones first
- **Parameter Pre-filtering**: Skips elements without required parameters
- **Skip Unchanged Values**: Only writes parameters that changed

### Performance Tips

1. **Use Rooms/Spaces when possible** - Much faster than Areas/Elements
2. **Limit Target Categories** - Use target filter categories to reduce checks
3. **Run Write manually** - More control for large updates
4. **Check element counts** - Large numbers of elements take longer

---

## Common Issues and Solutions

### Elements Not Getting Parameters

**Check**:
1. Source elements have the source parameters with values
2. Target elements have the target parameters (and they're writable)
3. Elements are actually contained within zones
4. Configuration is enabled
5. Target filter categories include your elements

### Containment Not Detected

**For Rooms/Spaces**:
- Ensure rooms/spaces are placed (not unplaced)
- Check that elements are actually inside room/space boundaries
- Verify room/space boundaries are correct
- **If using linked documents**: Verify the linked model is loaded and the link name matches exactly

**For Areas**:
- Check Area boundaries are closed and valid
- Ensure Area has a Level assigned
- Verify Area boundaries don't have gaps

**For Elements**:
- Verify elements have solid geometry (volume > 0)
- Check geometry is valid and not corrupted
- For 3D Zone families, ensure family name contains "3DZone"

### Performance Issues

**If processing is slow**:
1. Use Rooms/Spaces instead of Areas/Elements
2. Limit target filter categories
3. Reduce number of source elements
4. Check for geometry issues (corrupted elements)

### Phase Issues (Rooms)

**If elements get wrong room parameters**:
- Check element CreatedPhaseId and DemolishedPhaseId
- Verify room Phase parameter is set correctly
- Tool uses latest applicable phase - ensure phases are correct
- **If using linked documents**: If phase names don't match between documents, the tool checks all phases in the linked document and selects the room with lowest ElementId

---

## Best Practices

1. **Start with Rooms/Spaces**: Fastest and most reliable
2. **Use 3D Zone families**: When you need custom zones, create them from Rooms using "3D Zones from Rooms" tool
3. **Test with small sets**: Verify containment works before processing entire model
4. **Check parameter names**: Ensure source and target parameters exist and are writable
5. **Use target filters**: Limit which elements get updated for better performance
6. **Configuration order**: Order matters - earlier configurations run first
7. **Linked documents**: 
   - Ensure linked models are loaded and up-to-date
   - Verify link names match exactly (case-sensitive)
   - Test containment detection before processing large element sets
   - Coordinate with team to ensure linked models are synchronized

---

## Technical Notes

### Units

- All internal calculations use **Revit internal units (feet)**
- Level elevations, XYZ coordinates, and heights are in feet
- No unit conversion needed for internal calculations

### Parameter Types Supported

- **String**: Text parameters
- **Integer**: Whole numbers
- **Double**: Decimal numbers
- **ElementId**: References to other elements

### Storage

Configurations are stored in the Revit document using **Extensible Storage**, so they persist with the model and can be shared with team members.

---

## Examples

### Example 1: Copy Fire Rating from Rooms to Walls

**Configuration**:
- Source Category: Rooms
- Source Parameter: Fire Rating
- Target Parameter: Fire Rating
- Target Filter: Walls

**Result**: All walls get the Fire Rating from their containing room.

### Example 2: Copy MMI Code from 3D Zones to All Elements

**Configuration**:
- Source Category: 3D Zone (Generic Models with "3DZone" in name)
- Source Parameter: MMI
- Target Parameter: MMI
- Target Filter: (none - all elements)

**Result**: All elements get the MMI code from their containing 3D Zone.

### Example 3: Copy Zone Name from Areas to Equipment

**Configuration**:
- Source Category: Areas
- Source Parameter: Name
- Target Parameter: Zone Name
- Target Filter: Equipment

**Result**: All equipment gets the Name from their containing Area.

### Example 4: Copy Room Data from Architectural Model to MEP Elements

**Scenario**: MEP model needs room information (fire rating, room number, etc.) from the linked Architectural model.

**Configuration**:
- Source Category: Rooms
- Use Linked Document: ✅ **Enabled**
- Linked Document: Select "Architectural Model.rvt"
- Source Parameters: Fire Rating, Room Number, Room Name
- Target Parameters: Fire Rating, Room Number, Room Name
- Target Filter: Ducts, Pipes, Electrical Equipment

**Result**: All MEP elements in the MEP model get room parameters from rooms in the linked Architectural model.

**Use Case**: MEP designers can automatically assign room-based properties to their systems without manually entering data.

---

### Example 5: Copy Taktzon from Site Model to Building Elements

**Scenario**: Building model needs zone information (Taktzon, construction phase) from the linked Site/Planning model.

**Configuration**:
- Source Category: Rooms (or Areas)
- Use Linked Document: ✅ **Enabled**
- Linked Document: Select "Site Model.rvt"
- Source Parameters: Taktzon, Construction Phase
- Target Parameters: Taktzon, Construction Phase
- Target Filter: Walls, Floors, Structural Elements

**Result**: All building elements get zone information from the linked Site model.

**Use Case**: Construction planning where site zones define takt zones for different building areas.

---

### Example 6: Cross-Discipline Zone Coordination

**Scenario**: Structural model needs zone assignments from Architectural model for load calculations.

**Configuration**:
- Source Category: Rooms
- Use Linked Document: ✅ **Enabled**
- Linked Document: Select "Architectural Model.rvt"
- Source Parameters: Zone Classification, Occupancy Type
- Target Parameters: Zone Classification, Occupancy Type
- Target Filter: Structural Framing, Structural Columns

**Result**: Structural elements automatically get zone classifications from the architectural model.

**Use Case**: Ensuring structural elements match architectural zone assignments for code compliance and load calculations.

---

## Summary

The 3D Zone tool provides flexible parameter mapping from spatial zones to contained elements. Choose the right strategy for your needs:

- **Rooms**: Fastest, phase-aware, best for most projects
- **Spaces**: Fast, best for MEP projects
- **Areas**: Moderate speed, requires valid boundaries
- **Elements**: Slowest, most flexible, best for custom zones

**Cross-Model Support**: The tool can now search for source elements in linked Revit documents, enabling:
- MEP models to use Room data from Architectural models
- Site models to propagate zone information to Building models
- Automated parameter propagation across models
