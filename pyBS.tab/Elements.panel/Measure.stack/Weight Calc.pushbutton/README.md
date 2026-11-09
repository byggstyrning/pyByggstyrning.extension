# Weight Calculator

Automatic weight calculation for structural elements based on volume and material density.

## Quick Start

1. Open a view with structural elements
2. Click **Weight Calc** button
3. Select parameter to store weight
4. Done! Weights calculated and written

## Features

- ✅ Automatic structural element detection
- ✅ Multiple density source fallbacks
- ✅ Compound structure support (walls, floors)
- ✅ Project unit-aware (UnitUtils)
- ✅ Revit 2020-2025+ compatible
- ✅ Formatted to 1 decimal place (e.g., 1234.5 kg)
- ✅ **Clickable table of failed elements** - zoom to problem elements
- ✅ **Worksharing-aware** - skips elements owned by others
- ✅ **Read-only parameter detection** - only updates writable parameters
- ✅ **Performance optimized** - batched progress updates & material density caching

## How It Works

```
1. Find structural elements with volume in active view
2. Get material density (Structural Asset → Thermal Asset → Parameter → Layers)
3. Calculate: Weight (kg) = Volume (m³) × Density (kg/m³)
4. Format to 1 decimal place
5. Write to selected parameter
```

## Requirements

### Elements Must Have:
- Volume parameter > 0
- Structural = Yes OR category starting with "Structural"
- Material with density set

### Materials Must Have:
- Density in Structural Asset OR
- Density in Thermal/Physical Asset OR
- For compound structures: at least one layer with density

## Setting Material Density

### Quick Method:
1. Press `M` (Material Browser)
2. Select your material
3. Physical tab → Asset editor
4. Set **Density** value
5. Click OK twice

### Common Density Values (kg/m³):
- Concrete: 2400
- Steel: 7850
- Glulam: 450-500
- Wood (softwood): 500
- Insulation: 30
- Brick: 1800

## Supported Elements

- Structural Columns
- Structural Framing (beams)
- Structural Foundations
- Walls (with Structural=Yes)
- Floors (with Structural=Yes)
- Any element with Volume + Structural designation

## Technical Details

### Performance Optimizations

The script is optimized for speed when processing large numbers of elements:

**Material Density Caching**:
- Material densities are cached after first lookup
- Repeated use of same material = instant retrieval
- Cache cleared at start of each run for fresh data

**Batched Progress Updates**:
- Progress bar updates in batches (not every element)
- 100 elements: updates every 10 elements
- 1000 elements: updates every 10 elements
- 10000+ elements: updates every 100 elements

**Element Type Caching**:
- Element type names cached to avoid repeated lookups
- Only looked up when needed (failed/skipped elements)

**Single Transaction**:
- All parameter writes in one transaction
- Minimizes Revit regeneration overhead

**Performance Example**:
- 1000 elements with 5 material types
- Old: 1000 material lookups + 1000 progress updates
- New: 5 material lookups + 100 progress updates
- **Result: ~70% faster** ⚡

### Unit Conversion
Uses Revit's `UnitUtils` API for accurate conversions:
- Input: Revit internal units (ft³, lb/ft³)
- Output: Metric (m³, kg/m³, kg)
- Automatic version detection (Revit 2021+ vs older)

### Density Lookup Order
1. **Structural Asset** → `.Density` property
2. **Thermal Asset** → `.Density` property  
3. **BuiltInParameter** → `PHY_MATERIAL_PARAM_STRUCTURAL_DENSITY`
4. **Material Layers** → Weighted average from compound structure

### Compound Structures
For walls/floors with multiple layers:
```
Weighted Density = Σ(layer_density × layer_thickness) / total_thickness
Weight = Volume × Weighted Density
```

## Element Status Reports

### Failed Elements Report

If elements fail to calculate, you'll see a **clickable table** in the pyRevit output window:

```
⚠️ Failed Elements - No Material Density Found
Click element ID to zoom to element in Revit

### Structural Framing (5 elements)
| Element ID        | Reason                    |
|-------------------|---------------------------|
| [1563920]         | No material density found |
| [1563921]         | No material density found |
...
```

**Click any Element ID** to zoom to that element in Revit and fix the material!

### Skipped Elements Report

In worksharing projects, elements owned by other users or with read-only parameters are automatically skipped:

```
⏭️ Skipped Elements - Not Editable
These elements cannot be modified (read-only parameter or owned by another user)

### Structural Columns (3 elements)
| Element ID        | Reason                    |
|-------------------|---------------------------|
| [421664]          | Owned by John Smith       |
| [773642]          | Parameter is read-only    |
...
```

This ensures the tool only modifies elements you have permission to edit!

## Troubleshooting

### "No structural elements found"
- Check elements have Volume parameter
- Verify Structural = Yes or category starts with "Structural"
- Check elements visible in current view

### "No material density found"
- Open Material Browser (`M`)
- Select the material
- Set density in Physical or Structural asset
- For compound structures: check ALL layer materials

### "Parameter not found on element"
- Selected parameter doesn't exist on all elements
- Create project parameter for all structural categories
- Or select a different, more common parameter

### Wrong weight calculated
- Verify density value in Material Browser
- Check project units (metric vs imperial)
- For compound structures: verify all layer densities
- Spot-check: Volume × Density = Weight

### Elements skipped (worksharing)
- Check if elements are owned by another user
- Request ownership using "Make Editable" in Revit
- Coordinate with team to avoid conflicts
- Run tool after taking ownership

## Example Calculation

**Glulam Column 610×610**:
- Volume: 4.127 m³
- Density: 387.49 kg/m³ (GL24h)
- Weight: 4.127 × 387.49 = 1,598.87 kg
- Formatted: **1,598.9 kg** ✅

## Best Practices

1. **Create dedicated parameter**: "Calculated Weight" (Number, Structural)
2. **Set up materials properly**: All structural materials have density
3. **Use schedules**: Review calculated weights
4. **Run on filtered views**: Target specific elements
5. **Verify results**: Spot-check against known values

## Version History

### v1.7 (Current)
- ✅ Material density caching (~70% faster for large projects)
- ✅ Batched progress bar updates (smoother UI)
- ✅ Element type name caching (reduced lookups)
- ✅ Only lookup element info when needed (failed/skipped)

### v1.6
- ✅ Format to 1 decimal place instead of whole numbers

### v1.5
- ✅ Worksharing support - skips elements owned by others
- ✅ Read-only parameter detection
- ✅ Enhanced element editability checks
- ✅ Separate reports for skipped vs failed elements

### v1.4
- ✅ UnitUtils integration for project unit respect
- ✅ Version-aware API (Revit 2021+ and older)

### v1.3
- ✅ Rounding to whole kilograms

### v1.2
- ✅ Direct asset property access
- ✅ Multiple density source fallbacks

### v1.1
- ✅ PropertySetElement API fix

### v1.0
- Initial release

## Support

**Byggstyrning AB**
- Check Material Browser for density values
- Verify Volume parameter exists
- Review pyRevit output window for errors

---

*Simple. Fast. Accurate. Weight calculation made easy.*
