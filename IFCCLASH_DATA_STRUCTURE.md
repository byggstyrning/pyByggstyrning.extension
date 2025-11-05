# IfcClash Data Structure Analysis

## Source
https://github.com/IfcOpenShell/IfcOpenShell/blob/v0.8.0/src/ifcclash/ifcclash/ifcclash.py

## Data Structures

### ClashResult (Individual Clash)
```python
{
    "a_global_id": "2O2Fr$t4X7Zf8NOew3FLOH",  # IFC GUID for element A
    "b_global_id": "2O2Fr$t4X7Zf8NOew3FL0I",  # IFC GUID for element B
    "a_ifc_class": "IfcWall",                  # IFC class for element A
    "b_ifc_class": "IfcDuct",                  # IFC class for element B
    "a_name": "Wall-001",                      # Element A name
    "b_name": "Duct-002",                      # Element B name
    "type": "collision",                       # Clash type
    "p1": [10.5, 20.3, 5.0],                  # Clash point 1 (XYZ)
    "p2": [10.6, 20.4, 5.1],                  # Clash point 2 (XYZ)
    "distance": -0.15                          # Distance (negative = overlap)
}
```

### ClashSet (Collection of Clashes)
```python
{
    "name": "Architecture vs MEP",
    "mode": "collision",  # or "intersection", "clearance"
    "tolerance": 0.01,    # For intersection mode
    "allow_touching": false,  # For collision mode
    "clearance": 0.5,     # For clearance mode
    "a": [
        {
            "file": "architecture.ifc",
            "selector": "IfcWall",
            "mode": "i"
        }
    ],
    "b": [
        {
            "file": "mep.ifc",
            "selector": "IfcDuct",
            "mode": "i"
        }
    ],
    "clashes": {
        "2O2Fr$t4X7Zf8NOew3FLOH-2O2Fr$t4X7Zf8NOew3FL0I": ClashResult,
        "guid_a-guid_b": ClashResult,
        ...
    }
}
```

### Smart Grouped Clashes (from smart_group_clashes)
```python
{
    "Architecture vs MEP": [
        {
            "Architecture vs MEP - 1": [
                ["2O2Fr$t4X7Zf8NOew3FLOH", "2O2Fr$t4X7Zf8NOew3FL0I"],
                ["guid_a3", "guid_b3"]
            ],
            "Architecture vs MEP - 2": [
                ["guid_a4", "guid_b4"]
            ]
        }
    ]
}
```

## JSON Export Format

The `export_json()` method saves: `list[ClashSet]`

Example:
```json
[
    {
        "name": "Arch vs MEP",
        "mode": "collision",
        "allow_touching": false,
        "a": [...],
        "b": [...],
        "clashes": {
            "guid-guid": {...}
        }
    }
]
```

## Key Differences from Original Implementation

| Original | IfcClash Actual |
|----------|-----------------|
| `guid_a` / `guid_b` | `a_global_id` / `b_global_id` |
| `category_a` / `category_b` | `a_ifc_class` / `b_ifc_class` |
| `object_a` / `object_b` | `a_global_id` / `b_global_id` |
| Clash has `id` field | Clash keyed by "guid_a-guid_b" |
| Groups are flat list | Groups from smart_group_clashes are nested |
| No position field | Has `p1`, `p2` (3D points) |

## Updates Required

1. **API Client**: Update field names to match ifcclash
2. **Script**: Parse clash sets correctly
3. **Enrichment**: Use `a_global_id`/`b_global_id` instead of `guid_a`/`guid_b`
4. **Grouping**: Handle smart_group_clashes format
5. **UI**: Display correct field names

## Smart Grouping Logic

The `smart_group_clashes` method:
1. Takes clash sets with individual clashes
2. Uses OPTICS clustering on clash positions
3. Groups nearby clashes (within max_clustering_distance)
4. Returns dict with clash set names containing grouped GUID pairs

Output format:
```
{
    "ClashSetName": [
        {
            "ClashSetName - 1": [[guid_a, guid_b], ...],
            "ClashSetName - 2": [[guid_a, guid_b], ...]
        }
    ]
}
```
