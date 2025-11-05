# Clash Explorer Panel

The Clash Explorer panel provides advanced tools for visualizing and interacting with clash detection results from external clash detection systems in Revit.

## Features

### Clash Explorer

The Clash Explorer tool connects to external clash detection APIs to import, visualize, and interact with clash test results directly within your Revit model.

#### Key Features

1. **API Integration**
   - Connect to external clash detection services
   - Secure API key storage in project using extensible storage
   - Settings saved per project for team collaboration

2. **Fast GUID Lookup**
   - Builds comprehensive GUID index across ALL Revit models
   - Searches both open document and all linked models
   - Instant element lookup for thousands of clashes

3. **Clash Matrix Display**
   - View all clash tests for your project
   - See clash counts and status at a glance
   - Sort and filter clash tests

4. **Advanced Grouping and Filtering**
   - Group clashes by Category, Level, or both
   - Multi-level grouping for better organization
   - Filter by Revit categories and levels
   - Sort groups by clash count (descending)

5. **Rich Metadata**
   - Automatically enriches clash data with Revit information
   - Shows Revit categories for clashing elements
   - Displays level information for spatial context

6. **Visual Highlighting**
   - Creates dedicated 3D views for each clash or clash group
   - Color-coded visualization:
     - **Green**: Elements in your model
     - **Red**: Elements in linked models
   - Automatic section boxes around clash areas
   - Other models set as underlay for context

7. **Clash Details**
   - Drill down into individual clashes
   - View element GUIDs, categories, and distances
   - Highlight individual clashes or entire groups

## Usage

### Setup

1. Click the **Clash Explorer** button in the Clashes panel
2. In the Settings tab:
   - Enter your clash detection API URL
   - Enter your API key
   - Click **Save Settings** (stored in project)

### Loading Clash Tests

1. Click **Load Clash Tests** to retrieve available tests
2. Switch to the **Clash Tests** tab
3. Select a clash test from the list
4. Click **Select Clash Test**

### Viewing and Filtering Clashes

1. In the **Clash Groups** tab:
   - View all clash groups sorted by count
   - Use **Filter by Category** to show specific categories
   - Use **Filter by Level** to show specific levels
   - Use **Group by** to organize clashes hierarchically

2. Click **Apply Filters** to update the view

### Highlighting Clashes

#### Highlighting a Group
1. Select a clash group in the table
2. Click the **Highlight** button in the Actions column
3. A new 3D view is created showing:
   - All elements in the clash group
   - Green highlighting for your model elements
   - Red highlighting for linked model elements
   - Section box focused on the clash area

#### Highlighting Individual Clashes
1. Select a clash group to load details
2. Switch to the **Clash Details** tab
3. Click **Highlight** for any individual clash
4. Or click **Highlight All in Group** to show all clashes

### View Naming Convention

Created views are named: `Clash - [Group/Clash Name]`

This allows you to:
- Keep multiple clash views organized
- Reference specific clashes in coordination meetings
- Track which clashes have been reviewed

## API Requirements

The Clash Explorer expects the following API endpoints:

### Get Clash Tests
```
GET /api/v1/clash-tests?project_id={project_id}
```

Response:
```json
{
  "data": [
    {
      "id": "test123",
      "name": "Arch vs MEP",
      "total_clashes": 45,
      "status": "completed",
      "last_run": "2025-10-04T10:30:00Z"
    }
  ]
}
```

### Get Clash Groups
```
GET /api/v1/clash-tests/{test_id}/groups
```

Response:
```json
{
  "data": [
    {
      "id": "group123",
      "name": "Walls vs Ducts",
      "clash_count": 12,
      "clashes": [
        {
          "id": "clash001",
          "guid_a": "2O2Fr$t4X7Zf8NOew3FLOH",
          "guid_b": "2O2Fr$t4X7Zf8NOew3FL0I",
          "category_a": "Walls",
          "category_b": "Ducts",
          "distance": -0.15
        }
      ]
    }
  ]
}
```

### Search by GUIDs (Optional)
```
POST /api/v1/clashes/search
{
  "guids": ["2O2Fr$t4X7Zf8NOew3FLOH", ...]
}
```

## Technical Details

### GUID Lookup Performance

The GUID lookup system is optimized for performance:
- **One-time indexing**: Built once when tool opens
- **All models included**: Host + all linked models
- **Fast dictionary lookup**: O(1) element retrieval
- **Handles thousands of clashes**: Suitable for large projects

### Extensible Storage Schema

Settings are stored using a unique schema GUID:
```
a7f3d8e2-5c1a-4b9e-8f2d-3e4a5b6c7d8e
```

This ensures:
- Per-project configuration
- No conflicts with other tools
- Team-wide settings sharing

### Color Coding Logic

- **Green (0, 200, 0)**: Elements you can modify in your model
- **Red (255, 0, 0)**: Elements in linked models requiring coordination

### View Creation

Views are created as:
- 3D Isometric views
- Detail Level: Fine
- Section boxes with 10% padding around clash elements
- Uninvolved models set to halftone

## Best Practices

1. **Setup Once**: Save your API settings once per project
2. **Filter Early**: Use category and level filters to focus on relevant clashes
3. **Group Wisely**: Use appropriate grouping to organize clashes logically
4. **Review Systematically**: Work through groups from highest to lowest count
5. **Name Views**: Keep created views for documentation and coordination
6. **Clean Up**: Delete resolved clash views periodically

## Troubleshooting

### "No elements found in current model"
- The elements may be in a linked model not currently loaded
- Check that all relevant links are loaded in Revit
- Verify IFC GUID parameters are present on elements

### "Failed to create clash view"
- Ensure you have permission to create views
- Check that view names are unique
- Verify the model has a valid 3D view type

### "Authentication failed"
- Verify your API key is correct
- Check that the API URL is accessible
- Ensure your API key has not expired

### Slow Performance
- GUID indexing happens once at startup
- Subsequent operations are fast
- Consider filtering to reduce visible clashes

## Future Enhancements

Potential features for future releases:
- Poll API for selected elements to show clash count bubble
- Export clash reports to Excel
- Clash status tracking (New, Active, Resolved)
- Integration with issue tracking systems
- Batch clash resolution workflows

## Support

For issues or feature requests, please contact Byggstyrning AB or submit an issue to the project repository.
