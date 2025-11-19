# MMI Colorizer

A PyRevit tool for visualizing MMI (Modell-, Modenhets- og Informasjonsnivå) values in Revit using official color codes from the [Norwegian MMI Guide](https://mmi-veilederen.no/?page_id=85).

## Features

### 🖱️ Normal Click - Temporary Color Overrides
- **Toggle On**: Colors elements in the active view based on their MMI values
- **Toggle Off**: Removes all color overrides
- **Use Case**: Quick visual inspection, temporary coloring for presentations

### ⇧ Shift+Click - Permanent View Filters
- Creates reusable view filters for each MMI level
- Applies filters with graphics overrides to the active view
- **Use Case**: Long-term solution, filters can be reused across multiple views

## Official MMI Color Codes

All colors conform to the official [MMI-veilederen 2.0 standard](https://mmi-veilederen.no/?page_id=85):

| MMI Value | Description | RGB Color | Preview |
|-----------|-------------|-----------|---------|
| **000** | Tidligfase | (215, 50, 150) | 🟪 Magenta |
| **100** | Grunnlagsinformasjon | (190, 40, 35) | 🔴 Dark Red |
| **125** | Etablert konsept | (210, 75, 70) | 🔴 Red |
| **150** | Tverrfaglig kontrollert konsept | (225, 120, 115) | 🔴 Light Red |
| **175** | Valgt konsept | (240, 170, 170) | 🔴 Very Light Red |
| **200** | Ferdig konsept | (230, 150, 55) | 🟠 Orange |
| **225** | Etablert prinsipielle løsninger | (235, 175, 100) | 🟠 Light Orange |
| **250** | Tverrfaglig kontrollert prinsipielle løsninger | (240, 200, 140) | 🟠 Very Light Orange |
| **275** | Valgt prinsipielle løsninger | (245, 230, 215) | 🟠 Beige |
| **300** | Underlag for detaljering | (250, 240, 80) | 🟡 Yellow |
| **325** | Etablert detaljerte løsninger | (215, 205, 65) | 🟡 Darker Yellow |
| **350** | Tverrfaglig kontrollert detaljerte løsninger | (185, 175, 60) | 🟡 Olive Yellow |
| **375** | Detaljerte løsninger (anbud/bestilling) | (150, 150, 50) | 🟡 Olive |
| **400** | Arbeidsgrunnlag | (55, 130, 70) | 🟢 Dark Green |
| **425** | Etablert/utført | (75, 170, 90) | 🟢 Green |
| **450** | Kontrollert utførelse | (100, 195, 125) | 🟢 Light Green |
| **475** | Godkjent utførelse | (155, 215, 165) | 🟢 Very Light Green |
| **500** | Som bygget | (30, 70, 175) | 🔵 Blue |
| **600** | I drift | (175, 50, 205) | 🟣 Purple |

## How It Works

### Normal Click (script.py)
1. Scans all elements in the active view
2. Reads the MMI parameter value from each element
3. Assigns colors based on the **closest lower** MMI level
   - Example: Value `140` uses color from `125`
   - Example: Value `360` uses color from `350`
4. Applies color overrides using `SetElementOverrides()`

### Shift+Click (config.py)
1. Creates 19 view filters (one for each official MMI level)
2. Each filter uses an **exact match** rule (e.g., equals "400")
3. Filters are named systematically: `MMI_000_Tidligfase`, `MMI_400_Arbeidsgrunnlag`, etc.
4. Applies filters to the active view with graphics overrides
5. Filters remain in the model and can be reused

## Requirements

- **MMI Parameter**: Must be configured in MMI Settings
- **Parameter Type**: String/Text parameter
- **Value Format**: Three-digit format (e.g., "000", "125", "400")
- **View Type**: Works with all model views (not sheets or templates)

## Usage Instructions

### Normal Coloring (Temporary)
1. Open a model view
2. Click **Colorizer** button
3. Elements will be colored based on their MMI values
4. Click again to remove colors

### Create View Filters (Permanent)
1. Open a model view
2. Hold **Shift** and click **Colorizer** button
3. 19 view filters will be created in the model
4. Filters will be applied to the active view
5. Manage filters via Revit's **Visibility/Graphics** (VV/VG) dialog

## Filter Management

Once created, view filters can be:
- Applied to other views manually
- Edited in the Filters dialog (Manage → Filters)
- Modified in Visibility/Graphics settings
- Copied between projects
- Included in view templates

## Technical Details

### Color Override Settings
Both modes apply:
- **Projection line color**: Set to MMI color
- **Cut line color**: Set to MMI color
- **Surface foreground pattern**: Solid fill with MMI color
- **Cut foreground pattern**: Solid fill with MMI color

### Filter Logic
- **Filter Type**: `ParameterFilterElement`
- **Rule Type**: `FilterStringRule` with `Equals` comparison
- **Categories**: Applied to all filterable model categories
- **Case Sensitivity**: Enabled (values must match exactly)

### State Management
- Active state stored in pyRevit user config
- Colored view ID and element IDs cached for cleanup
- Filters persist in the model (not session-specific)

## Error Handling

The tool validates:
- MMI parameter existence and configuration
- Active view type (must be model view)
- View template status (cannot apply to templates)
- Parameter availability on elements
- Transaction success

## References

- [MMI-veilederen 2.0](https://mmi-veilederen.no/?page_id=85) - Official Norwegian MMI Guide
- [Grunnleggende MMI-nivåer](https://mmi-veilederen.no/?page_id=85) - Core MMI levels and color codes

## Version History

- **v2.0** - Updated to official MMI color codes (19 levels)
- **v1.0** - Initial release with 4 simplified color ranges

## Author

**Byggstyrning AB**

---

*This tool follows the official MMI standard and uses the exact RGB color codes specified in the MMI-veilederen.*

