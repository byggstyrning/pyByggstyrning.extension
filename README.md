![pyByggstyrning](pyByggstyrning.png)

# Overview

pyByggstyrning is a pyRevit extension with specialized tools for model-based construction workflows.

![pyBS tab](screenshot-tab.png)

# Tools

| Panel | Tools |
|-------|-------|
| **MMI** | Quick-set buttons for MMI values (200-475) |
| **MMI Settings** | Settings, Monitor, Colorizer |
| **View** | Color Elements, Reset Colors |
| **3D Zone** | Create 3D Zones from Rooms/Areas/Regions, Spatial Mappings, Write Mappings, Isolate |
| **MEP Spaces** | Create Spaces from Link, Update Spaces, Tag All Spaces |
| **StreamBIM** | Checklist Importer, Edit Configs, Run Everything |
| **Project Browser** | IFC Classification (Load Defaults, Quick Class, Class Crawler), Better Schedule |
| **Elements** | Weight Calc |
| **Documentation** | Load Family, Create References, Add to View |

# Extra Features

- **MMI in Modify tab** - Cloned on startup for quick access
- **Switchback API** - HTTP endpoint to select elements by ID (`http://localhost:48884/switchback/id/<element_id>`)
- **IFC Export Handler** - Automatic 3D Zone parameter mapping during export
- **Batch Importer** - Automated StreamBIM import via `pyrevit run`

# Installation

## 1. Install pyRevit

Download and install pyRevit from the [GitHub releases](https://github.com/pyrevitlabs/pyRevit/releases) or follow the [pyRevit Installation Guide](https://pyrevitlabs.notion.site/Install-pyRevit-98ca4359920a42c3af5c12a7c99a196d).

## 2. Install pyByggstyrning

### Using Extension Manager (Recommended)

1. Open Revit with pyRevit installed
2. Go to **pyRevit tab** → **Extensions**
3. Find **pyByggstyrning** in the list and click **Install Extension**

<img width="464" alt="Extension Manager" src="https://github.com/user-attachments/assets/af3110ea-9a77-44c2-881b-01068a334792" />

For more details, see the [pyRevit Extensions Installation Guide](https://pyrevitlabs.notion.site/Install-Extensions-0753ab78c0ce46149f962acc50892491).

### Using pyRevit CLI

```bash
pyrevit extend ui pyByggstyrning https://github.com/byggstyrning/pyByggstyrning.extension.git
```

## 3. Switchback Support (Optional)

```bash
pyrevit configs routes port 48884
```

# Video Demos

- [Color Elements](https://github.com/user-attachments/assets/4628d6dd-39a4-44ff-af8b-22d1dab4f7a7)
- [StreamBIM Checklist Importer](https://github.com/user-attachments/assets/bf7c74ad-be95-4c0f-a815-b736b2d31a3a)
- [Switchback](https://github.com/user-attachments/assets/ce8b4bb5-0ae3-49b4-9656-75130f450b33)

# Credits

- [Ehsan Iran-Nejad](https://github.com/eirannejad) - pyRevit
- [Icons8](https://icons8.com/) - Icons
- [Erik Frits](https://github.com/ErikFrits) - LearnRevitAPI
- [Giuseppe Dotto](https://github.com/GiuseppeDotto) - Filters/Views from pyM4B
- [khorn06](https://github.com/khorn06/extensible-storage-pyrevit) - Extensible storage library
