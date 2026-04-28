pyByggstyrning

# Overview

pyByggstyrning is a pyRevit extension designed to enhance Revit workflows with specialized tools for model-based construction.

pyBS tab

# Tools


| Panel               | Tools                                                                                                                    |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| **Development**     | Reload, Select Branch                                                                                                    |
| **Project Browser** | Better Schedule, Load Type Defaults, Quick Class, Class Crawler                                                          |
| **Documentation**   | Load Family, Create References, Add to View                                                                              |
| **StreamBIM**       | Checklist Importer, Edit Configs, Run Everything                                                                         |
| **MMI Settings**    | Settings, Monitor, Colorizer                                                                                             |
| **MMI**             | Quick-set buttons for MMI values 200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 450, 475                              |
| **Elements**        | Weight Calc, Family Elevations                                                                                           |
| **MEP Spaces**      | Create Spaces from Link, Update Spaces, Tag All Spaces                                                                   |
| **3D Zone**         | Edit Spatial Mappings, Create 3D Zones from Rooms / Areas / Regions, Mass from Regions, Write Mappings, Isolate 3D Zones |
| **View**            | Color Elements, Reset                                                                                                    |
| **Coordination**    | Clash Views                                                                                                              |


# Extra Features

- **MMI in Modify tab** — The MMI ribbon panel is cloned onto the **Modify** tab at startup. On pyRevit 6.x, cloning also runs from an **Idling** handler until the pyBS MMI panel exists, and runs again when a document is opened.
- **Switchback API** — HTTP endpoint for selecting elements by id is maintained in **pyValidator.extension** (avoids binding the same route from two extensions). Configure the listener port as needed, for example: `pyrevit configs routes port 48884`.
- **IFC export handler** — Registers automatic 3D Zone parameter mapping during IFC export (see `startup.py` / `lib/zone3d`).
- **Batch importer** — Automated StreamBIM import via `pyrevit run` where applicable.

# Installation

## 1. Install pyRevit

Download and install pyRevit from the [GitHub releases](https://github.com/pyrevitlabs/pyRevit/releases) or follow the [pyRevit Installation Guide](https://pyrevitlabs.notion.site/Install-pyRevit-98ca4359920a42c3af5c12a7c99a196d).

## 2. Install pyByggstyrning

### Using Extension Manager (Recommended)

1. Open Revit with pyRevit installed
2. Go to **pyRevit tab** → **Extensions**
3. Find **pyByggstyrning** in the list and click **Install Extension**



For more details, see the [pyRevit Extensions Installation Guide](https://pyrevitlabs.notion.site/Install-Extensions-0753ab78c0ce46149f962acc50892491).

### Using pyRevit CLI

```bash
pyrevit extend ui pyByggstyrning https://github.com/byggstyrning/pyByggstyrning.extension.git
```

## 3. Switchback Support (Optional)

If you use **pyValidator** (or another extension) for Switchback, set the routes port as documented for that extension, for example:

```bash
pyrevit configs routes port 48884
```

# Credits

- [Ehsan Iran-Nejad](https://github.com/eirannejad) - pyRevit
- [Icons8](https://icons8.com/) - Icons
- [Erik Frits](https://github.com/ErikFrits) - LearnRevitAPI
- [Giuseppe Dotto](https://github.com/GiuseppeDotto) - Filters/Views from pyM4B
- [khorn06](https://github.com/khorn06/extensible-storage-pyrevit) - Extensible storage library

