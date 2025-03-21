![pyByggstyrning](pyByggstyrning.png)

# Overview
pyByggstyrning is a pyRevit extension designed to enhance Revit workflows with some specialized tools for use in model based construction.

- **MMI Panel**: Tools for quickly setting MMI statuses on selected elements.
- **Project Browser Panel**: Enhanced tools for navigating and managing views and filters, credits to Giuseppe Dotto.
- **StreamBIM Panel**: Tools for StreamBIM projects for importing checklist values from site to Revit. 
- **View Panel**: Color Elements tool that can help you quickly identify and select elements based on parameter values.

![pyBS tab](screenshot-tab.png)

# Extra features

- **Quick access to MMI in the Modify tab**: In the [startup.py](https://github.com/byggstyrning/pyByggstyrning.extension/blob/master/startup.py), we're cloning the MMI Panel from our panel into the Modify tab in Revit for easy access when selecting elements.
- **StreamBIM API utilities**: Utilities for interacting with StreamBIM API, including authentication, project management, and checklist item retrieval. [StreamBIM utilities](https://github.com/byggstyrning/pyByggstyrning.extension/tree/master/lib/streambim).
- **Batch Importer tool**: The StreamBIM Checklist import can be automated using the 'pyrevit run' command to start revit and execute the 'Run Everything' script, this can be schedules using the .bat located in the tool folder. For more, checkout the [Batch Importer Tool](https://github.com/byggstyrning/pyByggstyrning.extension/tree/master/pyBS.tab/StreamBIM.panel/Batch%20Importer%20Tool).

# Installation
To use the extension follow the steps:

1. Install pyRevit or make sure it's already installed
2. Add pyByggstyrning extension:
   - Open command prompt (Win + R) => cmd
   - Type following command: `pyrevit extend ui pyByggstyrning https://github.com/byggstyrning/pyByggstyrning.extension.git`
   - A tab pyBS should appear on the next start of Revit

Or via the built-in extension manager in pyRevit

# Credits
- [Ehsan Iran-Nejad](https://github.com/eirannejad) for developing pyRevit
- [Icons8](https://icons8.com/) for its sweet icons
- [Erik Frits](https://github.com/ErikFrits) and the [LearnRevitAPI](https://learnrevitapi.com/) course
- [Giuseppe Dotto](https://github.com/GiuseppeDotto) for the Filters and Views dropdowns from the [pyM4B extension](https://github.com/GiuseppeDotto/pyM4B.extension).
- [khorn06](https://github.com/khorn06/extensible-storage-pyrevit) for the extensible storage library
