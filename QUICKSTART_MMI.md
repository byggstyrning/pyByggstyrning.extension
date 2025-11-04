# MMI Quick Start Guide

This guide will walk you through installing pyRevit, the pyByggstyrning extension, and setting up the MMI (Model Maturity Index) tools for model-based construction workflows.

## Table of Contents
1. [Install pyRevit](#1-install-pyrevit)
2. [Install pyByggstyrning Extension](#2-install-pybyggstyrning-extension)
3. [Setting Up MMI Parameters](#3-setting-up-mmi-parameters)
4. [Setting MMI Values on Elements](#4-setting-mmi-values-on-elements)
5. [MMI Settings Configuration](#5-mmi-settings-configuration)
6. [Using MMI Monitor](#6-using-mmi-monitor)

---

## 1. Install pyRevit

pyRevit is a free and open-source plugin framework for Revit that allows you to run custom Python scripts and extensions.

### Installation Steps:

1. **Download pyRevit installer** from the official website or GitHub repository
2. **Run the installer** and follow the on-screen instructions
3. **Restart Revit** after installation completes

For detailed installation instructions, visit:
**[pyRevit Installation Guide](https://pyrevitlabs.notion.site/Install-pyRevit-98ca4359920a42c3af5c12a7c99a196d)**

![Screenshot: pyRevit Installation]
*[PLACEHOLDER: Screenshot of pyRevit installer]*

---

## 2. Install pyByggstyrning Extension

Once pyRevit is installed, you can add the pyByggstyrning extension which includes the MMI tools.

### Method 1: Using pyRevit Extension Manager (Recommended)

This is the easiest method using the built-in pyRevit GUI.

1. **Open Revit** with pyRevit installed

2. **Open pyRevit Extension Manager:**
   - Click on the **pyRevit tab** in the Revit ribbon
   - Click the **"Extensions"** button (or use the pyRevit menu)
   
   ![Screenshot: pyRevit Extensions Button]
   *[PLACEHOLDER: Screenshot highlighting the Extensions button in pyRevit tab]*

3. **Add the extension source:**
   - In the Extensions window, look for the pyByggstyrning extension in the list
   - If not found, click on the **gear icon** or **settings** to add a custom extension source
   - Add the GitHub repository URL: `https://github.com/byggstyrning/pyByggstyrning.extension.git`
   
   ![Screenshot: pyRevit Extension Manager]
   *[PLACEHOLDER: Screenshot of pyRevit Extensions manager window]*

4. **Install the extension:**
   - Find **pyByggstyrning** in the extensions list
   - Click the **Install** or **Toggle** button to enable it
   
   ![Screenshot: Installing Extension]
   *[PLACEHOLDER: Screenshot showing pyByggstyrning extension being installed]*

5. **Restart Revit** to load the extension

For detailed extension installation instructions, visit:
**[pyRevit Extensions Installation Guide](https://pyrevitlabs.notion.site/Install-Extensions-0753ab78c0ce46149f962acc50892491)**

### Method 2: Using Command Line

1. **Open Command Prompt** (Windows Key + R, type `cmd`, press Enter)
2. **Run the extension install command:**
   ```
   pyrevit extend ui pyByggstyrning https://github.com/byggstyrning/pyByggstyrning.extension.git
   ```
3. **Restart Revit** to load the extension

![Screenshot: Command Prompt Installation]
*[PLACEHOLDER: Screenshot of command prompt with extension install command]*

### Verify Installation

After restarting Revit, you should see a new **"pyBS"** tab in the Revit ribbon with several panels including:
- **MMI Panel** - Quick access buttons for setting MMI values (200, 225, 250, etc.)
- **MMI Settings Panel** - Configuration tools for MMI parameters and monitoring

![Screenshot: pyBS Tab in Revit]
*[PLACEHOLDER: Screenshot of pyBS tab showing all panels]*

---

## 3. Setting Up MMI Parameters

Before using the MMI tools, you need to configure which parameter in your Revit project will store the MMI values.

### Prerequisites

Your Revit project should have a **shared parameter** or **project parameter** (Instance parameter, Text type) that will be used to store MMI values. Suggested parameter name: `MMI`

### Configuration Steps

1. **Open your Revit project**
2. **Navigate to the pyBS tab** in the Revit ribbon
3. **Go to the MMI Settings panel**
4. **Click the "Settings" button**

   ![Screenshot: MMI Settings Button]
   *[PLACEHOLDER: Screenshot highlighting the Settings button in MMI Settings panel]*

5. **In the MMI Settings dialog:**
   - Click **"Set MMI Parameter"**
   
   ![Screenshot: MMI Settings Dialog]
   *[PLACEHOLDER: Screenshot of MMI Settings dialog with "Set MMI Parameter" option]*

6. **Select your MMI parameter:**
   - A list of available instance parameters will appear
   - Select the parameter you want to use for MMI values
   - Click OK

   ![Screenshot: Parameter Selection]
   *[PLACEHOLDER: Screenshot of parameter selection dialog]*

7. **Confirmation:**
   - You'll see a balloon notification confirming the parameter has been set
   - The parameter setting is stored in the project file using Extensible Storage

   ![Screenshot: Confirmation Balloon]
   *[PLACEHOLDER: Screenshot of confirmation balloon notification]*

---

## 4. Setting MMI Values on Elements

Once your MMI parameter is configured, you can quickly assign MMI values to elements in your model.

### What is MMI?

MMI (Model Maturity Index) is a Norwegian standard indicating the level of detail and approval status of BIM elements. Values range from 200-475 in increments of 25:

- **200** - Finished concept
- **250** - Finished detailed design
- **300** - As-built model
- **400+** - Approved/locked elements

For complete definitions, visit: [MMI Veilederen](https://mmi-veilederen.no/?page_id=85)

### Using MMI Buttons

![Screenshot: MMI Panel]
*[PLACEHOLDER: Screenshot of MMI Panel showing all MMI value buttons]*

**Steps:**

1. **Select elements** in your Revit view

   ![Screenshot: Selected Elements]
   *[PLACEHOLDER: Screenshot showing multiple elements selected in Revit]*

2. **Click the MMI button** for the desired value (200, 250, 300, etc.)
   - Buttons are organized by hundreds (200 series, 300 series, 400 series)

   ![Screenshot: Clicking MMI Button]
   *[PLACEHOLDER: Screenshot highlighting an MMI button being clicked]*

3. **Values are applied** to all selected elements instantly

   ![Screenshot: MMI Values Applied]
   *[PLACEHOLDER: Screenshot showing elements with updated MMI parameter values]*

**Quick Tip:** The MMI panel also appears in the **Modify tab** when elements are selected for quick access.

![Screenshot: MMI in Modify Tab]
*[PLACEHOLDER: Screenshot of MMI panel appearing in Modify tab]*

---

## 5. MMI Settings Configuration (Optional)

Configure automated monitoring features for quality control.

### Available Settings

1. **Validate MMI** - Auto-corrects MMI value formats (e.g., "2 00" → "200")
2. **Pin elements >=400** - Auto-pins high-status elements to prevent modification
3. **Warn when moving elements >400** - Shows warnings when moving approved elements
4. **Check MMI after sync** - Verifies MMI values after sync with central

### Configuration Steps

1. Click **"Settings"** button in the MMI Settings panel
2. Check the boxes for features you want to enable
3. Click **"Save Config"** to save settings

![Screenshot: MMI Configuration Options]
*[PLACEHOLDER: Screenshot showing all configuration checkboxes]*

**Recommended:** Enable Validate MMI, Pin elements >=400, and Check MMI after sync for most projects.

---

## 6. Using MMI Monitor (Optional)

The MMI Monitor automatically enforces the rules you configured in MMI Settings as you work.

### Starting the Monitor

1. Click the **"Monitor"** button in the MMI Settings panel to toggle it on

![Screenshot: Monitor Button - OFF State]
*[PLACEHOLDER: Screenshot of Monitor button in OFF state]*

2. The button icon changes to green and shows active features

![Screenshot: Monitor Button - ON State]
*[PLACEHOLDER: Screenshot of Monitor button in ON state]*

![Screenshot: Monitor Active Notification]
*[PLACEHOLDER: Screenshot of balloon notification showing "Monitor activated"]*

### What the Monitor Does

The monitor runs in the background and:
- **Validates MMI formats** - Auto-corrects spacing and formatting issues
- **Pins high-status elements** - Prevents accidental changes to elements with MMI >=400
- **Warns on moves** - Notifies when approved elements are moved
- **Checks after sync** - Verifies MMI values after syncing with central

![Screenshot: Validation Notification]
*[PLACEHOLDER: Screenshot of MMI validation correction notification]*

![Screenshot: Pinning Notification]
*[PLACEHOLDER: Screenshot showing elements pinned notification]*

### Stopping the Monitor

Click the **"Monitor"** button again to toggle it off. The monitor state persists across Revit sessions.

---

## Troubleshooting

### MMI Buttons Don't Work
- **Check:** Is the MMI parameter configured? Go to Settings and set your MMI parameter
- **Check:** Are elements selected? MMI buttons require a selection
- **Check:** Does the selected element support the parameter? Some system families may not have access to all parameters

### Monitor Doesn't Start
- **Check:** Is the MMI parameter configured in Settings?
- **Check:** Are any monitor features enabled in Settings?
- **Check:** Check the pyRevit log for error messages (pyRevit tab > Settings > Open Log Folder)

### Parameter Not Saving
- **Check:** Is the parameter an Instance parameter (not Type parameter)?
- **Check:** Is the parameter Text type (not Integer)?
- **Check:** Do you have write access to the project?

### Elements Not Pinning Automatically
- **Check:** Is "Pin elements >=400" enabled in MMI Settings?
- **Check:** Is the MMI Monitor running (button should show ON state)?
- **Check:** Are the elements already pinned?
- **Check:** Does the element type support pinning? (Some elements like rooms cannot be pinned)

---

## Support and Resources

- **GitHub Repository:** [pyByggstyrning Extension](https://github.com/byggstyrning/pyByggstyrning.extension)
- **MMI Standard:** [MMI Veilederen](https://mmi-veilederen.no/)
- **pyRevit Documentation:** [pyRevit Labs](https://pyrevitlabs.notion.site/)

For issues or feature requests, please visit the GitHub repository and create an issue.

---

**Last Updated:** November 2025  
**Extension Version:** 1.x  
**Author:** Byggstyrning AB

