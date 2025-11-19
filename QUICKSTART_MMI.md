# MMI Quick Start Guide

This guide will walk you through installing pyRevit, the pyByggstyrning extension, and setting up the MMI (Model Maturity Index) tools for model-based construction workflows.

---

## 1. Install pyRevit

pyRevit is a free and open-source plugin framework for Revit that allows you to run custom Python scripts and extensions.

### Installation Steps:

1. **Download pyRevit installer** from the official website or **[GitHub repository](https://github.com/pyrevitlabs/pyRevit/releases)**
2. **Run the installer** and follow the on-screen instructions
3. **Restart Revit** after installation completes

For detailed installation instructions, visit:
**[pyRevit Installation Guide](https://pyrevitlabs.notion.site/Install-pyRevit-98ca4359920a42c3af5c12a7c99a196d)**

---

## 2. Install pyByggstyrning Extension

Once pyRevit is installed, you can add the pyByggstyrning extension which includes the MMI tools.

### Method 1: Using pyRevit Extension Manager (Recommended)

This is the easiest method using the built-in pyRevit GUI.

1. **Open Revit** with pyRevit installed

2. **Open pyRevit Extension Manager:**
   - Click on the **pyRevit tab** in the Revit ribbon
   - Click the **"Extensions"** button (or use the pyRevit menu)
   
<img width="579" height="303" alt="image" src="https://github.com/user-attachments/assets/bb16c156-7901-4844-931c-2eee9079190b" />

3. **Install the extension:**
   - Find and Select **pyByggstyrning** in the extensions list
   - Click the **Install Extension** to install it it
   
<img width="464" height="715" alt="image" src="https://github.com/user-attachments/assets/af3110ea-9a77-44c2-881b-01068a334792" />


4. **Restart Revit** to load the extension

For detailed extension installation instructions, visit:
**[pyRevit Extensions Installation Guide](https://pyrevitlabs.notion.site/Install-Extensions-0753ab78c0ce46149f962acc50892491)**

### Method 2: Using Command Line

1. **Open Command Prompt** (Windows Key + R, type `cmd`, press Enter)
2. **Run the extension install command:**
   ```
   pyrevit extend ui pyByggstyrning https://github.com/byggstyrning/pyByggstyrning.extension.git
   ```
3. **Restart Revit** to load the extension

### Verify Installation

After restarting Revit, you should see a new **"pyBS"** tab in the Revit ribbon with several panels including:
- **MMI Panel** - Quick access buttons for setting MMI values (200, 225, 250, etc.)
- **MMI Settings Panel** - Configuration tools for MMI parameters and monitoring

<img width="914" height="224" alt="image" src="https://github.com/user-attachments/assets/75f5374f-ebcd-490d-b303-071c7375466a" />

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

<img width="600" height="181" alt="image" src="https://github.com/user-attachments/assets/e7e6189e-4f9f-48a0-9174-2674bd89c172" />

5. **Select your MMI parameter:**
   - Click **"Set MMI Parameter"**
   - A list of available instance parameters will appear
   - Select the parameter you want to use for MMI values
   - Click OK
   - <img width="430" height="185" alt="image" src="https://github.com/user-attachments/assets/e558e38e-bd60-4943-8c28-517349610c3a" />


6. **Confirmation:**
   - You'll see a balloon notification confirming the parameter has been set
   - The parameter setting is stored in the project file using Extensible Storage
   - <img width="318" height="118" alt="image" src="https://github.com/user-attachments/assets/a84d0943-9803-4d44-9fc7-a4da848f7277" />


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

**Steps:**

1. **Select elements** in your Revit view
2. **Click the MMI button** for the desired value (200, 250, 300, etc.)
   - Buttons are organized by hundreds (200 series, 300 series, 400 series)
   - <img width="367" height="468" alt="image" src="https://github.com/user-attachments/assets/5add79fa-53c0-418e-954e-9826a93dcbf6" />
3. **Values are applied** to all selected elements instantly
<img width="648" height="411" alt="image" src="https://github.com/user-attachments/assets/deb5b8e2-1c55-4d40-a877-4d050940afaf" />


**Quick Tip:** The MMI panel also appears in the **Modify tab** when elements are selected for quick access.

---

## 5. MMI Settings Configuration

Configure automated monitoring features for quality control.

### Available Settings

1. **Validate MMI** - Auto-corrects MMI value formats (e.g., "2 00" → "200")
2. **Pin elements >=400** - Auto-pins high-status elements to prevent modification
3. **Warn when moving elements >=425** - Shows warnings when moving approved elements
4. **Check MMI after sync** - Verifies MMI values after sync with central

### Configuration Steps

1. Click **"Settings"** button in the MMI Settings panel
2. Check the boxes for features you want to enable
3. Click **"Save Config"** to save settings

<img width="600" height="181" alt="image" src="https://github.com/user-attachments/assets/105a62c2-fbf7-4287-9197-69ca9325ccd9" />

**Recommended:** Enable Validate MMI, Pin elements >=400, and Check MMI after sync for most projects.

---

## 6. Using MMI Monitor
The MMI Monitor automatically enforces the rules you configured in MMI Settings as you work.

### Starting the Monitor

1. Click the **"Monitor"** button in the MMI Settings panel to toggle it on

<img width="87" height="106" alt="image" src="https://github.com/user-attachments/assets/213e8ddb-c15b-4a26-bce4-8e66393f2a5b" />


2. The button icon changes to orange when active 

<img width="83" height="26" alt="image" src="https://github.com/user-attachments/assets/e3d3777e-997f-4665-898b-43b28b50ec0a" />

3. Bubble appears to display selected settings
4. 
<img width="287" height="193" alt="image" src="https://github.com/user-attachments/assets/0399d25d-d3f4-418f-8ada-243346165440" />

### What the Monitor Does

The monitor runs in the background and:
- **Validates MMI formats** - Auto-corrects spacing and formatting issues
- **Pins high-status elements** - Prevents accidental changes to elements with MMI >=400
- **Warns on moves** - Notifies when approved elements are moved
- **Checks after sync** - Verifies MMI values after syncing with central

<img width="526" height="389" alt="image" src="https://github.com/user-attachments/assets/fd3ef22d-6246-4b7c-807f-87f2a9adb8f0" />


### Stopping the Monitor

Click the **"Monitor"** button again to toggle it off. The monitor state persists across Revit sessions.


## Support and Resources

- **GitHub Repository:** [pyByggstyrning Extension](https://github.com/byggstyrning/pyByggstyrning.extension)
- **MMI Standard:** [MMI Veilederen](https://mmi-veilederen.no/)
- **pyRevit Documentation:** [pyRevit Labs](https://pyrevitlabs.notion.site/)

For issues or feature requests, please visit the GitHub repository and create an issue.

---

**Last Updated:** November 2025  
**Author:** Byggstyrning AB
