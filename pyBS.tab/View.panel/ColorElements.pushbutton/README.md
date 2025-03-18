# Color Elements

A tool for colorizing elements in Revit based on parameter values, with support for multiple categories and parameter types.

## Features

- Colorize elements based on parameter values across multiple categories
- Support for both instance and type parameters
- Automatic color assignment with visually distinct color ranges
- Automatic color assignment based on predefined color schemas (.cschn files)
- Double-click to customize colors for specific values
- Apply and reset colors with a single click
- Maintain color overrides across views
- Select elements in Revit based on parameter values
- Modern UI with easy-to-use controls
- Automatic color scheme handling for rooms, spaces, and areas
- Support for various parameter types (Double, ElementId, Integer, String, Boolean)

## Usage

1. Select one or more categories from the list
2. Choose between Instance or Type parameters
3. Select a parameter from the dropdown (shows only parameters common to all selected categories)
4. The tool will automatically:
   - Populate values from the selected parameter
   - Assign visually distinct colors based on the number of values
   - Apply colors to elements in the current view
5. Double-click any value to customize its color using the system color picker
6. Click "Apply Colors" to update colors in the current view
7. Click "Reset Colors" to remove all color overrides
8. Use the "Show Elements" checkbox to enable/disable element selection in Revit

## Options

- **Override Projection/Cut Lines**: Apply colors to lines as well as surfaces
- **Show Elements in Properties**: Show colored elements in the Properties palette
- **Keep color overrides in view when exiting**: Colors are automatically reset when changing views or closing the tool

## Special Features

- Automatic handling of color schemes for rooms, spaces, and areas
- Smart color assignment based on value count:
  - Up to 5 values: Distinct colors (Red, Green, Blue, Gold, Purple)
  - Up to 20 values: Gradient between Red → Green → Blue
  - More than 20 values: HSV color wheel distribution
- Automatic sorting of values:
  - Numeric values are sorted numerically
  - Text values are sorted alphabetically
  - "None" values are always placed at the end
- Gray color (192,192,192) is automatically assigned to "None" values
- Maintains category selection when switching between views
- Automatic parameter refresh capability

## Requirements

- Revit 2019 or newer
- pyRevit 4.8 or newer

## License

This tool is part of the pyByggstyrning extension.

# Color Elements Tool

This tool allows you to colorize Revit elements based on parameter values. You can easily identify and distinguish elements by applying different colors based on their parameter values.

## Features

- Color elements in the current view by parameter values
- Support for both instance and type parameters
- Save and load color schemes
- Automatic color scheme application by parameter name
- Create view filters based on colors
- Multiple categories support

## Using Color Schemas

The tool supports automatic loading of color schemas. When you select a parameter, the tool will check if there's a matching color schema file in the `coloringschemas` folder, and if found, it will apply that schema automatically.

### Predefined Color Schemas

The following predefined color schemas are included:

1. **MMI.cschn**: For MMI (Miljøkartlegging) grading parameters, using standard colors:

   - A1: Green (good condition)
   - A2: Light green
   - B1: Yellow
   - B2: Light yellow
   - C1: Orange
   - C2: Light orange
   - D1: Red (hazardous)
   - D2: Light red
2. **NS_3451_Color.cschn**: For building elements according to Norwegian Standard NS 3451, with color coding for different building systems.

### Adding Your Own Color Schemas

You can save your own color schemas:

1. Select a parameter and adjust colors as needed
2. Click "Save/Load Color Scheme"
3. Choose "Save Color Scheme"
4. Name the file after the parameter name to enable automatic loading (e.g., for a parameter named "Phase", save as "Phase.cschn")

### Color Schema File Format

Each line in a color schema file follows this format:

```
value::RrGgBb
```

Where:

- `value` is the parameter value
- `r`, `g`, `b` are the RGB color values (0-255)

Example:

```
A1::R0G149B67
B1::R252G175B22
C1::R250G122B17
D1::R230G33B23
```

## Tips

- To quickly apply standard colors for known parameters, name them exactly as the schema files
- You can create multiple schemas for the same parameter with different filenames and load them as needed
- When loading a schema manually, you can choose to match by value or position in the list
