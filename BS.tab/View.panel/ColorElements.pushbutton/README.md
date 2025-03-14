# PyRevit Colorizer

A tool for colorizing elements in Revit based on parameter values.

## Features

- Colorize elements based on parameter values
- Support for both instance and type parameters
- Pick custom colors for each value
- Apply and reset colors with a single click
- Maintain color overrides across views (optional)
- Modern UI with easy-to-use controls

## Usage

1. Select a category from the list
2. Choose between Instance or Type parameters
3. Select a parameter from the dropdown
4. The tool will automatically populate values and assign random colors
5. Double-click any value to customize its color
6. Click "Apply Colors" to colorize elements in the current view
7. Click "Reset Colors" to remove all color overrides

## Options

- **Override Projection/Cut Lines**: Apply colors to lines as well as surfaces
- **Show Elements in Properties**: Show colored elements in the Properties palette
- **Keep color overrides in view when exiting**: Preserve colors when closing the tool

## Requirements

- Revit 2019 or newer
- pyRevit 4.8 or newer

## License

This tool is part of the pyByggstyrning extension. 