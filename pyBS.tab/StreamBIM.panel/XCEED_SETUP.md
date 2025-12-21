# Xceed WPF Toolkit Setup

This extension uses the Xceed Extended WPF Toolkit (Community Edition) to enhance the UI with modern controls like WatermarkTextBox and BusyIndicator.

## Installation

1. Download the Xceed Extended WPF Toolkit Community Edition from:
   https://github.com/xceedsoftware/wpftoolkit

2. Extract the `Xceed.Wpf.Toolkit.dll` file from the downloaded package.

3. Place the DLL in one of the following locations:
   - **Recommended**: Copy `Xceed.Wpf.Toolkit.dll` to the `lib/` folder in your extension root directory
   - **Alternative**: Install it globally in the GAC (Global Assembly Cache) if you have admin rights

## Location

The extension will automatically look for the DLL in:
- `lib/Xceed.Wpf.Toolkit.dll` (relative to extension root)

If the DLL is not found, the extension will still work but without the enhanced UI features (it will fall back to standard WPF controls).

## Features Enabled

With Xceed Toolkit installed, the following enhancements are available:

- **WatermarkTextBox**: Placeholder hints in text fields (username, server URL, search box)
- **BusyIndicator**: Animated loading indicator during operations (login, data retrieval, imports)

## Troubleshooting

If you see warnings in the log about Xceed Toolkit not being found:
- Verify the DLL is in the correct location (`lib/Xceed.Wpf.Toolkit.dll`)
- Check that the DLL file name matches exactly: `Xceed.Wpf.Toolkit.dll`
- Ensure the DLL version is compatible with .NET Framework 4.x (required for Revit)

The extension will continue to function without the toolkit, but with standard WPF controls instead of the enhanced ones.
