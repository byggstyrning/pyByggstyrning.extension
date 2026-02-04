# Common UI Styles

This directory contains reusable WPF styles for all custom XAML-based UI in the extension.

## Overview

All custom UI windows now use a centralized style system (`CommonStyles.xaml`) that provides:
- Consistent look and feel across all tools
- **Automatic dark mode support** - Follows Revit's UI theme setting
- Easy maintenance and updates
- No external dependencies (no Xceed toolkit required)
- Custom placeholder textboxes, busy indicators, and styled controls

## Dark Mode Support

The styles automatically detect and adapt to Revit's dark/light theme setting (Revit 2024+). This is handled by the `lib/styles/__init__.py` module.

### How It Works

1. When styles are loaded, the module detects Revit's current UI theme
2. Color resources are automatically updated to match the theme
3. Dark mode uses dark grey backgrounds (#2D2D2D) with light text (#E0E0E0)
4. Accent colors (yellow #ffbb00) remain consistent across themes

### Quick Start

```python
from lib.styles import load_styles_to_window

class MyWindow(WPFWindow):
    def __init__(self):
        # Initialize window first
        WPFWindow.__init__(self, xaml_file)
        
        # Load styles with automatic theme detection AFTER WPFWindow.__init__
        # Styles are window-scoped and will NOT affect Revit's UI
        load_styles_to_window(self)
```

### Force a Specific Theme

```python
# Force dark mode regardless of Revit setting
load_styles_to_window(self, force_theme='dark')

# Force light mode
load_styles_to_window(self, force_theme='light')
```

### Theme Detection API

```python
from lib.styles import get_revit_theme, is_dark_theme

# Get current theme as string ('dark' or 'light')
theme = get_revit_theme()

# Check if dark theme is active
if is_dark_theme():
    # Dark mode specific logic
    pass
```

## Usage

### In XAML Files

1. **Remove inline styles** - Don't define styles directly in XAML
2. **Reference common styles** - Use `Style="{DynamicResource StyleName}"` on controls (must use DynamicResource since styles load after XAML parsing)
3. **Load styles programmatically** - Styles are loaded in Python code (see below)

### In Python Scripts

Use the `load_styles_to_window()` function from `lib.styles`:

```python
from lib.styles import load_styles_to_window

class MyWindow(WPFWindow):
    def __init__(self):
        # Initialize window first
        WPFWindow.__init__(self, xaml_file)
        
        # Load styles into window Resources (window-scoped, isolated from Revit UI)
        load_styles_to_window(self)
```

**Important Notes:**
- Styles must be loaded AFTER `WPFWindow.__init__()` because the window must exist first
- Styles are window-scoped and will NOT affect Revit's UI
- Use `DynamicResource` (not `StaticResource`) in XAML since styles load after parsing

### Busy Indicator Helper Method

Add this helper method to your window class:

```python
def set_busy(self, is_busy, message="Loading..."):
    """Show or hide the busy overlay indicator."""
    try:
        if is_busy:
            self.busyOverlay.Visibility = Visibility.Visible
            self.busyTextBlock.Text = message
        else:
            self.busyOverlay.Visibility = Visibility.Collapsed
    except Exception as e:
        logger.debug("Error setting busy indicator: {}".format(str(e)))
```

## Available Styles

### Buttons

- **StandardButtonStyle** - Primary action button (blue)
- **SecondaryButtonStyle** - Secondary action button (gray)
- **SuccessButtonStyle** - Success/apply actions (green)
- **DangerButtonStyle** - Delete/destructive actions (red)

### TextBlocks

- **HeaderTextBlockStyle** - Main window title (16pt, bold)
- **SubheaderTextBlockStyle** - Section headers (14pt, semi-bold)
- **LabelTextBlockStyle** - Form labels
- **BodyTextBlockStyle** - Body text (12pt, wraps)
- **SecondaryTextBlockStyle** - Secondary/helper text (italic)
- **StatusTextBlockStyle** - Status bar text

### Input Controls

- **PlaceholderTextBoxStyle** - TextBox with placeholder placeholder (use `Tag` property for placeholder text)
- **StandardComboBoxStyle** - Styled ComboBox
- **StandardCheckBoxStyle** - Styled CheckBox
- **StandardRadioButtonStyle** - Styled RadioButton
- **ToggleSwitchStyle** - Toggle switch style for CheckBox

### Data Display

- **StandardDataGridStyle** - Styled DataGrid with alternating rows
- **StandardListBoxStyle** - Styled ListBox
- **StandardGroupBoxStyle** - Styled GroupBox

### Layout

- **PanelBorderStyle** - Border for content panels
- **HeaderBorderStyle** - Border for section headers
- **StandardStatusBarStyle** - Status bar styling

### Busy Indicator

Add this to your XAML Grid:

```xml
<Grid>
    <!-- Busy Indicator Overlay -->
    <Border x:Name="busyOverlay" Style="{StaticResource BusyOverlayStyle}">
        <StackPanel Style="{StaticResource BusyContentStyle}">
            <ProgressBar x:Name="busyProgressBar" Style="{StaticResource BusyProgressBarStyle}"/>
            <TextBlock x:Name="busyTextBlock" Style="{StaticResource BusyTextBlockStyle}" Text="Loading..."/>
        </StackPanel>
    </Border>
    
    <!-- Your content here -->
</Grid>
```

Then use `self.set_busy(True, "Loading...")` and `self.set_busy(False)` in Python.

## Examples

### Placeholder TextBox

```xml
<TextBox x:Name="usernameTextBox" 
         Style="{StaticResource PlaceholderTextBoxStyle}" 
         Tag="Enter username"/>
```

### Styled Button

```xml
<Button Content="Save" Style="{StaticResource StandardButtonStyle}"/>
<Button Content="Cancel" Style="{StaticResource SecondaryButtonStyle}"/>
<Button Content="Delete" Style="{StaticResource DangerButtonStyle}"/>
```

### Header Text

```xml
<TextBlock Text="My Tool" Style="{StaticResource HeaderTextBlockStyle}"/>
<TextBlock Text="Section Title" Style="{StaticResource SubheaderTextBlockStyle}"/>
```

## Color Palette

### Light Theme (Default)

| Color | Hex | Usage |
|-------|-----|-------|
| Accent | #ffbb00 | Primary buttons, highlights |
| Success | #4CAF50 | Success actions, toggles |
| Error | #D32F2F | Danger buttons, errors |
| Warning | #FF9800 | Warning indicators |
| Border | #CCCCCC | Control borders |
| Background Light | #F5F5F5 | Panels, headers |
| Background Lighter | #F0F0F0 | Secondary backgrounds |
| Text | #333333 | Primary text |
| Text Secondary | #666666 | Secondary text |
| Window Background | #FFFFFF | Window/control backgrounds |

### Dark Theme

| Color | Hex | Usage |
|-------|-----|-------|
| Accent | #ffbb00 | Primary buttons, highlights (same) |
| Success | #66BB6A | Success actions (lighter) |
| Error | #EF5350 | Danger buttons (lighter) |
| Warning | #FFA726 | Warning indicators (lighter) |
| Border | #555555 | Control borders |
| Background Light | #3C3C3C | Panels, headers |
| Background Lighter | #454545 | Secondary backgrounds |
| Text | #E0E0E0 | Primary text (light) |
| Text Secondary | #B0B0B0 | Secondary text |
| Window Background | #2D2D2D | Window/control backgrounds |

## Migration Checklist

When updating existing XAML files:

- [ ] Remove inline `<Style>` definitions from `<Window.Resources>`
- [ ] Add comment: `<!-- Styles will be loaded programmatically -->`
- [ ] Replace TextBox with PlaceholderTextBox style where needed
- [ ] Apply button styles (StandardButtonStyle, SecondaryButtonStyle, etc.)
- [ ] Apply text block styles (HeaderTextBlockStyle, LabelTextBlockStyle, etc.)
- [ ] Add busy overlay to Grid if needed
- [ ] Import `load_styles_to_window` from `lib.styles`
- [ ] Call `load_styles_to_window(self)` in `__init__` AFTER WPFWindow initialization
- [ ] Add `set_busy()` method if using busy indicator

## Files Using Common Styles

- ✅ ChecklistImporter.xaml
- ✅ ConfigEditor.xaml
- ✅ Generate3DViewReferencesWindow.xaml
- ✅ ColorElementsWindow.xaml

All custom XAML-based UI should use these styles for consistency.
