# Common UI Styles

This directory contains reusable WPF styles for all custom XAML-based UI in the extension.

## Overview

All custom UI windows now use a centralized style system (`CommonStyles.xaml`) that provides:
- Consistent look and feel across all tools
- Easy maintenance and updates
- No external dependencies (no Xceed toolkit required)
- Custom placeholder textboxes, busy indicators, and styled controls

## Usage

### In XAML Files

1. **Remove inline styles** - Don't define styles directly in XAML
2. **Reference common styles** - Use `Style="{StaticResource StyleName}"` on controls
3. **Load styles programmatically** - Styles are loaded in Python code (see below)

### In Python Scripts

Add this method to your WPFWindow class:

```python
def load_styles(self):
    """Load the common styles ResourceDictionary."""
    try:
        import styles
        styles.load_common_styles(self)
    except Exception as e:
        logger.warning("Could not load styles: {}".format(e))

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

Call `self.load_styles()` in your `__init__` method after `WPFWindow.__init__()`.
This will automatically load the correct styles (Light/Dark) based on the Revit setting.


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

The styles use a consistent color palette:
- **Accent**: #0078D4 (blue)
- **Success**: #4CAF50 (green)
- **Error**: #D32F2F (red)
- **Warning**: #FF9800 (orange)
- **Border**: #CCCCCC (light gray)
- **Background Light**: #F5F5F5
- **Text**: #333333 (dark gray)

## Migration Checklist

When updating existing XAML files:

- [ ] Remove inline `<Style>` definitions from `<Window.Resources>`
- [ ] Add comment: `<!-- Styles will be loaded programmatically -->`
- [ ] Replace TextBox with PlaceholderTextBox style where needed
- [ ] Apply button styles (StandardButtonStyle, SecondaryButtonStyle, etc.)
- [ ] Apply text block styles (HeaderTextBlockStyle, LabelTextBlockStyle, etc.)
- [ ] Add busy overlay to Grid if needed
- [ ] Add `load_styles()` method to Python script
- [ ] Add `set_busy()` method if using busy indicator
- [ ] Call `self.load_styles()` in `__init__` after WPFWindow initialization

## Files Using Common Styles

- ✅ ChecklistImporter.xaml
- ✅ ConfigEditor.xaml
- ✅ Generate3DViewReferencesWindow.xaml
- ✅ ColorElementsWindow.xaml

All custom XAML-based UI should use these styles for consistency.
