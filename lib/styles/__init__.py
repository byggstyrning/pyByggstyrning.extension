# Styles module for reusable WPF UI styles
"""
This module provides reusable WPF styles for PyRevit extensions.
Supports automatic dark mode detection based on Revit's UI theme.

Usage in Python:
    from lib.styles import ensure_styles_loaded, get_revit_theme, is_dark_theme
    
    # Load styles with automatic theme detection
    ensure_styles_loaded()
"""

import os
import os.path as op

# Color palettes for light and dark themes
LIGHT_THEME_COLORS = {
    'PlaceholderForegroundColor': '#999999',
    'BusyOverlayColor': '#80000000',
    'AccentColor': '#ffbb00',
    'AccentHoverColor': '#e6a800',
    'AccentPressedColor': '#cc9900',
    'ErrorColor': '#D32F2F',
    'SuccessColor': '#4CAF50',
    'WarningColor': '#FF9800',
    'BorderColor': '#CCCCCC',
    'BackgroundLightColor': '#F5F5F5',
    'BackgroundLighterColor': '#F0F0F0',
    'TextColor': '#333333',
    'TextSecondaryColor': '#666666',
    'TextLightColor': '#999999',
    'DisabledColor': '#CCCCCC',
    'DisabledTextColor': '#666666',
    # Additional colors used in styles
    'WindowBackgroundColor': '#FFFFFF',
    'ControlBackgroundColor': '#FFFFFF',
    'PopupBackgroundColor': '#FFFFFF',
    'DataGridRowBackgroundColor': '#FFFFFF',
    'ArrowColor': '#2196F3',  # Blue for light mode
    'ColoredButtonTextColor': '#000000',  # Black text for colored buttons (orange/green/red)
}

DARK_THEME_COLORS = {
    'PlaceholderForegroundColor': '#808080',
    'BusyOverlayColor': '#80000000',
    'AccentColor': '#ffbb00',  # Keep accent color consistent
    'AccentHoverColor': '#e6a800',
    'AccentPressedColor': '#cc9900',
    'ErrorColor': '#EF5350',  # Slightly lighter for dark bg
    'SuccessColor': '#66BB6A',  # Slightly lighter for dark bg
    'WarningColor': '#FFA726',  # Slightly lighter for dark bg
    'BorderColor': '#555555',
    'BackgroundLightColor': '#3C3C3C',  # Dark grey
    'BackgroundLighterColor': '#454545',  # Slightly lighter dark grey
    'TextColor': '#E0E0E0',  # Light text for dark bg
    'TextSecondaryColor': '#B0B0B0',
    'TextLightColor': '#808080',
    'DisabledColor': '#555555',
    'DisabledTextColor': '#808080',
    # Additional colors used in styles
    'WindowBackgroundColor': '#2D2D2D',  # Dark grey window background
    'ControlBackgroundColor': '#3C3C3C',  # Control background
    'PopupBackgroundColor': '#383838',  # Popup/dropdown background
    'DataGridRowBackgroundColor': '#2D2D2D',
    'ArrowColor': '#64B5F6',  # Light blue for dark mode
    'ColoredButtonTextColor': '#000000',  # Pure black text for colored buttons in dark mode (ensures contrast against bright colors)
}


def get_styles_path():
    """Get the absolute path to the styles directory."""
    current_dir = op.dirname(__file__)
    return current_dir


def get_common_styles_path():
    """Get the absolute path to CommonStyles.xaml."""
    return op.join(get_styles_path(), 'CommonStyles.xaml')


def get_revit_theme():
    """
    Detect Revit's current UI theme.
    
    Returns:
        str: 'dark' or 'light' based on Revit's theme setting.
             Defaults to 'light' if detection fails.
    """
    try:
        # Try Revit 2024+ API first (UIThemeManager)
        import clr
        clr.AddReference('RevitAPIUI')
        from Autodesk.Revit.UI import UIThemeManager
        
        # UIThemeManager.CurrentTheme returns UITheme enum
        # UITheme.Dark = 0, UITheme.Light = 1
        current_theme = UIThemeManager.CurrentTheme
        
        # Check if it's dark theme
        if hasattr(current_theme, 'value__'):
            # Enum value: 0 = Dark, 1 = Light
            result = 'dark' if current_theme.value__ == 0 else 'light'
            return result
        else:
            # Try string comparison
            theme_str = str(current_theme).lower()
            result = 'dark' if 'dark' in theme_str else 'light'
            return result
            
    except Exception as e:
        pass
    
    try:
        # Fallback: Try to detect from Revit application colors
        # This method checks the actual UI colors being used
        from Autodesk.Revit.UI import RevitCommandId, UIApplication
        from pyrevit import HOST_APP
        
        # Get the application's active ribbon background color
        # In dark mode, backgrounds are typically darker
        uiapp = HOST_APP.uiapp
        if uiapp:
            # Check window background - Revit 2019+ uses System.Windows
            from System.Windows import Application as WpfApp
            if WpfApp.Current is not None:
                # Try to detect from system theme
                try:
                    from System.Windows.Media import Color
                    # Get system window color
                    bg = WpfApp.Current.MainWindow
                    if bg is not None and hasattr(bg, 'Background'):
                        brush = bg.Background
                        if hasattr(brush, 'Color'):
                            color = brush.Color
                            # Calculate luminance - dark themes have low luminance
                            luminance = (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255
                            return 'dark' if luminance < 0.5 else 'light'
                except:
                    pass
    except Exception:
        pass
    
    # Default to light theme
    return 'light'


def is_dark_theme():
    """
    Check if Revit is currently using dark theme.
    
    Returns:
        bool: True if dark theme, False otherwise.
    """
    return get_revit_theme() == 'dark'


def get_theme_colors(theme=None):
    """
    Get the color palette for the specified theme.
    
    Args:
        theme: 'dark', 'light', or None (auto-detect)
    
    Returns:
        dict: Color palette dictionary
    """
    if theme is None:
        theme = get_revit_theme()
    
    colors = DARK_THEME_COLORS if theme == 'dark' else LIGHT_THEME_COLORS
    return colors


def apply_theme_to_resources(resources, theme=None):
    """
    Apply theme colors to a ResourceDictionary.
    
    This function updates the Color and SolidColorBrush resources
    in the provided ResourceDictionary to match the current theme.
    
    Args:
        resources: A WPF ResourceDictionary
        theme: 'dark', 'light', or None (auto-detect)
    
    Returns:
        bool: True if colors were applied successfully
    """
    try:
        from System.Windows.Media import Color, SolidColorBrush, ColorConverter
        
        colors = get_theme_colors(theme)
        
        # Update Color resources
        colors_applied = 0
        for color_key, color_value in colors.items():
            try:
                # Convert hex string to Color
                color = ColorConverter.ConvertFromString(color_value)
                resources[color_key] = color
                colors_applied += 1
            except Exception:
                pass
        
        # Update corresponding brush resources
        brush_mappings = {
            'PlaceholderForegroundBrush': 'PlaceholderForegroundColor',
            'BusyOverlayBrush': 'BusyOverlayColor',
            'AccentBrush': 'AccentColor',
            'AccentHoverBrush': 'AccentHoverColor',
            'AccentPressedBrush': 'AccentPressedColor',
            'ErrorBrush': 'ErrorColor',
            'SuccessBrush': 'SuccessColor',
            'WarningBrush': 'WarningColor',
            'BorderBrush': 'BorderColor',
            'BackgroundLightBrush': 'BackgroundLightColor',
            'BackgroundLighterBrush': 'BackgroundLighterColor',
            'TextBrush': 'TextColor',
            'TextSecondaryBrush': 'TextSecondaryColor',
            'TextLightBrush': 'TextLightColor',
            'DisabledBrush': 'DisabledColor',
            'DisabledTextBrush': 'DisabledTextColor',
            'WindowBackgroundBrush': 'WindowBackgroundColor',
            'ControlBackgroundBrush': 'ControlBackgroundColor',
            'PopupBackgroundBrush': 'PopupBackgroundColor',
            'DataGridRowBackgroundBrush': 'DataGridRowBackgroundColor',
            'ArrowBrush': 'ArrowColor',
            'ColoredButtonTextBrush': 'ColoredButtonTextColor',
        }
        
        brushes_applied = 0
        for brush_key, color_key in brush_mappings.items():
            try:
                if color_key in colors:
                    color = ColorConverter.ConvertFromString(colors[color_key])
                    brush = SolidColorBrush(color)
                    # Use indexer syntax to ensure it works with both ResourceDictionary and MergedDictionary
                    if brush_key in resources.Keys:
                        resources[brush_key] = brush
                    else:
                        resources.Add(brush_key, brush)
                    brushes_applied += 1
            except Exception as ex:
                pass
        
        return True
        
    except Exception as e:
        return False


def ensure_styles_loaded(force_theme=None):
    """
    Ensure CommonStyles are loaded into Application.Resources with theme support.
    
    This function loads the CommonStyles.xaml and applies the appropriate
    theme colors based on Revit's current UI theme setting.
    
    Args:
        force_theme: Optional. Force 'dark' or 'light' theme instead of auto-detecting.
    
    Usage:
        # In your script, call before creating WPFWindow:
        from lib.styles import ensure_styles_loaded
        ensure_styles_loaded()
        
        # Or force a specific theme:
        ensure_styles_loaded(force_theme='dark')
    """
    try:
        from System.Windows import Application, ResourceDictionary
        from System.Windows.Markup import XamlReader
        from System.IO import File
        
        styles_path = get_common_styles_path()
        
        if not op.exists(styles_path):
            return False
        
        # Check if styles are already loaded
        # NOTE: We always reload styles to ensure Background setters and other changes are applied
        # The early return was preventing new style definitions from being loaded
        styles_already_loaded = False
        if Application.Current is not None and Application.Current.Resources is not None:
            try:
                test_resource = Application.Current.Resources['BusyOverlayStyle']
                if test_resource is not None:
                    styles_already_loaded = True
            except:
                pass
        
        # Always reload styles to ensure latest changes (Background setters, etc.) are applied
        # Remove old merged dictionaries first
        if styles_already_loaded and Application.Current is not None and Application.Current.Resources is not None:
            try:
                # Remove existing CommonStyles dictionary if present
                merged_dicts_to_remove = []
                for i in range(Application.Current.Resources.MergedDictionaries.Count):
                    merged_dict = Application.Current.Resources.MergedDictionaries[i]
                    # Check if this dictionary contains our styles
                    if 'BusyOverlayStyle' in merged_dict.Keys:
                        merged_dicts_to_remove.append(i)
                # Remove in reverse order to maintain indices
                for i in reversed(merged_dicts_to_remove):
                    Application.Current.Resources.MergedDictionaries.RemoveAt(i)
            except:
                pass
        
        # Load the XAML
        xaml_content = File.ReadAllText(styles_path)
        styles_dict = XamlReader.Parse(xaml_content)
        
        # Apply theme colors to the loaded dictionary
        apply_theme_to_resources(styles_dict, force_theme)
        
        # Ensure Application.Current exists
        if Application.Current is None:
            return False
        
        # Merge into Application.Resources
        if Application.Current.Resources is None:
            Application.Current.Resources = ResourceDictionary()
        
        # Try to merge
        try:
            Application.Current.Resources.MergedDictionaries.Add(styles_dict)
            
            # CRITICAL: Apply theme to the merged dictionary AND Application.Resources
            # This ensures brushes are available in both places for DynamicResource resolution
            apply_theme_to_resources(styles_dict, force_theme)
            apply_theme_to_resources(Application.Current.Resources, force_theme)
            
            # CRITICAL: Also ensure ColoredButtonTextBrush is directly in Application.Resources
            # This ensures DynamicResource can find it even if MergedDictionaries lookup fails
            try:
                if 'ColoredButtonTextBrush' in styles_dict.Keys:
                    brush = styles_dict['ColoredButtonTextBrush']
                    Application.Current.Resources['ColoredButtonTextBrush'] = brush
            except: pass
        except Exception as e:
            # Fallback: copy resources manually
            for key in styles_dict.Keys:
                try:
                    Application.Current.Resources[key] = styles_dict[key]
                except:
                    pass
            # Apply theme to fallback resources
            apply_theme_to_resources(Application.Current.Resources, force_theme)
        
        return True
        
    except Exception as e:
        return False


def load_styles_to_window(window, force_theme=None):
    """
    Load styles directly into a window's Resources with theme support.
    
    This is useful when Application.Resources is not available.
    
    Args:
        window: A WPF Window instance
        force_theme: Optional. Force 'dark' or 'light' theme.
    
    Returns:
        bool: True if styles were loaded successfully
    """
    try:
        from System.Windows import ResourceDictionary
        from System.Windows.Markup import XamlReader
        from System.IO import File
        
        styles_path = get_common_styles_path()
        
        if not op.exists(styles_path):
            return False
        
        # Load the XAML
        xaml_content = File.ReadAllText(styles_path)
        styles_dict = XamlReader.Parse(xaml_content)
        
        # Apply theme colors
        apply_theme_to_resources(styles_dict, force_theme)
        
        # Ensure window has Resources
        if window.Resources is None:
            window.Resources = ResourceDictionary()
        
        # Copy resources to window
        for key in styles_dict.Keys:
            try:
                window.Resources[key] = styles_dict[key]
            except:
                pass
        
        return True
        
    except Exception as e:
        return False
