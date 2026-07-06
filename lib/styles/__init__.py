# -*- coding: utf-8 -*-
# Styles module for reusable WPF UI styles
"""
This module provides reusable WPF styles for PyRevit extensions.
Supports automatic dark mode detection based on Revit's UI theme.

Usage in Python:
    from lib.styles import load_styles_to_window, get_revit_theme, is_dark_theme
    
    # Load styles into window (window-scoped, does not affect Revit UI)
    class MyWindow(WPFWindow):
        def __init__(self):
            WPFWindow.__init__(self, xaml_file)
            load_styles_to_window(self)  # Load styles AFTER window creation
"""

import os
import os.path as op

# Color palettes for light and dark themes.
# IMPORTANT: these values are the runtime source of truth - they override the
# Color/Brush resources parsed from CommonStyles.xaml at load time, so any
# palette change must be mirrored in both places.
# Color roles: Accent (brand yellow) = primary actions/selected tabs;
# Selection (blue) = selection, checked state, focus; Success/Error/Warning =
# semantic status; neutrals for surfaces/borders/text.
LIGHT_THEME_COLORS = {
    'PlaceholderForegroundColor': '#6B7280',
    'BusyOverlayColor': '#A6000000',
    'AccentColor': '#FFBB00',
    'AccentHoverColor': '#EBA700',
    'AccentPressedColor': '#D69800',
    'SelectionColor': '#0078D4',
    'SelectionHoverColor': '#106EBE',
    'SelectionIndicatorColor': '#0078D4',
    'ProgressFillColor': '#0078D4',
    'ErrorColor': '#C62828',
    'ErrorHoverColor': '#B02525',
    'ErrorPressedColor': '#8E1C1C',
    'SuccessColor': '#2E7D32',
    'SuccessHoverColor': '#27692B',
    'SuccessPressedColor': '#1B5E20',
    'WarningColor': '#B45309',
    'BorderColor': '#D5D9DE',
    'BackgroundLightColor': '#F4F5F7',
    'BackgroundLighterColor': '#EAECEF',
    'TextColor': '#1F2328',
    'TextSecondaryColor': '#5B6470',
    'TextLightColor': '#8B939E',
    'DisabledColor': '#E4E7EB',
    'DisabledTextColor': '#9AA1AA',
    # Additional colors used in styles
    'WindowBackgroundColor': '#FAFAFA',
    'ControlBackgroundColor': '#FFFFFF',
    'PopupBackgroundColor': '#FFFFFF',
    'DataGridRowBackgroundColor': '#FFFFFF',
    'ArrowColor': '#6B7280',  # Neutral grey dropdown/sort glyphs
    'ColoredButtonTextColor': '#1A1A1A',  # Near-black text on the yellow accent
    'SemanticButtonTextColor': '#FFFFFF',  # White text on deep green/red buttons
}

DARK_THEME_COLORS = {
    'PlaceholderForegroundColor': '#9CA3AF',
    'BusyOverlayColor': '#66000000',
    'AccentColor': '#FFBB00',  # Keep accent color consistent
    'AccentHoverColor': '#EBA700',
    'AccentPressedColor': '#D69800',
    'SelectionColor': '#0078D4',
    'SelectionHoverColor': '#106EBE',
    'SelectionIndicatorColor': '#4CA0E0',  # Lighter blue so focus rings hit 3:1 on dark surfaces
    'ProgressFillColor': '#FFBB00',  # Yellow reads well against dark tracks
    'ErrorColor': '#E57373',  # Lighter for readability on dark surfaces
    'ErrorHoverColor': '#D75F5F',
    'ErrorPressedColor': '#D26666',
    'SuccessColor': '#81C784',
    'SuccessHoverColor': '#66BB6A',
    'SuccessPressedColor': '#4CAF50',
    'WarningColor': '#FFB74D',
    'BorderColor': '#4D5157',
    'BackgroundLightColor': '#37393D',
    'BackgroundLighterColor': '#43464B',
    'TextColor': '#E8EAED',
    'TextSecondaryColor': '#B4BAC2',
    'TextLightColor': '#8B939E',
    'DisabledColor': '#43464B',
    'DisabledTextColor': '#7D848D',
    # Additional colors used in styles
    'WindowBackgroundColor': '#2B2C2E',
    'ControlBackgroundColor': '#353639',
    'PopupBackgroundColor': '#38393C',
    'DataGridRowBackgroundColor': '#303134',
    'ArrowColor': '#9CA3AF',  # Neutral grey dropdown/sort glyphs
    'ColoredButtonTextColor': '#1A1A1A',  # Near-black text on the yellow accent
    'SemanticButtonTextColor': '#1A1A1A',  # Dark text on the pastel green/red buttons
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
            'SelectionBrush': 'SelectionColor',
            'SelectionHoverBrush': 'SelectionHoverColor',
            'SelectionIndicatorBrush': 'SelectionIndicatorColor',
            'ProgressFillBrush': 'ProgressFillColor',
            'ErrorBrush': 'ErrorColor',
            'ErrorHoverBrush': 'ErrorHoverColor',
            'ErrorPressedBrush': 'ErrorPressedColor',
            'SuccessBrush': 'SuccessColor',
            'SuccessHoverBrush': 'SuccessHoverColor',
            'SuccessPressedBrush': 'SuccessPressedColor',
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
            'SemanticButtonTextBrush': 'SemanticButtonTextColor',
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


def load_styles_to_window(window, force_theme=None):
    """
    Load styles directly into a window's Resources with theme support.
    
    This is the PRIMARY method for loading styles. It loads styles into the window's
    Resources collection, ensuring complete isolation from Revit's UI. Styles loaded
    this way will NOT affect Revit's own UI elements.
    
    IMPORTANT: This function must be called AFTER WPFWindow.__init__() because
    the window must exist before its Resources can be accessed.
    
    Args:
        window: A WPF Window instance (must be initialized)
        force_theme: Optional. Force 'dark' or 'light' theme instead of auto-detecting.
    
    Returns:
        bool: True if styles were loaded successfully
    
    Usage:
        class MyWindow(WPFWindow):
            def __init__(self):
                WPFWindow.__init__(self, xaml_file)
                load_styles_to_window(self)  # Load styles AFTER window creation
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
