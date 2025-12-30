# Styles module for reusable WPF UI styles
"""
This module provides reusable WPF styles for PyRevit extensions.

Usage in XAML:
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/lib/styles/CommonStyles.xaml"/>
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
"""

import os
import os.path as op
from pyrevit import script

logger = script.get_logger()

def get_styles_path():
    """Get the absolute path to the styles directory."""
    current_dir = op.dirname(__file__)
    return current_dir

def get_common_styles_path():
    """Get the absolute path to CommonStyles.xaml (Light)."""
    return op.join(get_styles_path(), 'CommonStyles.xaml')

def get_dark_styles_path():
    """Get the absolute path to CommonStyles.Dark.xaml (Dark)."""
    return op.join(get_styles_path(), 'CommonStyles.Dark.xaml')

def get_active_theme_styles_path():
    """Get the path to the styles file matching the current Revit theme."""
    try:
        from Autodesk.Revit.UI import UIThemeManager, UITheme
        if UIThemeManager.CurrentTheme == UITheme.Dark:
            return get_dark_styles_path()
    except Exception:
        # Fallback for older Revit versions or if UIThemeManager is missing
        pass
    return get_common_styles_path()

def load_common_styles(window):
    """Load common styles into the window based on current Revit theme.
    
    Args:
        window (WPFWindow): The window to load styles into.
    """
    try:
        from System.Windows.Markup import XamlReader
        from System.IO import File
        from System.Windows import ResourceDictionary

        styles_path = get_active_theme_styles_path()
        logger.debug("Loading styles from: {}".format(styles_path))
        
        if op.exists(styles_path):
            # Read XAML content
            xaml_content = File.ReadAllText(styles_path)
            
            # Parse as ResourceDictionary
            styles_dict = XamlReader.Parse(xaml_content)
            
            # Merge into window resources
            if window.Resources is None:
                window.Resources = ResourceDictionary()
            
            # Merge styles into existing resources
            if hasattr(styles_dict, 'MergedDictionaries'):
                for merged_dict in styles_dict.MergedDictionaries:
                    window.Resources.MergedDictionaries.Add(merged_dict)
            
            # Copy individual resources
            for key in styles_dict.Keys:
                window.Resources[key] = styles_dict[key]
    except Exception as e:
        logger.debug("Could not load styles: {}".format(e))
