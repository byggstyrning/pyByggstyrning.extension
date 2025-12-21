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

def get_styles_path():
    """Get the absolute path to the styles directory."""
    current_dir = op.dirname(__file__)
    return current_dir

def get_common_styles_path():
    """Get the absolute path to CommonStyles.xaml."""
    return op.join(get_styles_path(), 'CommonStyles.xaml')
