# -*- coding: utf-8 -*-
"""Shared selector dialog base classes for 3D Zone creation."""

import os.path as op
from pyrevit import forms, script

logger = script.get_logger()


class SpatialElementItem(object):
    """Generic wrapper for spatial element items in selector dialogs."""
    
    def __init__(self, element, doc, has_zone=False, adapter=None, element_attr_name=None):
        """Initialize spatial element item.
        
        Args:
            element: Area or Room element
            doc: Revit document
            has_zone: Whether element already has a 3D Zone
            adapter: SpatialElementAdapter instance
            element_attr_name: Attribute name for element (e.g., "area" or "room")
        """
        self.element = element
        self.has_zone = has_zone
        self.adapter = adapter
        
        # Store element with the appropriate attribute name for compatibility
        if element_attr_name:
            setattr(self, element_attr_name, element)
        
        # Get element display info using adapter
        if adapter:
            self.element_number = adapter.get_number(element)
            self.element_name = adapter.get_name(element)
        else:
            self.element_number = "?"
            self.element_name = "Unnamed"
        
        # Get level name
        level_name = "?"
        level_id = adapter.get_level_id(element) if adapter else None
        if level_id:
            level_elem = doc.GetElement(level_id)
            if level_elem:
                level_name = level_elem.Name
        
        # Store level_name as separate attribute
        self.level_name = level_name
        
        # Get phase name
        phase_name = "?"
        if adapter and hasattr(adapter, 'get_phase_id'):
            phase_id = adapter.get_phase_id(element)
            if phase_id:
                phase_elem = doc.GetElement(phase_id)
                if phase_elem:
                    phase_name = phase_elem.Name
        else:
            phase_name = "?"
        
        # Store phase_name as separate attribute
        self.phase_name = phase_name
        
        # Store exists_display for DataGrid (checkmark or empty)
        self.exists_display = "✓" if has_zone else ""
        
        # Create display text with icon: (*) if zone exists, empty if not
        icon = "(*)" if has_zone else "   "
        self.display_text = "{} {} - {} ({})".format(icon, self.element_number, self.element_name, level_name)
    
    def __str__(self):
        """String representation for pyRevit forms."""
        return self.display_text
    
    def __repr__(self):
        """Representation for debugging."""
        return self.display_text


class SpatialSelectorWindow(forms.WPFWindow):
    """Base class for spatial element selector dialogs."""
    
    def __init__(self, element_items, pushbutton_dir, extension_dir):
        """Initialize the selector window.
        
        Args:
            element_items: List of SpatialElementItem objects
            pushbutton_dir: Pushbutton directory (for XAML files)
            extension_dir: Extension root directory (for styles)
        """
        # Load XAML file
        xaml_path = op.join(pushbutton_dir, self.get_xaml_filename())
        forms.WPFWindow.__init__(self, xaml_path)
        
        # Load styles AFTER window initialization (window-scoped, does not affect Revit UI)
        # Note: lib is added to sys.path by scripts that import this module
        from styles import load_styles_to_window
        load_styles_to_window(self)
        
        # Store all items and filtered items
        self.all_items = element_items
        self.filtered_items = list(element_items)
        
        # Store selected elements (will be populated when Create is clicked)
        self.selected_elements = None
        
        # Bind collection to ListView (if it exists - subclasses may use DataGrid instead)
        listview_name = self.get_listview_name()
        if hasattr(self, listview_name):
            listview = getattr(self, listview_name)
            if listview:
                listview.ItemsSource = self.filtered_items
        
        # Set up event handlers
        if hasattr(self, 'createButton'):
            self.createButton.Click += self.create_button_click
        if hasattr(self, 'cancelButton'):
            self.cancelButton.Click += self.cancel_button_click
    
    def get_xaml_filename(self):
        """Return XAML filename. Must be implemented by subclass.
        
        Returns:
            str: XAML filename (e.g., "AreaSelector.xaml")
        """
        raise NotImplementedError
    
    def get_listview_name(self):
        """Return ListView control name. Must be implemented by subclass.
        
        Returns:
            str: ListView name (e.g., "areasListView")
        """
        raise NotImplementedError
    
    def get_element_attribute(self):
        """Return element attribute name. Must be implemented by subclass.
        
        Returns:
            str: Attribute name (e.g., "area" or "room")
        """
        raise NotImplementedError
    
    def load_styles(self, extension_dir):
        """Load the common styles ResourceDictionary."""
        try:
            styles_path = op.join(extension_dir, 'lib', 'styles', 'CommonStyles.xaml')
            
            if op.exists(styles_path):
                from System.Windows.Markup import XamlReader
                from System.IO import File
                
                # Read XAML content
                xaml_content = File.ReadAllText(styles_path)
                
                # Parse as ResourceDictionary
                styles_dict = XamlReader.Parse(xaml_content)
                
                # Merge into window resources
                if self.Resources is None:
                    from System.Windows import ResourceDictionary
                    self.Resources = ResourceDictionary()
                
                # Merge styles into existing resources
                if hasattr(styles_dict, 'MergedDictionaries'):
                    for merged_dict in styles_dict.MergedDictionaries:
                        self.Resources.MergedDictionaries.Add(merged_dict)
                
                # Copy individual resources
                for key in styles_dict.Keys:
                    self.Resources[key] = styles_dict[key]
        except Exception as e:
            logger.debug("Could not load styles: {}".format(e))
    
    def searchTextBox_TextChanged(self, sender, args):
        """Handle search text box text changed event."""
        try:
            search_text = sender.Text.lower() if sender.Text else ""
            
            # Filter items based on search text
            if search_text:
                self.filtered_items = [
                    item for item in self.all_items
                    if search_text in item.display_text.lower()
                ]
            else:
                self.filtered_items = list(self.all_items)
            
            # Update ListView
            listview_name = self.get_listview_name()
            listview = getattr(self, listview_name)
            if listview:
                listview.ItemsSource = self.filtered_items
        except Exception as e:
            logger.debug("Error filtering elements: {}".format(e))
    
    def create_button_click(self, sender, args):
        """Handle Create button click - collect selected elements and close."""
        # Collect all selected items
        selected_items = []
        listview_name = self.get_listview_name()
        listview = getattr(self, listview_name)
        if listview:
            for item in listview.SelectedItems:
                selected_items.append(item)
        
        # Extract elements from selected items
        element_attr = self.get_element_attribute()
        if selected_items:
            # Try to get element via attribute, fallback to .element
            self.selected_elements = []
            for item in selected_items:
                if hasattr(item, element_attr):
                    self.selected_elements.append(getattr(item, element_attr))
                elif hasattr(item, 'element'):
                    self.selected_elements.append(item.element)
        else:
            self.selected_elements = []
        
        # Close window
        self.Close()
    
    def cancel_button_click(self, sender, args):
        """Handle Cancel button click - close without selection."""
        self.selected_elements = None
        self.Close()

