# -*- coding: utf-8 -*-
"""Clash Explorer - Visualize and interact with clash detection results.

This tool connects to a clash detection API to display clash tests, groups,
and individual clashes with advanced filtering, grouping, and highlighting
capabilities in Revit.

Compatible with ifcclash JSON format.
"""
__title__ = "Clash\nExplorer"
__author__ = "Byggstyrning AB"
__doc__ = "Visualize and interact with clash detection results from ifcclash"

import os
import sys
import clr
from collections import namedtuple

# Add the extension directory to the path
import os.path as op
script_path = __file__
panel_dir = op.dirname(script_path)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(op.dirname(tab_dir))
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Add reference to WPF
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *

from System.Collections.ObjectModel import ObservableCollection
from System.Windows import MessageBox, Visibility

from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Import custom modules
from clashes import clash_api
from clashes import clash_utils

# Initialize logger
logger = script.get_logger()

# Define data classes using namedtuple for IronPython compatibility
ClashSet = namedtuple('ClashSet', ['Name', 'Mode', 'TotalClashes', 'Data'])
ClashGroup = namedtuple('ClashGroup', ['Id', 'GroupName', 'ClashCount', 'CategoryA', 'CategoryB', 'LevelA', 'LevelB', 'Data'])
ClashDetail = namedtuple('ClashDetail', ['ClashKey', 'ElementA', 'ElementB', 'CategoryA', 'CategoryB', 'Distance', 'Data'])

class ClashExplorerUI(forms.WPFWindow):
    """Clash Explorer UI implementation."""
    
    def __init__(self):
        """Initialize the Clash Explorer UI."""
        # Initialize WPF window
        forms.WPFWindow.__init__(self, 'ClashExplorer.xaml')
        
        # Initialize API client
        self.clash_api_client = None
        
        # Initialize data collections
        self.clash_sets = ObservableCollection[object]()
        self.clash_groups = ObservableCollection[object]()
        self.clash_details = ObservableCollection[object]()
        self.all_clash_groups = []  # Store unfiltered groups
        
        # Current selections
        self.current_clash_set = None
        self.current_group = None
        
        # GUID lookup dictionary
        self.guid_dict = None
        
        # Set up event handlers
        self.saveSettingsButton.Click += self.save_settings_button_click
        self.loadClashTestsButton.Click += self.load_clash_sets_button_click
        self.selectClashTestButton.Click += self.select_clash_set_button_click
        self.applyFiltersButton.Click += self.apply_filters_button_click
        self.highlightAllButton.Click += self.highlight_all_button_click
        self.exportButton.Click += self.export_button_click
        
        # Data grid selection changed events
        self.clashTestsDataGrid.SelectionChanged += self.clash_set_selection_changed
        self.clashGroupsDataGrid.SelectionChanged += self.clash_group_selection_changed
        
        # Initialize UI
        self.clashTestsDataGrid.ItemsSource = self.clash_sets
        self.clashGroupsDataGrid.ItemsSource = self.clash_groups
        self.clashDetailsDataGrid.ItemsSource = self.clash_details
        
        # Load saved settings
        self.load_saved_settings()
        
        # Build GUID lookup dictionary in background
        self.update_status("Building GUID lookup dictionary...")
        try:
            self.guid_dict = clash_utils.build_guid_lookup_dict(revit.doc)
            self.update_status("Ready - {} elements in GUID index".format(len(self.guid_dict)))
        except Exception as e:
            logger.error("Error building GUID lookup: {}".format(str(e)))
            self.update_status("Ready - GUID lookup failed")
    
    def load_saved_settings(self):
        """Load saved API settings from extensible storage."""
        try:
            api_url, api_key = clash_api.load_clash_settings(revit.doc)
            if api_url:
                self.apiUrlTextBox.Text = api_url
            if api_key:
                self.apiKeyPasswordBox.Password = api_key
            
            if api_url and api_key:
                self.update_status("Loaded saved settings")
                # Automatically initialize API client
                self.clash_api_client = clash_api.ClashAPIClient(api_url, api_key)
        except Exception as e:
            logger.error("Error loading saved settings: {}".format(str(e)))
    
    def save_settings_button_click(self, sender, args):
        """Handle save settings button click."""
        api_url = self.apiUrlTextBox.Text
        api_key = self.apiKeyPasswordBox.Password
        
        if not api_url or not api_key:
            self.update_status("Please enter both API URL and API key")
            return
        
        # Save to extensible storage
        if clash_api.save_clash_settings(revit.doc, api_url, api_key):
            self.update_status("Settings saved successfully")
            # Initialize API client
            self.clash_api_client = clash_api.ClashAPIClient(api_url, api_key)
            self.clashTestsTab.IsEnabled = True
        else:
            self.update_status("Failed to save settings")
    
    def load_clash_sets_button_click(self, sender, args):
        """Handle load clash sets button click."""
        if not self.clash_api_client:
            self.update_status("Please save settings first")
            return
        
        self.update_status("Loading clash sets...")
        self.progressBar.Visibility = Visibility.Visible
        
        try:
            # Get clash sets from API (ifcclash format)
            clash_sets = self.clash_api_client.get_clash_sets()
            
            if clash_sets:
                self.clash_sets.Clear()
                for clash_set in clash_sets:
                    # Parse ifcclash ClashSet format
                    clashes = clash_set.get('clashes', {})
                    set_obj = ClashSet(
                        Name=clash_set.get('name', 'Unknown'),
                        Mode=clash_set.get('mode', 'collision'),
                        TotalClashes=len(clashes),
                        Data=clash_set
                    )
                    self.clash_sets.Add(set_obj)
                
                self.update_status("Loaded {} clash sets".format(len(clash_sets)))
                self.clashTestsTab.IsEnabled = True
                self.tabControl.SelectedItem = self.clashTestsTab
            else:
                error_msg = self.clash_api_client.last_error or "No clash sets found"
                self.update_status(error_msg)
        except Exception as e:
            logger.error("Error loading clash sets: {}".format(str(e)))
            self.update_status("Error loading clash sets: {}".format(str(e)))
        finally:
            self.progressBar.Visibility = Visibility.Collapsed
    
    def clash_set_selection_changed(self, sender, args):
        """Handle clash set selection change."""
        selected = self.clashTestsDataGrid.SelectedItem
        if selected:
            self.selectClashTestButton.IsEnabled = True
        else:
            self.selectClashTestButton.IsEnabled = False
    
    def select_clash_set_button_click(self, sender, args):
        """Handle select clash set button click."""
        selected_set = self.clashTestsDataGrid.SelectedItem
        
        if not selected_set:
            self.update_status("Please select a clash set")
            return
        
        self.current_clash_set = selected_set
        self.selectedTestNameRun.Text = selected_set.Name
        
        self.update_status("Processing clash set: {}".format(selected_set.Name))
        self.progressBar.Visibility = Visibility.Visible
        
        try:
            # Get clashes from the clash set (ifcclash format: dict of clash_id -> ClashResult)
            clash_set_data = selected_set.Data
            clashes_dict = clash_set_data.get('clashes', {})
            
            if clashes_dict:
                # Enrich clashes with Revit data
                enriched_clashes = clash_utils.enrich_clash_data_with_revit_info(revit.doc, clashes_dict)
                
                # Create groups from enriched clashes
                groups = self.create_groups_from_clashes(enriched_clashes, selected_set.Name)
                
                self.all_clash_groups = groups
                self.populate_clash_groups(groups)
                self.populate_filter_dropdowns(groups)
                
                self.update_status("Created {} clash groups from {} clashes".format(len(groups), len(enriched_clashes)))
                self.clashGroupsTab.IsEnabled = True
                self.tabControl.SelectedItem = self.clashGroupsTab
            else:
                self.update_status("No clashes found in this clash set")
        except Exception as e:
            logger.error("Error loading clash groups: {}".format(str(e)))
            self.update_status("Error loading clash groups: {}".format(str(e)))
        finally:
            self.progressBar.Visibility = Visibility.Collapsed
    
    def create_groups_from_clashes(self, clashes_dict, set_name):
        """Create logical groups from clash dictionary.
        
        Groups clashes by category pairs for better organization.
        """
        from collections import defaultdict
        
        # Group by category pair
        category_groups = defaultdict(list)
        
        for clash_key, clash in clashes_dict.items():
            cat_a = clash.get('revit_category_a', clash.get('a_ifc_class', 'Unknown'))
            cat_b = clash.get('revit_category_b', clash.get('b_ifc_class', 'Unknown'))
            
            # Create a consistent group key
            group_key = "{} vs {}".format(cat_a, cat_b)
            category_groups[group_key].append((clash_key, clash))
        
        # Create group objects
        groups = []
        for group_name, clashes_list in sorted(category_groups.items(), key=lambda x: len(x[1]), reverse=True):
            # Get representative clash for metadata
            first_clash = clashes_list[0][1]
            
            group_data = {
                'name': group_name,
                'clash_count': len(clashes_list),
                'category_a': first_clash.get('revit_category_a', first_clash.get('a_ifc_class', 'Unknown')),
                'category_b': first_clash.get('revit_category_b', first_clash.get('b_ifc_class', 'Unknown')),
                'level_a': first_clash.get('revit_level_a', 'Unknown'),
                'level_b': first_clash.get('revit_level_b', 'Unknown'),
                'clashes': dict(clashes_list)  # Convert list of tuples to dict
            }
            groups.append(group_data)
        
        return groups
    
    def populate_clash_groups(self, groups):
        """Populate the clash groups data grid."""
        self.clash_groups.Clear()
        
        for group in groups:
            group_obj = ClashGroup(
                Id=group.get('name', ''),
                GroupName=group.get('name', 'Unknown'),
                ClashCount=group.get('clash_count', 0),
                CategoryA=group.get('category_a', 'Unknown'),
                CategoryB=group.get('category_b', 'Unknown'),
                LevelA=group.get('level_a', 'Unknown'),
                LevelB=group.get('level_b', 'Unknown'),
                Data=group
            )
            self.clash_groups.Add(group_obj)
    
    def populate_filter_dropdowns(self, groups):
        """Populate filter dropdowns with unique values."""
        categories = set()
        levels = set()
        
        for group in groups:
            categories.add(group.get('category_a', 'Unknown'))
            categories.add(group.get('category_b', 'Unknown'))
            levels.add(group.get('level_a', 'Unknown'))
            levels.add(group.get('level_b', 'Unknown'))
        
        # Populate category filter
        self.categoryFilterComboBox.Items.Clear()
        self.categoryFilterComboBox.Items.Add("All Categories")
        for cat in sorted(categories):
            self.categoryFilterComboBox.Items.Add(cat)
        self.categoryFilterComboBox.SelectedIndex = 0
        
        # Populate level filter
        self.levelFilterComboBox.Items.Clear()
        self.levelFilterComboBox.Items.Add("All Levels")
        for level in sorted(levels):
            self.levelFilterComboBox.Items.Add(level)
        self.levelFilterComboBox.SelectedIndex = 0
    
    def apply_filters_button_click(self, sender, args):
        """Apply filters to clash groups."""
        category_filter = self.categoryFilterComboBox.SelectedItem
        level_filter = self.levelFilterComboBox.SelectedItem
        
        filtered_groups = self.all_clash_groups
        
        # Apply category filter
        if category_filter and category_filter != "All Categories":
            filtered_groups = [g for g in filtered_groups 
                             if g.get('category_a') == category_filter or g.get('category_b') == category_filter]
        
        # Apply level filter
        if level_filter and level_filter != "All Levels":
            filtered_groups = [g for g in filtered_groups 
                             if g.get('level_a') == level_filter or g.get('level_b') == level_filter]
        
        self.populate_clash_groups(filtered_groups)
        self.update_status("Filtered to {} groups".format(len(filtered_groups)))
    
    def clash_group_selection_changed(self, sender, args):
        """Handle clash group selection change."""
        selected = self.clashGroupsDataGrid.SelectedItem
        if selected:
            self.current_group = selected
            self.load_clash_details(selected)
    
    def load_clash_details(self, group):
        """Load clash details for a selected group."""
        try:
            self.selectedGroupNameRun.Text = group.GroupName
            self.clash_details.Clear()
            
            # Get clashes from the group data (ifcclash format)
            clashes_dict = group.Data.get('clashes', {})
            
            for clash_key, clash in clashes_dict.items():
                # Use ifcclash field names
                guid_a = clash.get('a_global_id', '')
                guid_b = clash.get('b_global_id', '')
                
                detail = ClashDetail(
                    ClashKey=clash_key,
                    ElementA="{}...".format(guid_a[:8]) if guid_a else "Unknown",
                    ElementB="{}...".format(guid_b[:8]) if guid_b else "Unknown",
                    CategoryA=clash.get('revit_category_a', clash.get('a_ifc_class', 'Unknown')),
                    CategoryB=clash.get('revit_category_b', clash.get('b_ifc_class', 'Unknown')),
                    Distance="{:.2f}".format(clash.get('distance', 0.0)),
                    Data=clash
                )
                self.clash_details.Add(detail)
            
            self.clashDetailsTab.IsEnabled = True
            self.update_status("Loaded {} clashes in group".format(len(clashes_dict)))
            
        except Exception as e:
            logger.error("Error loading clash details: {}".format(str(e)))
            self.update_status("Error loading clash details: {}".format(str(e)))
    
    def highlight_button_click(self, sender, args):
        """Handle highlight button click for a clash group."""
        button = sender
        group = button.Tag
        
        if not group:
            return
        
        try:
            self.update_status("Highlighting clash group: {}".format(group.GroupName))
            
            # Get all GUIDs from the group (ifcclash format)
            clashes_dict = group.Data.get('clashes', {})
            guids = []
            for clash_key, clash in clashes_dict.items():
                guid_a = clash.get('a_global_id')
                guid_b = clash.get('b_global_id')
                if guid_a:
                    guids.append(guid_a)
                if guid_b:
                    guids.append(guid_b)
            
            # Find elements
            element_dict = {}
            for guid in guids:
                if self.guid_dict and guid in self.guid_dict:
                    element_dict[guid] = self.guid_dict[guid]
            
            if not element_dict:
                self.update_status("No elements found in current model for this clash group")
                return
            
            # Create view name
            view_name = "Clash - {}".format(group.GroupName)
            
            # Create 3D view
            view = clash_utils.create_clash_view(revit.doc, view_name, element_dict)
            
            if view:
                # Highlight elements
                clash_utils.highlight_clash_elements(revit.doc, view, element_dict)
                
                # Set other models as underlay
                clash_utils.set_other_models_as_underlay(revit.doc, view, element_dict)
                
                # Set as active view
                with revit.Transaction("Set Active View"):
                    revit.uidoc.ActiveView = view
                
                self.update_status("Created and highlighted clash view: {}".format(view_name))
            else:
                self.update_status("Failed to create clash view")
                
        except Exception as e:
            logger.error("Error highlighting clash group: {}".format(str(e)))
            self.update_status("Error highlighting clash group: {}".format(str(e)))
    
    def highlight_clash_button_click(self, sender, args):
        """Handle highlight button click for an individual clash."""
        button = sender
        clash_detail = button.Tag
        
        if not clash_detail:
            return
        
        try:
            self.update_status("Highlighting clash: {}".format(clash_detail.ClashKey))
            
            # Get GUIDs from clash (ifcclash format)
            clash_data = clash_detail.Data
            guid_a = clash_data.get('a_global_id')
            guid_b = clash_data.get('b_global_id')
            
            guids = []
            if guid_a:
                guids.append(guid_a)
            if guid_b:
                guids.append(guid_b)
            
            # Find elements
            element_dict = {}
            for guid in guids:
                if self.guid_dict and guid in self.guid_dict:
                    element_dict[guid] = self.guid_dict[guid]
            
            if not element_dict:
                self.update_status("No elements found in current model for this clash")
                return
            
            # Create view name
            view_name = "Clash - {}".format(clash_detail.ClashKey[:16])
            
            # Create 3D view
            view = clash_utils.create_clash_view(revit.doc, view_name, element_dict)
            
            if view:
                # Highlight elements
                clash_utils.highlight_clash_elements(revit.doc, view, element_dict)
                
                # Set as active view
                with revit.Transaction("Set Active View"):
                    revit.uidoc.ActiveView = view
                
                self.update_status("Created and highlighted clash view: {}".format(view_name))
            else:
                self.update_status("Failed to create clash view")
                
        except Exception as e:
            logger.error("Error highlighting clash: {}".format(str(e)))
            self.update_status("Error highlighting clash: {}".format(str(e)))
    
    def highlight_all_button_click(self, sender, args):
        """Handle highlight all button click."""
        if not self.current_group:
            self.update_status("Please select a clash group first")
            return
        
        # Create a button object with the current group as Tag
        class FakeButton:
            def __init__(self, tag):
                self.Tag = tag
        
        fake_button = FakeButton(self.current_group)
        self.highlight_button_click(fake_button, None)
    
    def export_button_click(self, sender, args):
        """Handle export button click."""
        self.update_status("Export functionality not yet implemented")
    
    def update_status(self, message):
        """Update status text."""
        self.statusTextBlock.Text = message
        logger.debug(message)

# Main execution
if __name__ == '__main__':
    # Show the Clash Explorer UI
    ClashExplorerUI().ShowDialog()
