# -*- coding: utf-8 -*-
"""Clash Explorer - Visualize and interact with clash detection results.

This tool connects to a clash detection API to display clash tests, groups,
and individual clashes with advanced filtering, grouping, and highlighting
capabilities in Revit.
"""
__title__ = "Clash\nExplorer"
__author__ = "Byggstyrning AB"
__doc__ = "Visualize and interact with clash detection results from external APIs"

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
ClashTest = namedtuple('ClashTest', ['Id', 'Name', 'TotalClashes', 'Status', 'LastRun'])
ClashGroup = namedtuple('ClashGroup', ['Id', 'GroupName', 'ClashCount', 'CategoryA', 'CategoryB', 'LevelA', 'LevelB', 'Data'])
ClashDetail = namedtuple('ClashDetail', ['ClashId', 'ElementA', 'ElementB', 'CategoryA', 'CategoryB', 'Distance', 'Data'])

class ClashExplorerUI(forms.WPFWindow):
    """Clash Explorer UI implementation."""
    
    def __init__(self):
        """Initialize the Clash Explorer UI."""
        # Initialize WPF window
        forms.WPFWindow.__init__(self, 'ClashExplorer.xaml')
        
        # Initialize API client
        self.clash_api_client = None
        
        # Initialize data collections
        self.clash_tests = ObservableCollection[object]()
        self.clash_groups = ObservableCollection[object]()
        self.clash_details = ObservableCollection[object]()
        self.all_clash_groups = []  # Store unfiltered groups
        
        # Current selections
        self.current_test = None
        self.current_group = None
        
        # GUID lookup dictionary
        self.guid_dict = None
        
        # Set up event handlers
        self.saveSettingsButton.Click += self.save_settings_button_click
        self.loadClashTestsButton.Click += self.load_clash_tests_button_click
        self.selectClashTestButton.Click += self.select_clash_test_button_click
        self.applyFiltersButton.Click += self.apply_filters_button_click
        self.highlightAllButton.Click += self.highlight_all_button_click
        self.exportButton.Click += self.export_button_click
        
        # Data grid selection changed events
        self.clashTestsDataGrid.SelectionChanged += self.clash_test_selection_changed
        self.clashGroupsDataGrid.SelectionChanged += self.clash_group_selection_changed
        
        # Initialize UI
        self.clashTestsDataGrid.ItemsSource = self.clash_tests
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
    
    def load_clash_tests_button_click(self, sender, args):
        """Handle load clash tests button click."""
        if not self.clash_api_client:
            self.update_status("Please save settings first")
            return
        
        self.update_status("Loading clash tests...")
        self.progressBar.Visibility = Visibility.Visible
        
        try:
            # Get clash tests from API
            tests = self.clash_api_client.get_clash_tests()
            
            if tests:
                self.clash_tests.Clear()
                for test in tests:
                    test_obj = ClashTest(
                        Id=test.get('id', ''),
                        Name=test.get('name', 'Unknown'),
                        TotalClashes=test.get('total_clashes', 0),
                        Status=test.get('status', 'Unknown'),
                        LastRun=test.get('last_run', 'N/A')
                    )
                    self.clash_tests.Add(test_obj)
                
                self.update_status("Loaded {} clash tests".format(len(tests)))
                self.clashTestsTab.IsEnabled = True
                self.tabControl.SelectedItem = self.clashTestsTab
            else:
                error_msg = self.clash_api_client.last_error or "No clash tests found"
                self.update_status(error_msg)
        except Exception as e:
            logger.error("Error loading clash tests: {}".format(str(e)))
            self.update_status("Error loading clash tests: {}".format(str(e)))
        finally:
            self.progressBar.Visibility = Visibility.Collapsed
    
    def clash_test_selection_changed(self, sender, args):
        """Handle clash test selection change."""
        selected = self.clashTestsDataGrid.SelectedItem
        if selected:
            self.selectClashTestButton.IsEnabled = True
        else:
            self.selectClashTestButton.IsEnabled = False
    
    def select_clash_test_button_click(self, sender, args):
        """Handle select clash test button click."""
        selected_test = self.clashTestsDataGrid.SelectedItem
        
        if not selected_test:
            self.update_status("Please select a clash test")
            return
        
        self.current_test = selected_test
        self.selectedTestNameRun.Text = selected_test.Name
        
        self.update_status("Loading clash groups for: {}".format(selected_test.Name))
        self.progressBar.Visibility = Visibility.Visible
        
        try:
            # Get clash groups from API
            groups = self.clash_api_client.get_clash_groups(selected_test.Id)
            
            if groups:
                # Enrich with Revit data
                enriched_groups = self.enrich_groups_with_revit_data(groups)
                
                self.all_clash_groups = enriched_groups
                self.populate_clash_groups(enriched_groups)
                self.populate_filter_dropdowns(enriched_groups)
                
                self.update_status("Loaded {} clash groups".format(len(enriched_groups)))
                self.clashGroupsTab.IsEnabled = True
                self.tabControl.SelectedItem = self.clashGroupsTab
            else:
                error_msg = self.clash_api_client.last_error or "No clash groups found"
                self.update_status(error_msg)
        except Exception as e:
            logger.error("Error loading clash groups: {}".format(str(e)))
            self.update_status("Error loading clash groups: {}".format(str(e)))
        finally:
            self.progressBar.Visibility = Visibility.Collapsed
    
    def enrich_groups_with_revit_data(self, groups):
        """Enrich clash groups with Revit category and level information."""
        enriched = []
        
        for group in groups:
            try:
                # Get GUIDs from the group (assuming first clash is representative)
                clashes = group.get('clashes', [])
                if not clashes:
                    continue
                
                # Get categories and levels from first clash
                first_clash = clashes[0]
                guid_a = first_clash.get('guid_a') or first_clash.get('object_a')
                guid_b = first_clash.get('guid_b') or first_clash.get('object_b')
                
                # Look up in GUID dictionary
                category_a = "Unknown"
                category_b = "Unknown"
                level_a = "Unknown"
                level_b = "Unknown"
                
                if self.guid_dict and guid_a in self.guid_dict:
                    element_a, is_linked_a, link_a = self.guid_dict[guid_a]
                    if element_a.Category:
                        category_a = element_a.Category.Name
                    
                    # Get level
                    try:
                        level_param = element_a.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
                        if level_param and level_param.HasValue:
                            level_id = level_param.AsElementId()
                            level = revit.doc.GetElement(level_id)
                            if level:
                                level_a = level.Name
                    except:
                        pass
                
                if self.guid_dict and guid_b in self.guid_dict:
                    element_b, is_linked_b, link_b = self.guid_dict[guid_b]
                    if element_b.Category:
                        category_b = element_b.Category.Name
                    
                    # Get level
                    try:
                        level_param = element_b.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
                        if level_param and level_param.HasValue:
                            level_id = level_param.AsElementId()
                            level = revit.doc.GetElement(level_id)
                            if level:
                                level_b = level.Name
                    except:
                        pass
                
                # Create enriched group
                enriched_group = {
                    'id': group.get('id'),
                    'name': group.get('name', 'Unknown'),
                    'clash_count': group.get('clash_count', len(clashes)),
                    'category_a': category_a,
                    'category_b': category_b,
                    'level_a': level_a,
                    'level_b': level_b,
                    'clashes': clashes
                }
                enriched.append(enriched_group)
                
            except Exception as e:
                logger.error("Error enriching group: {}".format(str(e)))
                continue
        
        return enriched
    
    def populate_clash_groups(self, groups):
        """Populate the clash groups data grid."""
        self.clash_groups.Clear()
        
        for group in groups:
            group_obj = ClashGroup(
                Id=group.get('id', ''),
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
            
            # Get clashes from the group data
            clashes = group.Data.get('clashes', [])
            
            for clash in clashes:
                guid_a = clash.get('guid_a') or clash.get('object_a')
                guid_b = clash.get('guid_b') or clash.get('object_b')
                
                detail = ClashDetail(
                    ClashId=clash.get('id', ''),
                    ElementA=guid_a[:8] + "..." if guid_a else "Unknown",
                    ElementB=guid_b[:8] + "..." if guid_b else "Unknown",
                    CategoryA=clash.get('category_a', 'Unknown'),
                    CategoryB=clash.get('category_b', 'Unknown'),
                    Distance="{:.2f}".format(clash.get('distance', 0.0)),
                    Data=clash
                )
                self.clash_details.Add(detail)
            
            self.clashDetailsTab.IsEnabled = True
            self.update_status("Loaded {} clashes in group".format(len(clashes)))
            
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
            
            # Get all GUIDs from the group
            clashes = group.Data.get('clashes', [])
            guids = []
            for clash in clashes:
                guid_a = clash.get('guid_a') or clash.get('object_a')
                guid_b = clash.get('guid_b') or clash.get('object_b')
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
            self.update_status("Highlighting clash: {}".format(clash_detail.ClashId))
            
            # Get GUIDs from clash
            clash_data = clash_detail.Data
            guid_a = clash_data.get('guid_a') or clash_data.get('object_a')
            guid_b = clash_data.get('guid_b') or clash_data.get('object_b')
            
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
            view_name = "Clash - {}".format(clash_detail.ClashId)
            
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
        
        # Simulate clicking the highlight button for the current group
        self.highlight_button_click(sender, args)
    
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
