# -*- coding: utf-8 -*-
"""MMI Settings.

This tool allows users to set the MMI settings for the current project.

"""

__title__ = "Settings"
__author__ = "Byggstyrning AB"
__doc__ = "MMI Settings"
__highlight__ = 'updated'

# Import standard libraries
import sys
import os.path as op
import json
import time

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *

# Import .NET System
import System
from System import Object
from System.Collections.ObjectModel import ObservableCollection

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Add the extension directory to the path
script_path = __file__
pushbutton_dir = op.dirname(script_path)
stack_dir = op.dirname(pushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

# Initialize logger
logger = script.get_logger()

# Import from MMI library modules
from mmi.config import CONFIG_KEYS
from mmi.core import get_mmi_parameter_name, save_mmi_parameter, save_monitor_config, load_monitor_config

# Import revit utils
from revit.revit_utils import get_available_parameters

# Import styles for theme support
from styles import load_styles_to_window


class MMISettingsWindow(forms.WPFWindow):
    """WPF window for MMI Settings configuration."""
    
    def __init__(self):
        """Initialize the MMI Settings window."""
        try:
            # Load XAML file
            xaml_file = op.join(op.dirname(__file__), 'MMISettingsWindow.xaml')
            forms.WPFWindow.__init__(self, xaml_file)
            
            # Load styles AFTER window initialization (window-scoped, does not affect Revit UI)
            load_styles_to_window(self)
            
            # Set up ComboBox
            self.parameterComboBox.DisplayMemberPath = "display_name"
            
            # Flag to prevent saving during initialization
            self._is_initializing = True
            
            # Load parameters into ComboBox
            self.load_parameters()
            
            # Load current configuration
            self.load_config()
            
            # Mark initialization as complete
            self._is_initializing = False
            
            # Flag to prevent saving during search/filtering
            self._is_searching = False
            
            # Wire up DropDownOpened and DropDownClosed events for search functionality
            self.parameterComboBox.DropDownOpened += self.on_parameter_dropdown_opened
            self.parameterComboBox.DropDownClosed += self.on_parameter_dropdown_closed
            
        except Exception as e:
            logger.error("Error initializing MMI Settings window: {}".format(str(e)))
            forms.alert("Failed to initialize MMI Settings window. See log for details.", title="Error")
            raise
    
    def load_parameters(self):
        """Load available parameters into the ComboBox."""
        try:
            # Get all available parameters
            available_parameters = get_available_parameters()
            
            if not available_parameters:
                logger.warning("No parameters found in the project")
                return
            
            # Create simple objects with display_name for the ComboBox
            # The SearchableComboBoxStyle expects items with display_name property
            class ParameterItem(Object):
                def __init__(self, name):
                    self.display_name = name
                    self.name = name
                
                def __str__(self):
                    return self.display_name
                
                def ToString(self):
                    """Explicit ToString implementation for WPF binding."""
                    return self.display_name
            
            items = ObservableCollection[Object]()
            for param_name in available_parameters:
                items.Add(ParameterItem(param_name))
            
            # Set ItemsSource
            self.parameterComboBox.ItemsSource = items
            
            # Store original items for search filtering
            self.all_parameters = list(items)
            
        except Exception as e:
            logger.error("Error loading parameters: {}".format(str(e)))
    
    def load_config(self):
        """Load current MMI configuration and parameter name."""
        try:
            # Load current MMI parameter name
            current_parameter = get_mmi_parameter_name(revit.doc)
            
            # Set selected parameter in ComboBox
            if self.parameterComboBox.ItemsSource:
                for item in self.parameterComboBox.ItemsSource:
                    if item.display_name == current_parameter:
                        self.parameterComboBox.SelectedItem = item
                        break
            
            # Load current monitor configuration (using schema keys, not display names)
            current_config = load_monitor_config(revit.doc, use_display_names=False)
            
            # Map schema keys to checkboxes
            self.validateMmiCheckBox.IsChecked = current_config.get("validate_mmi", False)
            self.pinElementsCheckBox.IsChecked = current_config.get("pin_elements", False)
            self.warnOnMoveCheckBox.IsChecked = current_config.get("warn_on_move", False)
            self.checkAfterSyncCheckBox.IsChecked = current_config.get("check_mmi_after_sync", False)
            
        except Exception as e:
            logger.error("Error loading MMI configuration: {}".format(str(e)))
    
    def ParameterComboBox_SelectionChanged(self, sender, e):
        """Handle parameter selection change."""
        # Do NOT save automatically - only save when Save button is clicked
        # This prevents saving during search/filtering operations
        # The _is_searching flag is set when dropdown opens and cleared when it closes
        # but we don't save on selection change regardless to ensure user must click Save
        pass
    
    def SaveButton_Click(self, sender, e):
        """Handle click on Save button."""
        try:
            # Save the selected MMI parameter first
            selected_item = self.parameterComboBox.SelectedItem
            if selected_item is not None:
                parameter_name = selected_item.display_name
                if save_mmi_parameter(revit.doc, parameter_name):
                    logger.debug("MMI parameter set to: {}".format(parameter_name))
                else:
                    error_msg = "Failed to save MMI parameter setting. See log for details."
                    forms.alert(error_msg, title="MMI Parameter Error")
                    return  # Don't continue if parameter save failed
            
            # Build config dictionary using schema keys
            config = {
                "validate_mmi": self.validateMmiCheckBox.IsChecked or False,
                "pin_elements": self.pinElementsCheckBox.IsChecked or False,
                "warn_on_move": self.warnOnMoveCheckBox.IsChecked or False,
                "check_mmi_after_sync": self.checkAfterSyncCheckBox.IsChecked or False
            }
            
            # Save the configuration
            if save_monitor_config(revit.doc, config):
                forms.show_balloon(
                    header="MMI Settings Saved",
                    text="MMI monitor configuration has been saved.",
                    tooltip="MMI monitor configuration has been saved.",
                    is_new=True
                )
                # Close window
                self.Close()
            else:
                error_msg = "Failed to save MMI Monitor configuration. See log for details."
                forms.warning(error_msg, title="MMI Monitor Config Error")
                
        except Exception as e:
            logger.error("Error saving MMI configuration: {}".format(str(e)))
            forms.alert("Error saving configuration. See log for details.", title="Error")
    
    def CancelButton_Click(self, sender, e):
        """Handle click on Cancel button."""
        self.Close()
    
    def on_parameter_dropdown_opened(self, sender, args):
        """Handle parameter dropdown opened event - initialize search filter."""
        try:
            # Set flag to indicate we're searching (prevents auto-save)
            self._is_searching = True
            
            # Store current selection before filtering
            if self.parameterComboBox.SelectedItem:
                self._selected_item_before_search = self.parameterComboBox.SelectedItem
                try:
                    self._selected_param_name_before_search = self.parameterComboBox.SelectedItem.display_name
                except:
                    self._selected_param_name_before_search = None
            
            # Use Dispatcher to wait for popup to be fully rendered before finding SearchTextBox
            from System.Windows.Threading import DispatcherPriority
            self.parameterComboBox.Dispatcher.BeginInvoke(
                DispatcherPriority.Loaded,
                System.Action(self._initialize_search_textbox)
            )
        except Exception as ex:
            logger.error("Error initializing search filter: {}".format(str(ex)))
    
    def on_parameter_dropdown_closed(self, sender, args):
        """Handle parameter dropdown closed event - reset search flag."""
        try:
            # Reset search flag when dropdown closes
            self._is_searching = False
        except Exception as ex:
            logger.debug("Error handling dropdown closed: {}".format(str(ex)))
    
    def _initialize_search_textbox(self):
        """Initialize the search textbox after popup is loaded."""
        try:
            # Check if template exists
            if not self.parameterComboBox.Template:
                logger.warning("ComboBox has no template")
                return
            
            # Try to find Popup first
            popup = self.parameterComboBox.Template.FindName("Popup", self.parameterComboBox)
            
            # Try to find SearchTextBox directly
            search_textbox = self.parameterComboBox.Template.FindName("SearchTextBox", self.parameterComboBox)
            
            # If not found directly, try traversing visual tree from Popup
            if not search_textbox and popup:
                try:
                    # Try to find SearchTextBox in Popup's visual tree
                    from System.Windows.Media import VisualTreeHelper
                    if popup.Child:
                        # Try to find by name in visual tree
                        def find_child_by_name(parent, name):
                            if parent is None:
                                return None
                            if hasattr(parent, 'Name') and parent.Name == name:
                                return parent
                            for i in range(VisualTreeHelper.GetChildrenCount(parent)):
                                child = VisualTreeHelper.GetChild(parent, i)
                                result = find_child_by_name(child, name)
                                if result:
                                    return result
                            return None
                        
                        search_textbox = find_child_by_name(popup.Child, "SearchTextBox")
                except Exception as ex:
                    logger.debug("Error traversing visual tree: {}".format(str(ex)))
            
            if search_textbox:
                logger.debug("Found SearchTextBox successfully")
                # Store reference to search textbox for focus management
                self._search_textbox = search_textbox
                # Clear search text
                search_textbox.Text = ""
                # Wire up TextChanged event if not already done
                if not hasattr(self, '_search_textbox_wired'):
                    search_textbox.TextChanged += self.on_search_text_changed
                    search_textbox.KeyDown += self.on_search_textbox_keydown
                    search_textbox.PreviewKeyDown += self.on_search_textbox_preview_keydown
                    self._search_textbox_wired = True
                    logger.debug("Wired up SearchTextBox events")
                # Focus the search textbox
                search_textbox.Focus()
                logger.debug("Focused SearchTextBox")
            else:
                logger.warning("SearchTextBox not found in ComboBox template")
                if not popup:
                    logger.warning("Popup not found in ComboBox template")
        except Exception as ex:
            logger.error("Error finding search textbox: {}".format(str(ex)))
            import traceback
            logger.debug("Traceback: {}".format(traceback.format_exc()))
    
    def on_search_text_changed(self, sender, args):
        """Handle search text changed event - filter parameters."""
        try:
            search_text = sender.Text.lower().strip()
            
            # Clear current items - handle ItemsSource
            if self.parameterComboBox.ItemsSource is not None:
                self.parameterComboBox.ItemsSource = None
            self.parameterComboBox.Items.Clear()
            
            if not hasattr(self, 'all_parameters') or not self.all_parameters:
                return
            
            # Use ItemsSource for SearchableComboBox
            filtered_collection = ObservableCollection[Object]()
            
            if not search_text:
                # Show all parameters if search is empty
                for param_item in self.all_parameters:
                    filtered_collection.Add(param_item)
            else:
                # Filter parameters based on search text
                for param_item in self.all_parameters:
                    param_name = param_item.display_name.lower()
                    if search_text in param_name:
                        filtered_collection.Add(param_item)
            
            self.parameterComboBox.ItemsSource = filtered_collection
            
            # Try to restore selection if it matches the filter
            if hasattr(self, '_selected_item_before_search') and self._selected_item_before_search:
                try:
                    # Use ItemsSource since we just set it
                    for item in filtered_collection:
                        if item == self._selected_item_before_search:
                            self.parameterComboBox.SelectedItem = item
                            break
                except Exception as ex:
                    logger.debug("Error restoring selection in filter: {}".format(str(ex)))
            
            # Maintain focus on search textbox after filtering
            if hasattr(self, '_search_textbox') and self._search_textbox:
                from System.Windows.Threading import DispatcherPriority
                self.parameterComboBox.Dispatcher.BeginInvoke(
                    DispatcherPriority.Input,
                    System.Action(lambda: self._search_textbox.Focus())
                )
        except Exception as ex:
            logger.error("Error filtering parameters: {}".format(str(ex)))
            import traceback
            logger.debug("Traceback: {}".format(traceback.format_exc()))
    
    def on_search_textbox_preview_keydown(self, sender, args):
        """Handle PreviewKeyDown events - intercept arrow keys before they're handled."""
        try:
            from System.Windows.Input import Key
            
            # Handle Down arrow key - move focus to first item in results
            if args.Key == Key.Down:
                args.Handled = True
                # Use Dispatcher to set focus after current event completes
                from System.Windows.Threading import DispatcherPriority
                self.parameterComboBox.Dispatcher.BeginInvoke(
                    DispatcherPriority.Input,
                    System.Action(lambda: self._focus_first_item())
                )
                return
            
            # Handle Up arrow key
            elif args.Key == Key.Up:
                if self.parameterComboBox.Items.Count > 0:
                    args.Handled = True
                    from System.Windows.Threading import DispatcherPriority
                    self.parameterComboBox.Dispatcher.BeginInvoke(
                        DispatcherPriority.Input,
                        System.Action(lambda: self._focus_last_item())
                    )
                return
            
            # Handle Enter key
            elif args.Key == Key.Enter:
                if self.parameterComboBox.SelectedIndex >= 0:
                    args.Handled = True
                    self.parameterComboBox.IsDropDownOpen = False
                return
            
            # Handle Escape key
            elif args.Key == Key.Escape:
                args.Handled = True
                self.parameterComboBox.IsDropDownOpen = False
                return
                
        except Exception as ex:
            logger.error("Error handling search textbox preview keydown: {}".format(str(ex)))
    
    def _focus_first_item(self):
        """Focus the first ComboBoxItem."""
        try:
            first_item = self._find_first_combobox_item()
            if first_item:
                first_item.Focus()
                if self.parameterComboBox.Items.Count > 0:
                    self.parameterComboBox.SelectedIndex = 0
        except Exception as ex:
            logger.debug("Error focusing first item: {}".format(str(ex)))
    
    def _focus_last_item(self):
        """Focus the last ComboBoxItem."""
        try:
            last_item = self._find_last_combobox_item()
            if last_item:
                last_item.Focus()
                self.parameterComboBox.SelectedIndex = self.parameterComboBox.Items.Count - 1
        except Exception as ex:
            logger.debug("Error focusing last item: {}".format(str(ex)))
    
    def on_search_textbox_keydown(self, sender, args):
        """Handle keyboard events in search textbox - enable arrow key navigation."""
        # Note: PreviewKeyDown handles arrow keys, this handler is kept for compatibility
        # but arrow keys are intercepted in PreviewKeyDown before reaching here
        pass
    
    def _find_first_combobox_item(self):
        """Find the first ComboBoxItem in the dropdown."""
        try:
            from System.Windows.Controls import ComboBoxItem
            from System.Windows.Media import VisualTreeHelper
            
            if not self.parameterComboBox.Template:
                return None
            
            popup = self.parameterComboBox.Template.FindName("Popup", self.parameterComboBox)
            if not popup or not popup.Child:
                return None
            
            # Traverse visual tree to find first ComboBoxItem
            def find_first_combobox_item(parent):
                if parent is None:
                    return None
                if isinstance(parent, ComboBoxItem):
                    return parent
                child_count = VisualTreeHelper.GetChildrenCount(parent)
                for i in range(child_count):
                    child = VisualTreeHelper.GetChild(parent, i)
                    result = find_first_combobox_item(child)
                    if result:
                        return result
                return None
            
            return find_first_combobox_item(popup.Child)
        except Exception as ex:
            logger.debug("Error finding first ComboBoxItem: {}".format(str(ex)))
            return None
    
    def _find_last_combobox_item(self):
        """Find the last ComboBoxItem in the dropdown."""
        try:
            from System.Windows.Controls import ComboBoxItem
            from System.Windows.Media import VisualTreeHelper
            
            if not self.parameterComboBox.Template:
                return None
            
            popup = self.parameterComboBox.Template.FindName("Popup", self.parameterComboBox)
            if not popup or not popup.Child:
                return None
            
            # Traverse visual tree to find all ComboBoxItems and return the last one
            def find_all_combobox_items(parent, items_list):
                if parent is None:
                    return
                if isinstance(parent, ComboBoxItem):
                    items_list.append(parent)
                child_count = VisualTreeHelper.GetChildrenCount(parent)
                for i in range(child_count):
                    child = VisualTreeHelper.GetChild(parent, i)
                    find_all_combobox_items(child, items_list)
            
            items_list = []
            find_all_combobox_items(popup.Child, items_list)
            return items_list[-1] if items_list else None
        except Exception as ex:
            logger.debug("Error finding last ComboBoxItem: {}".format(str(ex)))
            return None


if __name__ == '__main__':
    try:
        window = MMISettingsWindow()
        window.ShowDialog()
    except Exception as e:
        logger.error("Error showing MMI Settings window: {}".format(str(e)))
        forms.alert("Failed to open MMI Settings window. See log for details.", title="Error")
