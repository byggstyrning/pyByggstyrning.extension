# -*- coding: utf-8 -*-
"""Write parameters to elements based on configured zone mappings.

Executes all enabled configurations in order, writing parameters from
spatial elements (Rooms, Spaces, Areas, Mass/Generic Model) to contained elements.
"""

__title__ = "Write"
__author__ = "Byggstyrning AB"
__doc__ = "Execute all enabled 3D Zone configurations to write parameters to elements"

# Import standard libraries
import sys
import os

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from Autodesk.Revit.DB import *

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Import WPF components
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

# Add the extension directory to the path
import os.path as op
script_path = __file__
pushbutton_dir = op.dirname(script_path)
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Initialize logger
logger = script.get_logger()

# Import zone3d libraries
try:
    from zone3d import config, core, containment
    from Autodesk.Revit.DB import BuiltInCategory
except ImportError as e:
    logger.error("Failed to import zone3d libraries: {}".format(e))
    forms.alert("Failed to import required libraries. Check logs for details.")
    script.exit()

# --- Configuration Selection ViewModel ---

class ConfigItemViewModel(INotifyPropertyChanged):
    """ViewModel for configuration item with checkbox binding."""
    def __init__(self, config_dict):
        self.config_dict = config_dict
        self.name = config_dict.get("name", "Unknown")
        self.order = config_dict.get("order", 0)
        self._is_selected = config_dict.get("enabled", False)
        self.display_name = self.name
        # Initialize the event handler list
        self._property_changed_handlers = []
    
    def add_PropertyChanged(self, handler):
        """Add a PropertyChanged event handler."""
        if handler is not None:
            self._property_changed_handlers.append(handler)
    
    def remove_PropertyChanged(self, handler):
        """Remove a PropertyChanged event handler."""
        if handler is not None and handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)
    
    def _notify_property_changed(self, property_name):
        """Notify that a property has changed."""
        if self._property_changed_handlers:
            args = PropertyChangedEventArgs(property_name)
            for handler in self._property_changed_handlers:
                try:
                    handler(self, args)
                except Exception:
                    pass
    
    @property
    def IsSelected(self):
        return self._is_selected
    
    @IsSelected.setter
    def IsSelected(self, value):
        if self._is_selected != value:
            self._is_selected = value
            self._notify_property_changed("IsSelected")
    
    @property
    def DisplayName(self):
        return self.display_name

# --- Configuration Selector Window ---

class ConfigSelectorWindow(forms.WPFWindow):
    """Window for selecting configurations to execute."""
    
    def __init__(self, configs):
        """Initialize the configuration selector window.
        
        Args:
            configs: List of configuration dictionaries
        """
        # Load XAML file
        xaml_path = op.join(pushbutton_dir, 'ConfigSelector.xaml')
        forms.WPFWindow.__init__(self, xaml_path)
        
        # Store selected configs (will be populated when Write is clicked)
        self.selected_configs = None
        
        # Create ObservableCollection and populate with ViewModels
        self.config_items = ObservableCollection[object]()
        for cfg in configs:
            view_model = ConfigItemViewModel(cfg)
            self.config_items.Add(view_model)
        
        # Bind collection to ListView
        self.configsListView.ItemsSource = self.config_items
        
        # Set up event handlers
        self.writeButton.Click += self.write_button_click
        self.cancelButton.Click += self.cancel_button_click
    
    def write_button_click(self, sender, args):
        """Handle Write button click - collect checked configs and close."""
        # Collect all checked configurations
        selected_configs = []
        for item in self.config_items:
            if item.IsSelected:
                selected_configs.append(item.config_dict)
        
        # Sort by order
        selected_configs.sort(key=lambda x: x.get("order", 0))
        
        # Store result and close window
        self.selected_configs = selected_configs
        self.Close()
    
    def cancel_button_click(self, sender, args):
        """Handle Cancel button click - close without selection."""
        self.selected_configs = None
        self.Close()

# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    class BatchProgressAdapter(object):
        """Adapter to map per-config progress (0..N) into one batch progress bar.
        
        `zone3d.core` reports progress as (current, total) for a single config.
        This adapter scales that into an overall 0..100% progress bar.
        """
        def __init__(self, progress_bar, config_idx, total_configs, steps_per_config=100):
            self._pb = progress_bar
            self._config_idx = int(config_idx)
            self._total_configs = int(total_configs) if total_configs else 1
            # steps_per_config is kept for backward compatibility / tuning,
            # but output is always mapped to a 0..100 total to match typical progress UIs.
            self._steps_per_config = int(steps_per_config) if steps_per_config else 100
        
        def update_progress(self, current, total):
            try:
                cur = int(current) if current is not None else 0
                tot = int(total) if total is not None else 0
            except Exception:
                cur = 0
                tot = 0
            
            if tot <= 0:
                local = 0
            else:
                ratio = float(cur) / float(tot)
                if ratio < 0.0:
                    ratio = 0.0
                elif ratio > 1.0:
                    ratio = 1.0
                local = int(round(ratio * self._steps_per_config))
                if local < 0:
                    local = 0
                elif local > self._steps_per_config:
                    local = self._steps_per_config
            
            overall_ratio = (float(self._config_idx) + (float(local) / float(self._steps_per_config))) / float(self._total_configs)
            if overall_ratio < 0.0:
                overall_ratio = 0.0
            elif overall_ratio > 1.0:
                overall_ratio = 1.0
            
            global_current = int(round(overall_ratio * 100.0))
            if global_current < 0:
                global_current = 0
            elif global_current > 100:
                global_current = 100
            
            self._pb.update_progress(global_current, 100)
        
        def mark_complete(self):
            # Mark this config as finished (advance to the start of the next slice)
            overall_ratio = float(self._config_idx + 1) / float(self._total_configs)
            if overall_ratio < 0.0:
                overall_ratio = 0.0
            elif overall_ratio > 1.0:
                overall_ratio = 1.0
            global_current = int(round(overall_ratio * 100.0))
            if global_current < 0:
                global_current = 0
            elif global_current > 100:
                global_current = 100
            self._pb.update_progress(global_current, 100)
    
    # Load all configurations (not just enabled)
    all_configs = config.load_configs(doc)
    
    if not all_configs:
        forms.alert(
            "No configurations found.\n\n"
            "Please configure mappings using the Config button first.",
            title="No Configurations",
            exitscript=True
        )
    
    # Sort configs by order for display
    all_configs.sort(key=lambda x: (x.get("order", 0), x.get("name", "")))
    
    # Show configuration selector window
    selector_window = ConfigSelectorWindow(all_configs)
    selector_window.ShowDialog()
    
    # Get selected configurations
    selected_configs = selector_window.selected_configs
    
    # Exit if cancelled or no selection
    if not selected_configs:
        script.exit()
    
    try:
        # Execute selected configurations manually (similar to execute_all_configurations)
        summary = {
            "total_configs": len(selected_configs),
            "config_results": [],
            "total_elements_updated": 0,
            "total_elements_already_correct": 0,
            "total_parameters_copied": 0,
            "total_parameters_already_correct": 0
        }
        
        # Clear geometry cache once at the start (not per config)
        containment.clear_geometry_cache()
        
        # Start single transaction for all selected configurations
        from Autodesk.Revit.DB import Transaction
        transaction = Transaction(doc, "3D Zone: Selected Configurations")
        
        try:
            transaction.Start()
            
            # Single progress bar for the entire batch
            batch_title = "3D Zone: Writing ({} configuration(s))".format(len(selected_configs))
            with forms.ProgressBar(title=batch_title) as pb:
                total_configs = len(selected_configs) if selected_configs else 1
                pb.update_progress(0, 100)
                
                # Process each selected configuration
                for config_idx, zone_config in enumerate(selected_configs):
                    config_name = zone_config.get("name", "Unknown")
                    config_order = zone_config.get("order", 0)
                    
                    logger.debug("[BATCH] Running configuration {}/{}: {}".format(
                        config_idx + 1, total_configs, config_name
                    ))
                    
                    adapter = BatchProgressAdapter(pb, config_idx, total_configs, steps_per_config=100)
                    
                    # Execute configuration with adapted progress reporter
                    # Use no-transaction path since we're already in a transaction
                    # Skip cache clear since we cleared it once at the start
                    result = core.execute_configuration(
                        doc, zone_config, adapter,
                        view_id=None, force_transaction=False, use_subtransaction=False,
                        skip_cache_clear=True
                    )
                    adapter.mark_complete()
                    
                    result["config_name"] = config_name
                    result["config_order"] = config_order
                    
                    # Add to summary
                    summary["config_results"].append(result)
                    summary["total_elements_updated"] += result.get("elements_updated", 0)
                    summary["total_elements_already_correct"] += result.get("elements_already_correct", 0)
                    summary["total_parameters_copied"] += result.get("parameters_copied", 0)
                    summary["total_parameters_already_correct"] += result.get("parameters_already_correct", 0)
            
            # Commit single transaction for all configurations
            transaction.Commit()
        
        except Exception as e:
            # Rollback on error
            try:
                if transaction.HasStarted() and not transaction.HasEnded():
                    transaction.RollBack()
            except:
                pass
            raise
        finally:
            # Ensure transaction is disposed
            try:
                if transaction.HasStarted() and not transaction.HasEnded():
                    transaction.RollBack()
                transaction.Dispose()
            except:
                pass
        
        # Build results message
        results_text = "Execution Complete\n\n"
        results_text += "Total Configurations: {}\n".format(summary["total_configs"])
        results_text += "Total Elements Updated: {} ({} parameters)\n".format(
            summary["total_elements_updated"],
            summary["total_parameters_copied"]
        )
        if summary.get("total_elements_already_correct", 0) > 0:
            results_text += "Total Elements Already Correct: {} ({} parameters)\n".format(
                summary["total_elements_already_correct"],
                summary["total_parameters_already_correct"]
            )
        results_text += "\n"
        
        # Add per-configuration results
        if summary["config_results"]:
            results_text += "Per Configuration:\n"
            for result in summary["config_results"]:
                config_name = result.get("config_name", "Unknown")
                elements_updated = result.get("elements_updated", 0)
                elements_already_correct = result.get("elements_already_correct", 0)
                params_copied = result.get("parameters_copied", 0)
                params_already_correct = result.get("parameters_already_correct", 0)
                errors = result.get("errors", [])
                
                results_text += "\n{}:\n".format(config_name)
                results_text += "  Elements Updated: {} ({} parameters)\n".format(elements_updated, params_copied)
                if elements_already_correct > 0:
                    results_text += "  Elements Already Correct: {} ({} parameters)\n".format(
                        elements_already_correct, params_already_correct
                    )
                
                if errors:
                    results_text += "  Errors: {}\n".format(len(errors))
                    for error in errors[:3]:  # Show first 3 errors
                        results_text += "    - {}\n".format(error[:60])
                    if len(errors) > 3:
                        results_text += "    ... and {} more\n".format(len(errors) - 3)
        
        # Show results using pyRevit alert
        forms.alert(results_text, title="3D Zone Write Results")
    
    except Exception as e:
        error_msg = "Error executing configurations: {}".format(str(e))
        logger.error(error_msg)
        forms.alert(error_msg, title="Error", exitscript=True)

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.

