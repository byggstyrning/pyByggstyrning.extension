# -*- coding: utf-8 -*-
"""Configure 3D Zone parameter mappings.

Allows users to create, edit, and manage configurations for mapping
parameters from spatial elements to contained elements.
"""

__title__ = "Config"
__author__ = "Byggstyrning AB"
__doc__ = "Configure 3D Zone parameter mappings"

# Import standard libraries
import sys
import os

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

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
    from zone3d import config
except ImportError as e:
    logger.error("Failed to import zone3d libraries: {}".format(e))
    forms.alert("Failed to import required libraries. Check logs for details.")
    script.exit()

# --- Helper Functions ---

def get_category_options():
    """Get list of common BuiltInCategory options for selection."""
    return [
        ("Rooms", BuiltInCategory.OST_Rooms),
        ("Spaces", BuiltInCategory.OST_MEPSpaces),
        ("Areas", BuiltInCategory.OST_Areas),
        ("Mass", BuiltInCategory.OST_Mass),
        ("Generic Model", BuiltInCategory.OST_GenericModel),
        ("Walls", BuiltInCategory.OST_Walls),
        ("Doors", BuiltInCategory.OST_Doors),
        ("Windows", BuiltInCategory.OST_Windows),
        ("Furniture", BuiltInCategory.OST_Furniture),
        ("Equipment", BuiltInCategory.OST_PlumbingFixtures),
    ]

def get_parameters_from_element(doc, element):
    """Get list of parameter names from an element."""
    params = []
    if element:
        for param in element.Parameters:
            if param and not param.IsReadOnly:
                params.append(param.Definition.Name)
    return sorted(set(params))

def get_parameters_from_category(doc, category):
    """Get sample parameters from a category."""
    collector = FilteredElementCollector(doc)\
        .OfCategory(category)\
        .WhereElementIsNotElementType()\
        .FirstElement()
    
    if collector:
        return get_parameters_from_element(doc, collector)
    return []

# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Load existing configurations
    configs = config.load_configs(doc)
    
    # Main menu
    options = [
        "View Configurations",
        "Add New Configuration",
        "Edit Configuration",
        "Delete Configuration",
        "Reorder Configurations"
    ]
    
    choice = forms.CommandSwitchWindow.show(
        options,
        message="Select an action:",
        title="3D Zone Configuration"
    )
    
    if not choice:
        script.exit()
    
    if choice == "View Configurations":
        if not configs:
            forms.alert("No configurations found.", title="No Configurations")
            script.exit()
        
        config_list = []
        for cfg in sorted(configs, key=lambda x: x.get("order", 999)):
            name = cfg.get("name", "Unknown")
            order = cfg.get("order", 0)
            enabled = "Enabled" if cfg.get("enabled", False) else "Disabled"
            source_cats = cfg.get("source_categories", [])
            source_cat_names = [str(cat) for cat in source_cats]
            source_params = cfg.get("source_params", [])
            target_params = cfg.get("target_params", [])
            
            # Build parameter mapping display
            param_mappings = []
            for i, (src_param, tgt_param) in enumerate(zip(source_params, target_params)):
                param_mappings.append("  {} -> {}".format(src_param, tgt_param))
            
            param_display = "\n".join(param_mappings) if param_mappings else "  (no parameters)"
            
            config_list.append("Order {}: {} ({})\n  Source Categories: {}\n  Parameter Mappings:\n{}".format(
                order, name, enabled,
                ", ".join(source_cat_names[:3]),
                param_display
            ))
        
        forms.alert("\n\n".join(config_list), title="3D Zone Configurations")
    
    elif choice == "Add New Configuration":
        # Get source categories first
        category_options = get_category_options()
        source_cat_choices = forms.SelectFromList.show(
            [opt[0] for opt in category_options],
            multiselect=True,
            title="Select Source Categories"
        )
        
        if not source_cat_choices:
            forms.alert("Source categories are required.", title="Error")
            script.exit()
        
        # Map display names to BuiltInCategory
        cat_map = {opt[0]: opt[1] for opt in category_options}
        source_categories = [cat_map[name] for name in source_cat_choices]
        
        # Generate default name from selected categories
        default_name = "{} to Elements".format(" + ".join(source_cat_choices))
        
        # Get configuration name (writable string input)
        name = forms.ask_for_string(
            default=default_name,
            prompt="Configuration Name:",
            title="New Configuration"
        )
        
        if not name:
            name = default_name
        
        # Get sample element to list parameters
        sample_collector = FilteredElementCollector(doc)\
            .OfCategory(source_categories[0])\
            .WhereElementIsNotElementType()\
            .FirstElement()
        
        if not sample_collector:
            forms.alert("No elements found in selected category.", title="Error")
            script.exit()
        
        source_params = get_parameters_from_element(doc, sample_collector)
        
        if not source_params:
            forms.alert("No parameters found on source elements.", title="Error")
            script.exit()
        
        # Select source parameters
        selected_source_params = forms.SelectFromList.show(
            source_params,
            multiselect=True,
            title="Select Source Parameters"
        )
        
        if not selected_source_params:
            forms.alert("Source parameters are required.", title="Error")
            script.exit()
        
        # Get target parameters (use same list for now)
        selected_target_params = forms.SelectFromList.show(
            source_params,
            multiselect=True,
            title="Select Target Parameters (same order as source)",
            default=selected_source_params
        )
        
        if len(selected_target_params) != len(selected_source_params):
            forms.alert("Target parameters count must match source parameters.", title="Error")
            script.exit()
        
        # Get target filter categories (optional)
        target_filter_choices = forms.SelectFromList.show(
            [opt[0] for opt in category_options],
            multiselect=True,
            title="Select Target Filter Categories (optional)",
            button_name="Skip"
        )
        
        target_filter_categories = []
        if target_filter_choices:
            target_filter_categories = [cat_map[name] for name in target_filter_choices]
        
        # Get order
        next_order = config.get_next_order(doc)
        order = forms.ask_for_one_item(
            [str(i) for i in range(1, next_order + 2)],
            default=str(next_order),
            prompt="Execution Order:",
            title="Configuration Order"
        )
        
        if not order:
            order = next_order
        else:
            try:
                order = int(order)
            except:
                order = next_order
        
        # Create configuration
        new_config = {
            "id": config.generate_config_id(),
            "name": name,
            "order": order,
            "enabled": True,
            "source_categories": source_categories,
            "source_params": selected_source_params,
            "target_params": selected_target_params,
            "target_filter_categories": target_filter_categories
        }
        
        # Add to list
        configs.append(new_config)
        
        # Save
        if config.save_configs(doc, configs):
            forms.alert("Configuration '{}' created successfully.".format(name), title="Success")
        else:
            forms.alert("Failed to save configuration.", title="Error")
    
    elif choice == "Edit Configuration":
        if not configs:
            forms.alert("No configurations found.", title="No Configurations")
            script.exit()
        
        # Select configuration
        config_names = ["{} (Order: {})".format(cfg.get("name"), cfg.get("order")) for cfg in configs]
        selected = forms.SelectFromList.show(config_names, title="Select Configuration to Edit")
        
        if not selected:
            script.exit()
        
        selected_index = config_names.index(selected)
        cfg = configs[selected_index]
        
        # Simple edit: toggle enabled
        new_enabled = forms.alert(
            "Configuration: {}\n\nCurrent status: {}\n\nToggle enabled status?".format(
                cfg.get("name"), "Enabled" if cfg.get("enabled") else "Disabled"
            ),
            title="Edit Configuration",
            ok=False,
            yes=True,
            no=True
        )
        
        if new_enabled is not None:
            cfg["enabled"] = new_enabled
            if config.save_configs(doc, configs):
                forms.alert("Configuration updated.", title="Success")
            else:
                forms.alert("Failed to save changes.", title="Error")
    
    elif choice == "Delete Configuration":
        if not configs:
            forms.alert("No configurations found.", title="No Configurations")
            script.exit()
        
        # Select configuration
        config_names = ["{} (Order: {})".format(cfg.get("name"), cfg.get("order")) for cfg in configs]
        selected = forms.SelectFromList.show(config_names, title="Select Configuration to Delete")
        
        if not selected:
            script.exit()
        
        selected_index = config_names.index(selected)
        cfg = configs[selected_index]
        
        # Confirm deletion
        if forms.alert(
            "Delete configuration '{}'?".format(cfg.get("name")),
            title="Confirm Deletion",
            ok=False,
            yes=True,
            no=True
        ):
            config_id = cfg.get("id")
            if config.delete_config(doc, config_id):
                forms.alert("Configuration deleted.", title="Success")
            else:
                forms.alert("Failed to delete configuration.", title="Error")
    
    elif choice == "Reorder Configurations":
        if not configs:
            forms.alert("No configurations found.", title="No Configurations")
            script.exit()
        
        # Show current order
        sorted_configs = sorted(configs, key=lambda x: x.get("order", 999))
        config_list = []
        for i, cfg in enumerate(sorted_configs):
            config_list.append("{}. {} (current order: {})".format(
                i+1, cfg.get("name"), cfg.get("order")
            ))
        
        forms.alert(
            "Current order:\n\n{}".format("\n".join(config_list)),
            title="Configuration Order"
        )
        
        # Simple reordering: ask for new order numbers
        forms.alert(
            "To reorder, edit configurations individually or recreate them in desired order.\n\n"
            "Full reordering UI coming in future update.",
            title="Reorder Configurations"
        )

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.

