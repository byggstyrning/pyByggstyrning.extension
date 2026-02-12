# -*- coding: utf-8 -*-
"""Set NestedParentId on nested shared family instances for IFC export.

This tool identifies nested shared family instances and populates a shared parameter
with their parent's Element ID, enabling reliable parent-child matching in IFC
post-processing workflows.
"""
# pylint: disable=import-error,invalid-name,broad-except

__title__ = "Set Nested\nParent ID"
__author__ = "Byggstyrning AB"
__doc__ = """Set NestedParentId parameter on nested shared family instances.

This tool identifies nested shared family instances in the current document
and sets a 'NestedParentId' shared parameter with the parent's Element ID.

This enables reliable parent-child matching in IFC post-processing workflows,
solving the issue where Revit exports nested shared families as disconnected
elements (GitHub: Autodesk/revit-ifc#374).

USAGE:
1. Run this tool and select "Create Parameter" to add NestedParentId
2. Run "Set Values" to populate the parameter on all nested family instances
3. Export to IFC with "Export Revit property sets" enabled
4. Use IFC post-processing to merge nested elements based on NestedParentId

IFC EXPORT:
Enable "Export Revit property sets" in IFC export settings. The parameter
will appear under 'Pset_Revit_Parameters' or similar property set in the IFC file.

AUTOMATIC UPDATE:
Once the parameter is created, NestedParentId values are automatically updated
before each IFC export and saved to the model after export completes. This is
handled by the extension's startup script - no manual action required.

Note: "Export IFC common property sets" only exports predefined IFC standard
properties, not custom parameters. You must enable "Export Revit property sets"
for custom parameters like NestedParentId to be included.

PARAMETER SETTINGS:
The parameter is created as Text type with "Values can vary by group instance" 
enabled, which is essential for nested elements inside groups to have individual
values. (Integer parameters don't support varying by group instance in Revit.)

OPTIONS:
- Check Status: See if parameter exists and count nested instances
- Create Parameter: Creates the NestedParentId shared parameter
- Set Values (Dry Run): Preview what would be set without making changes
- Set Values: Populates NestedParentId on all nested shared family instances

CATEGORIES:
Bound to all family-based model element categories including: Doors, Windows,
Furniture, Generic Models, MEP equipment/fixtures/fittings, Structural elements,
Casework, Planting, and many more. Excludes system families (walls, floors, etc.).
"""

import sys
import os
import os.path as op

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    FilteredElementCollector, FamilyInstance, Transaction,
    BuiltInCategory, StorageType,
    CategorySet, Category, InstanceBinding, TypeBinding,
    ExternalDefinitionCreationOptions, DefinitionFile, DefinitionGroup,
    ElementId
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult

# Version-specific imports - ParameterType and BuiltInParameterGroup deprecated in Revit 2022+
try:
    from Autodesk.Revit.DB import SpecTypeId, GroupTypeId
    USE_FORGE_SCHEMA = True
except ImportError:
    from Autodesk.Revit.DB import ParameterType, BuiltInParameterGroup
    USE_FORGE_SCHEMA = False

# Import pyRevit modules
from pyrevit import script, forms, revit, HOST_APP

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
output = script.get_output()

# Configuration
# Simple parameter name - exports via "Export Revit property sets" checkbox in IFC export
# Will appear under Pset_Revit_Parameters or similar in IFC file
PARAM_NAME = "NestedParentId"
PARAM_GROUP_NAME = "IFC Export"
PARAM_TOOLTIP = "Element ID of the parent family instance (for IFC nested family merging)"

# Categories to process - all family-based model element categories
# These are categories that can contain nested shared families
# Excludes system families (walls, floors, roofs, ceilings, stairs, railings, etc.)
FAMILY_BASED_CATEGORIES = [
    # Common architectural
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_FurnitureSystems,
    BuiltInCategory.OST_Casework,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_Planting,
    BuiltInCategory.OST_Site,
    BuiltInCategory.OST_Entourage,
    BuiltInCategory.OST_Parking,
    
    # Curtain wall components
    BuiltInCategory.OST_CurtainWallPanels,
    BuiltInCategory.OST_CurtainWallMullions,
    
    # Structural
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_StructConnections,
    BuiltInCategory.OST_StructuralStiffener,
    BuiltInCategory.OST_Rebar,
    BuiltInCategory.OST_FabricAreas,
    BuiltInCategory.OST_FabricReinforcement,
    
    # MEP - Mechanical
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctInsulations,
    BuiltInCategory.OST_DuctLinings,
    
    # MEP - Plumbing
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_PipeInsulations,
    BuiltInCategory.OST_Sprinklers,
    
    # MEP - Electrical
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_LightingFixtures,
    BuiltInCategory.OST_LightingDevices,
    BuiltInCategory.OST_FireAlarmDevices,
    BuiltInCategory.OST_DataDevices,
    BuiltInCategory.OST_TelephoneDevices,
    BuiltInCategory.OST_NurseCallDevices,
    BuiltInCategory.OST_SecurityDevices,
    BuiltInCategory.OST_CommunicationDevices,
    BuiltInCategory.OST_ElectricalCircuit,
    BuiltInCategory.OST_CableTrayFitting,
    BuiltInCategory.OST_ConduitFitting,
    
    # Other model elements
    BuiltInCategory.OST_Columns,
    BuiltInCategory.OST_MassForm,
    BuiltInCategory.OST_DetailComponents,
    BuiltInCategory.OST_Signage,
    BuiltInCategory.OST_AudioVisualDevices,
    BuiltInCategory.OST_FoodServiceEquipment,
    BuiltInCategory.OST_MedicalEquipment,
    BuiltInCategory.OST_VerticalCirculation,
    BuiltInCategory.OST_Hardscape,
    BuiltInCategory.OST_BridgeFraming,
    BuiltInCategory.OST_BridgeDecks,
    BuiltInCategory.OST_BridgeCables,
    BuiltInCategory.OST_BridgeBearings,
    BuiltInCategory.OST_BridgeFoundations,
    BuiltInCategory.OST_BridgeTowers,
    BuiltInCategory.OST_BridgeArches,
    BuiltInCategory.OST_BridgePiers,
    BuiltInCategory.OST_BridgeAbutments,
    BuiltInCategory.OST_TemporaryStructure,
]


def get_revit_version():
    """Get Revit version as integer."""
    return int(HOST_APP.version)


def get_elementid_value(element_id):
    """Get ElementId value based on Revit version (API changed in 2024)."""
    version = get_revit_version()
    if version > 2023:
        return element_id.Value
    else:
        return element_id.IntegerValue


def get_shared_param_file(doc):
    """Get or create shared parameter file."""
    app = doc.Application
    
    # Check if a shared parameter file is already set
    if app.SharedParametersFilename:
        try:
            return app.OpenSharedParameterFile()
        except Exception as e:
            logger.warning("Could not open existing shared parameter file: {}".format(e))
    
    # No shared parameter file set - prompt user
    return None


def get_internal_definition(doc, param_name):
    """Get the InternalDefinition for a bound parameter.
    
    Args:
        doc: Revit document
        param_name: Name of the parameter
        
    Returns:
        InternalDefinition or None
    """
    bindings = doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    while iterator.MoveNext():
        if iterator.Key.Name == param_name:
            return iterator.Key
    return None


def create_shared_parameter(doc, param_name=PARAM_NAME):
    """Create the NestedParentId shared parameter and bind to categories.
    
    Args:
        doc: Revit document
        param_name: Name of the parameter to create
        
    Returns:
        tuple: (success: bool, message: str)
    """
    app = doc.Application
    
    # Get shared parameter file
    sp_file = get_shared_param_file(doc)
    if not sp_file:
        return False, "No shared parameter file set. Please set a shared parameter file in Revit settings."
    
    # Check if parameter already exists in file
    existing_def = None
    existing_is_text = False
    for group in sp_file.Groups:
        for definition in group.Definitions:
            if definition.Name == param_name:
                existing_def = definition
                # Check if it's the correct type (Text/String)
                try:
                    if USE_FORGE_SCHEMA:
                        param_type = definition.GetDataType()
                        existing_is_text = (param_type == SpecTypeId.String.Text)
                    else:
                        existing_is_text = (definition.ParameterType == ParameterType.Text)
                except:
                    existing_is_text = False
                break
        if existing_def:
            break
    
    # If existing definition is wrong type, we need to inform user
    if existing_def and not existing_is_text:
        return False, ("Parameter '{}' exists in shared parameter file but is not Text type. "
                      "Text type is required for 'Values can vary by group instance'. "
                      "Please delete '{}' from your shared parameter file and run again.").format(param_name, param_name)
    
    # Create definition if it doesn't exist
    if not existing_def:
        # Get or create group
        group = sp_file.Groups.get_Item(PARAM_GROUP_NAME)
        if not group:
            group = sp_file.Groups.Create(PARAM_GROUP_NAME)
        
        # Create the definition - API differs by version
        # Using Text (String) type because Integer doesn't support "Values can vary by group instance"
        try:
            if USE_FORGE_SCHEMA:
                # Revit 2022+ uses ForgeTypeId for parameter type
                options = ExternalDefinitionCreationOptions(param_name, SpecTypeId.String.Text)
                options.Description = PARAM_TOOLTIP
                existing_def = group.Definitions.Create(options)
            else:
                # Older versions use ParameterType enum
                options = ExternalDefinitionCreationOptions(param_name, ParameterType.Text)
                options.Description = PARAM_TOOLTIP
                existing_def = group.Definitions.Create(options)
        except Exception as e:
            return False, "Failed to create parameter definition: {}".format(e)
    
    # Check if already bound in this document
    bindings = doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    while iterator.MoveNext():
        if iterator.Key.Name == param_name:
            # Parameter already bound - check if VariesAcrossGroups is set
            internal_def = iterator.Key
            try:
                if hasattr(internal_def, 'VariesAcrossGroups') and not internal_def.VariesAcrossGroups:
                    # Need to enable VariesAcrossGroups
                    internal_def.SetAllowVaryBetweenGroups(doc, True)
                    return True, "Parameter '{}' already bound. Enabled 'Values can vary by group instance'.".format(param_name)
            except Exception as e:
                logger.debug("Could not set VariesAcrossGroups: {}".format(e))
            return True, "Parameter '{}' already exists and is bound.".format(param_name)
    
    # Create category set for binding
    cat_set = CategorySet()
    categories_added = []
    
    for bic in FAMILY_BASED_CATEGORIES:
        try:
            cat = Category.GetCategory(doc, bic)
            if cat and cat.AllowsBoundParameters:
                cat_set.Insert(cat)
                categories_added.append(cat.Name)
        except:
            continue
    
    if cat_set.IsEmpty:
        return False, "No valid categories found for parameter binding."
    
    # Create instance binding
    binding = InstanceBinding(cat_set)
    
    # Bind parameter to document - place in Identity Data group
    try:
        if USE_FORGE_SCHEMA:
            success = doc.ParameterBindings.Insert(existing_def, binding, GroupTypeId.IdentityData)
        else:
            success = doc.ParameterBindings.Insert(existing_def, binding, BuiltInParameterGroup.PG_IDENTITY_DATA)
        
        if not success:
            return False, "Failed to bind parameter to document."
        
        # After binding, enable "Values can vary by group instance"
        # This must be done after the parameter is bound to the document
        internal_def = get_internal_definition(doc, param_name)
        if internal_def:
            try:
                internal_def.SetAllowVaryBetweenGroups(doc, True)
                return True, "Parameter '{}' created and bound to: {}. 'Values can vary by group instance' enabled.".format(
                    param_name, ", ".join(categories_added))
            except Exception as e:
                logger.warning("Parameter bound but could not enable VariesAcrossGroups: {}".format(e))
                return True, "Parameter '{}' created and bound to: {}. (Note: Could not enable 'Values can vary by group instance' - please set manually)".format(
                    param_name, ", ".join(categories_added))
        else:
            return True, "Parameter '{}' created and bound to: {}".format(
                param_name, ", ".join(categories_added))
    except Exception as e:
        return False, "Error binding parameter: {}".format(e)


def check_parameter_exists(doc, param_name=PARAM_NAME):
    """Check if the parameter is bound in the document.
    
    Args:
        doc: Revit document
        param_name: Name of the parameter
        
    Returns:
        bool: True if parameter exists
    """
    bindings = doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    while iterator.MoveNext():
        if iterator.Key.Name == param_name:
            return True
    return False


def get_nested_family_instances(doc):
    """Get all nested shared family instances in the document.
    
    Nested shared families are FamilyInstances that have a SuperComponent
    (parent family instance they're nested inside).
    
    Args:
        doc: Revit document
        
    Returns:
        list: List of tuples (nested_instance, parent_instance)
    """
    nested_pairs = []
    
    # Collect all family instances
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    
    for fi in collector:
        try:
            # Check if this instance has a SuperComponent (is nested)
            super_component = fi.SuperComponent
            if super_component and isinstance(super_component, FamilyInstance):
                nested_pairs.append((fi, super_component))
        except Exception as e:
            logger.debug("Error checking SuperComponent for {}: {}".format(fi.Id, e))
            continue
    
    return nested_pairs


def set_nested_parent_ids(doc, dry_run=False, param_name=PARAM_NAME):
    """Set NestedParentId on all nested shared family instances.
    
    Args:
        doc: Revit document
        dry_run: If True, only report what would be done
        param_name: Name of the parameter to set
        
    Returns:
        dict: Results with counts and details
    """
    results = {
        "total_nested": 0,
        "successfully_set": 0,
        "already_set": 0,
        "param_not_found": 0,
        "param_readonly": 0,
        "errors": 0,
        "details": []
    }
    
    # Get all nested family instances
    nested_pairs = get_nested_family_instances(doc)
    results["total_nested"] = len(nested_pairs)
    
    if not nested_pairs:
        return results
    
    # Process each nested instance
    for nested_instance, parent_instance in nested_pairs:
        try:
            nested_id = get_elementid_value(nested_instance.Id)
            parent_id = get_elementid_value(parent_instance.Id)
            
            # Get family names for reporting
            nested_family = ""
            parent_family = ""
            try:
                if nested_instance.Symbol and nested_instance.Symbol.Family:
                    nested_family = nested_instance.Symbol.Family.Name
                if parent_instance.Symbol and parent_instance.Symbol.Family:
                    parent_family = parent_instance.Symbol.Family.Name
            except:
                pass
            
            detail = {
                "nested_id": nested_id,
                "nested_family": nested_family,
                "parent_id": parent_id,
                "parent_family": parent_family,
                "status": "pending"
            }
            
            # Try to get the parameter
            param = nested_instance.LookupParameter(param_name)
            
            if not param:
                detail["status"] = "param_not_found"
                results["param_not_found"] += 1
                results["details"].append(detail)
                continue
            
            if param.IsReadOnly:
                detail["status"] = "param_readonly"
                results["param_readonly"] += 1
                results["details"].append(detail)
                continue
            
            # Check current value (stored as string)
            parent_id_str = str(parent_id)
            current_value = param.AsString() if param.HasValue else ""
            if current_value == parent_id_str:
                detail["status"] = "already_set"
                results["already_set"] += 1
                results["details"].append(detail)
                continue
            
            # Set the value (unless dry run)
            if not dry_run:
                try:
                    param.Set(parent_id_str)
                    detail["status"] = "set"
                    results["successfully_set"] += 1
                except Exception as e:
                    detail["status"] = "error"
                    detail["error"] = str(e)
                    results["errors"] += 1
            else:
                detail["status"] = "would_set"
                results["successfully_set"] += 1
            
            results["details"].append(detail)
            
        except Exception as e:
            results["errors"] += 1
            logger.error("Error processing nested instance: {}".format(e))
    
    return results


def print_results(results, dry_run=False):
    """Print results to output."""
    mode = "DRY RUN" if dry_run else "EXECUTION"
    
    output.print_md("# NestedParentId {} Results".format(mode))
    output.print_md("")
    
    output.print_md("## Summary")
    output.print_md("- **Total nested instances found:** {}".format(results["total_nested"]))
    
    if dry_run:
        output.print_md("- **Would set:** {}".format(results["successfully_set"]))
    else:
        output.print_md("- **Successfully set:** {}".format(results["successfully_set"]))
    
    output.print_md("- **Already set correctly:** {}".format(results["already_set"]))
    output.print_md("- **Parameter not found:** {}".format(results["param_not_found"]))
    output.print_md("- **Parameter read-only:** {}".format(results["param_readonly"]))
    output.print_md("- **Errors:** {}".format(results["errors"]))
    output.print_md("")
    
    # Print details if there are any
    if results["details"]:
        output.print_md("## Details (first 50)")
        output.print_md("")
        
        for i, detail in enumerate(results["details"][:50]):
            status_icon = {
                "set": "✅",
                "would_set": "🔄",
                "already_set": "✓",
                "param_not_found": "⚠️",
                "param_readonly": "🔒",
                "error": "❌"
            }.get(detail["status"], "?")
            
            output.print_md("{} **{}** (ID: {}) → Parent: **{}** (ID: {}) - {}".format(
                status_icon,
                detail.get("nested_family", "Unknown"),
                detail.get("nested_id", "?"),
                detail.get("parent_family", "Unknown"),
                detail.get("parent_id", "?"),
                detail.get("status", "unknown")
            ))
        
        if len(results["details"]) > 50:
            output.print_md("")
            output.print_md("*...and {} more*".format(len(results["details"]) - 50))


def show_action_dialog():
    """Show dialog to select action."""
    options = {
        "1. Check Status": "See if parameter exists and count nested instances",
        "2. Create Parameter": "Create NestedParentId shared parameter",
        "3. Set Values (Dry Run)": "Preview what would be set",
        "4. Set Values": "Populate NestedParentId on nested instances"
    }
    
    selected = forms.CommandSwitchWindow.show(
        sorted(options.keys()),
        message="Select action for NestedParentId tool:"
    )
    
    if selected:
        # Extract the number from the selection (e.g., "1. Check Status" -> 1)
        return int(selected.split(".")[0])
    return None


def main():
    """Main entry point."""
    doc = revit.doc
    
    # Show action selection
    action = show_action_dialog()
    if not action:
        return
    
    output.print_md("# NestedParentId Tool")
    output.print_md("")
    
    if action == 1:
        # Check status
        output.print_md("## Status Check")
        output.print_md("")
        
        # Check if parameter exists
        param_exists = check_parameter_exists(doc, PARAM_NAME)
        if param_exists:
            output.print_md("✅ Parameter '{}' is bound to the document.".format(PARAM_NAME))
        else:
            output.print_md("❌ Parameter '{}' is NOT bound to the document.".format(PARAM_NAME))
            output.print_md("   Use option 2 to create it.")
        
        output.print_md("")
        
        # Count nested instances
        nested_pairs = get_nested_family_instances(doc)
        output.print_md("**Nested shared family instances found:** {}".format(len(nested_pairs)))
        
        if nested_pairs:
            output.print_md("")
            output.print_md("### Sample (first 10):")
            for nested, parent in nested_pairs[:10]:
                try:
                    nested_name = nested.Symbol.Family.Name if nested.Symbol and nested.Symbol.Family else "Unknown"
                    parent_name = parent.Symbol.Family.Name if parent.Symbol and parent.Symbol.Family else "Unknown"
                    output.print_md("- {} (ID: {}) → Parent: {} (ID: {})".format(
                        nested_name, 
                        get_elementid_value(nested.Id),
                        parent_name,
                        get_elementid_value(parent.Id)
                    ))
                except:
                    pass
    
    elif action == 2:
        # Create parameter
        output.print_md("## Create Parameter")
        output.print_md("")
        
        # Check if already exists
        if check_parameter_exists(doc, PARAM_NAME):
            output.print_md("⚠️ Parameter '{}' already exists in this document.".format(PARAM_NAME))
            return
        
        # Create with transaction
        with revit.Transaction("Create NestedParentId Parameter"):
            success, message = create_shared_parameter(doc, PARAM_NAME)
            
            if success:
                output.print_md("✅ {}".format(message))
            else:
                output.print_md("❌ {}".format(message))
    
    elif action == 3:
        # Dry run
        output.print_md("## Set Values (Dry Run)")
        output.print_md("")
        
        if not check_parameter_exists(doc, PARAM_NAME):
            output.print_md("❌ Parameter '{}' not found. Create it first (option 2).".format(PARAM_NAME))
            return
        
        results = set_nested_parent_ids(doc, dry_run=True, param_name=PARAM_NAME)
        print_results(results, dry_run=True)
    
    elif action == 4:
        # Set values
        output.print_md("## Set Values")
        output.print_md("")
        
        if not check_parameter_exists(doc, PARAM_NAME):
            output.print_md("❌ Parameter '{}' not found. Create it first (option 2).".format(PARAM_NAME))
            return
        
        # Confirm action
        nested_pairs = get_nested_family_instances(doc)
        if not nested_pairs:
            output.print_md("ℹ️ No nested shared family instances found in the document.")
            return
        
        confirm = TaskDialog.Show(
            "Confirm Set NestedParentId",
            "This will set NestedParentId on {} nested family instances.\n\nProceed?".format(len(nested_pairs)),
            TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
        )
        
        if confirm != TaskDialogResult.Yes:
            output.print_md("Operation cancelled.")
            return
        
        # Execute with transaction
        with revit.Transaction("Set NestedParentId Values"):
            results = set_nested_parent_ids(doc, dry_run=False, param_name=PARAM_NAME)
            print_results(results, dry_run=False)
    
    output.print_md("")
    output.print_md("---")
    output.print_md("*Tool completed.*")


if __name__ == "__main__":
    main()
