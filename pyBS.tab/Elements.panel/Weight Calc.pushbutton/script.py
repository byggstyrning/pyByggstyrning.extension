# -*- coding: utf-8 -*-
"""Weight Calculator - Calculate weight for structural elements based on volume and density."""

__title__ = "Weight Calc"
__author__ = "Byggstyrning AB"
__doc__ = "Calculate weight for structural elements based on volume and density"

import sys
import os.path as op
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *

# Revit version compatibility
try:
    from Autodesk.Revit.DB import UnitTypeId
    USE_FORGE_UNITS = True
except ImportError:
    from Autodesk.Revit.DB import DisplayUnitType
    USE_FORGE_UNITS = False

from pyrevit import script, forms, revit

# Add lib path
script_path = __file__
extension_dir = op.dirname(op.dirname(op.dirname(op.dirname(script_path))))
lib_path = op.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)

from revit.revit_utils import (
    get_available_parameters, 
    set_parameter_value,
    is_element_editable,
    is_parameter_writable
)

logger = script.get_logger()
WEIGHT_PARAM_DEFAULT = "Weight"

# Global cache for material densities to improve performance
_material_density_cache = {}


def convert_volume_to_cubic_meters(volume_internal):
    """Convert volume from Revit internal units to cubic meters."""
    try:
        if USE_FORGE_UNITS:
            return UnitUtils.ConvertFromInternalUnits(volume_internal, UnitTypeId.CubicMeters)
        else:
            return UnitUtils.ConvertFromInternalUnits(volume_internal, DisplayUnitType.DUT_CUBIC_METERS)
    except:
        return volume_internal * 0.0283168


def convert_density_to_kg_per_cubic_meter(density_internal):
    """Convert density from Revit internal units to kg/m³."""
    try:
        if USE_FORGE_UNITS:
            return UnitUtils.ConvertFromInternalUnits(density_internal, UnitTypeId.KilogramsPerCubicMeter)
        else:
            return UnitUtils.ConvertFromInternalUnits(density_internal, DisplayUnitType.DUT_KILOGRAMS_PER_CUBIC_METER)
    except:
        return density_internal * 16.0185


def get_structural_elements_in_view(doc, view):
    """Get all structural elements with volume in the active view."""
    elements = []
    collector = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements()
    
    for element in collector:
        try:
            if not element.Category:
                continue
            
            volume_param = element.LookupParameter("Volume")
            if not volume_param or not volume_param.HasValue or volume_param.AsDouble() <= 0:
                continue
            
            # Check if structural
            is_structural = False
            structural_param = element.LookupParameter("Structural")
            if structural_param and structural_param.HasValue and structural_param.StorageType == StorageType.Integer:
                is_structural = structural_param.AsInteger() == 1
            
            if not is_structural and element.Category.Name and element.Category.Name.startswith("Structural"):
                is_structural = True
            
            if is_structural:
                elements.append(element)
        except:
            continue
    
    return elements


def get_material_density(material):
    """Get density from material in kg/m³ (with caching)."""
    if not material:
        return None
    
    # Check cache first
    material_id = material.Id.IntegerValue
    if material_id in _material_density_cache:
        return _material_density_cache[material_id]
    
    density = None
    
    try:
        doc = material.Document
        
        # Try Structural Asset
        structural_asset_id = material.StructuralAssetId
        if structural_asset_id and structural_asset_id != ElementId.InvalidElementId:
            try:
                pse = doc.GetElement(structural_asset_id)
                if pse:
                    structural_asset = pse.GetStructuralAsset()
                    if structural_asset and hasattr(structural_asset, 'Density'):
                        density_internal = structural_asset.Density
                        if density_internal and density_internal > 0:
                            density = convert_density_to_kg_per_cubic_meter(density_internal)
                            _material_density_cache[material_id] = density
                            return density
            except:
                pass
        
        # Try Thermal Asset
        thermal_asset_id = material.ThermalAssetId
        if thermal_asset_id and thermal_asset_id != ElementId.InvalidElementId:
            try:
                pse = doc.GetElement(thermal_asset_id)
                if pse:
                    thermal_asset = pse.GetThermalAsset()
                    if thermal_asset and hasattr(thermal_asset, 'Density'):
                        density_internal = thermal_asset.Density
                        if density_internal and density_internal > 0:
                            density = convert_density_to_kg_per_cubic_meter(density_internal)
                            _material_density_cache[material_id] = density
                            return density
            except:
                pass
        
        # Try Built-in Parameter
        try:
            density_param = material.get_Parameter(BuiltInParameter.PHY_MATERIAL_PARAM_STRUCTURAL_DENSITY)
            if density_param and density_param.HasValue:
                density_internal = density_param.AsDouble()
                if density_internal > 0:
                    density = convert_density_to_kg_per_cubic_meter(density_internal)
                    _material_density_cache[material_id] = density
                    return density
        except:
            pass
        
        # Try older API
        if hasattr(material, 'Density'):
            density_value = material.Density
            if density_value and density_value > 0:
                _material_density_cache[material_id] = density_value
                return density_value
    except:
        pass
    
    # Cache None result to avoid repeated lookups
    _material_density_cache[material_id] = None
    return None


def calculate_density_from_layers(element):
    """Calculate weighted average density from material layers."""
    try:
        doc = element.Document
        type_id = element.GetTypeId()
        if type_id == ElementId.InvalidElementId:
            return None
            
        element_type = doc.GetElement(type_id)
        if not element_type or not hasattr(element_type, 'GetCompoundStructure'):
            return None
        
        compound_structure = element_type.GetCompoundStructure()
        if not compound_structure:
            return None
        
        total_density_thickness = 0.0
        total_thickness = 0.0
        
        for layer in compound_structure.GetLayers():
            layer_thickness = layer.Width
            material_id = layer.MaterialId
            
            if material_id == ElementId.InvalidElementId:
                continue
            
            material = doc.GetElement(material_id)
            if not material:
                continue
            
            density = get_material_density(material)
            if density and density > 0:
                try:
                    if USE_FORGE_UNITS:
                        thickness_meters = UnitUtils.ConvertFromInternalUnits(layer_thickness, UnitTypeId.Meters)
                    else:
                        thickness_meters = UnitUtils.ConvertFromInternalUnits(layer_thickness, DisplayUnitType.DUT_METERS)
                except:
                    thickness_meters = layer_thickness * 0.3048
                
                total_density_thickness += density * thickness_meters
                total_thickness += thickness_meters
        
        if total_thickness > 0:
            return total_density_thickness / total_thickness
    except:
        pass
    
    return None


def calculate_element_weight(element):
    """Calculate weight for a structural element. Returns (weight in kg, method) or (None, error)."""
    try:
        volume_param = element.LookupParameter("Volume")
        if not volume_param or not volume_param.HasValue:
            return None, "No volume parameter"
        
        volume_internal = volume_param.AsDouble()
        if volume_internal <= 0:
            return None, "Volume is zero or negative"
        
        volume_cubic_meters = convert_volume_to_cubic_meters(volume_internal)
        if volume_cubic_meters <= 0:
            return None, "Volume conversion error"
        
        doc = element.Document
        density = None
        method = "unknown"
        
        # Try Structural Material parameter
        structural_material_param = element.LookupParameter("Structural Material")
        if structural_material_param and structural_material_param.HasValue:
            material_id = structural_material_param.AsElementId()
            if material_id != ElementId.InvalidElementId:
                material = doc.GetElement(material_id)
                if material:
                    density = get_material_density(material)
                    if density:
                        method = "Structural Material"
        
        # Try Material parameter
        if not density:
            material_param = element.LookupParameter("Material")
            if material_param and material_param.HasValue:
                material_id = material_param.AsElementId()
                if material_id != ElementId.InvalidElementId:
                    material = doc.GetElement(material_id)
                    if material:
                        density = get_material_density(material)
                        if density:
                            method = "Material parameter"
        
        # Try material layers
        if not density:
            weighted_density = calculate_density_from_layers(element)
            if weighted_density:
                density = weighted_density
                method = "Material layers"
        
        # Calculate weight
        if density and density > 0:
            weight = volume_cubic_meters * density
            return weight, method
        elif density == 0:
            return None, "Material density is zero"
        else:
            return None, "No material density found"
            
    except Exception as ex:
        logger.error("Error calculating weight for element {}: {}".format(element.Id, ex))
        return None, "Calculation error"


def print_failed_elements_table(failed_elements):
    """Print a table of failed elements with clickable links."""
    output_window = script.get_output()
    output_window.print_md("## ⚠️ Failed Elements - No Material Density Found")
    output_window.print_md("Click element ID to zoom to element in Revit")
    output_window.print_md("---")
    
    # Group by category for better overview
    by_category = {}
    for item in failed_elements:
        cat = item['category']
        if cat not in by_category:
            by_category[cat] = []
        by_category[cat].append(item)
    
    # Print by category
    for category, items in sorted(by_category.items()):
        output_window.print_md("### {} ({} elements)".format(category, len(items)))
        output_window.print_md("")
        
        # Print each element with clickable link
        for item in items:
            element = item['element']
            reason = item['reason']
            type_name = item.get('type_name', 'Unknown Type')
            # Use linkify method from output window
            element_link = output_window.linkify(element.Id)
            output_window.print_md("- **{} {}**: {}".format(type_name, element_link, reason))
        
        output_window.print_md("")  # Empty line between categories


def print_skipped_elements_table(skipped_elements):
    """Print a table of skipped elements with clickable links."""
    output_window = script.get_output()
    output_window.print_md("## ⏭️ Skipped Elements - Not Editable")
    output_window.print_md("These elements cannot be modified (read-only parameter or owned by another user)")
    output_window.print_md("---")
    
    # Group by category for better overview
    by_category = {}
    for item in skipped_elements:
        cat = item['category']
        if cat not in by_category:
            by_category[cat] = []
        by_category[cat].append(item)
    
    # Print by category
    for category, items in sorted(by_category.items()):
        output_window.print_md("### {} ({} elements)".format(category, len(items)))
        output_window.print_md("")
        
        # Print each element with clickable link
        for item in items:
            element = item['element']
            reason = item['reason']
            type_name = item.get('type_name', 'Unknown Type')
            # Use linkify method from output window
            element_link = output_window.linkify(element.Id)
            output_window.print_md("- **{} {}**: {}".format(type_name, element_link, reason))
        
        output_window.print_md("")  # Empty line between categories


def select_weight_parameter(doc):
    """Let user select which parameter to use for weight."""
    available_parameters = get_available_parameters()
    
    if not available_parameters:
        forms.alert("No writable parameters found in the project.", title="Weight Calculator")
        return None
    
    default_param = None
    for param in available_parameters:
        if param.lower() == WEIGHT_PARAM_DEFAULT.lower():
            default_param = param
            break
    
    selected_parameter = forms.ask_for_one_item(
        available_parameters,
        default=default_param,
        prompt="Select parameter to store weight values (kg):",
        title="Weight Calculator"
    )
    
    return selected_parameter


def main():
    """Main execution function."""
    # Clear material density cache for fresh run
    global _material_density_cache
    _material_density_cache = {}
    
    doc = revit.doc
    
    if not doc.ActiveView:
        forms.alert("No active view found.", title="Weight Calculator")
        return
    
    view = doc.ActiveView
    elements = get_structural_elements_in_view(doc, view)
    
    if not elements:
        forms.alert(
            "No structural elements with volume found in the active view.\n\n"
            "Elements must:\n"
            "- Have a Volume parameter with a value\n"
            "- Be structural (Structural=Yes or category starting with 'Structural')",
            title="Weight Calculator"
        )
        return
    
    weight_parameter = select_weight_parameter(doc)
    if not weight_parameter:
        return
    
    results = []
    failed_elements = []
    skipped_elements = []
    
    # Determine batch size for progress updates (update every N elements)
    total_count = len(elements)
    batch_size = max(1, total_count // 100) if total_count > 100 else 10
    
    # Cache for element types to avoid repeated lookups
    type_cache = {}
    
    def get_element_info(element):
        """Get element type name and category (only when needed for reporting)."""
        try:
            type_id = element.GetTypeId()
            if type_id not in type_cache:
                element_type = doc.GetElement(type_id)
                type_cache[type_id] = element_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if element_type else "Unknown Type"
            type_name = type_cache[type_id]
        except:
            type_name = "Unknown Type"
        
        category_name = element.Category.Name if element.Category else "No Category"
        return type_name, category_name
    
    with forms.ProgressBar(title="Calculating Weights ({} elements)".format(total_count)) as pb:
        for idx, element in enumerate(elements):
            # Batch progress updates for better performance
            if idx % batch_size == 0 or idx == total_count - 1:
                pb.update_progress(idx + 1, total_count)
            
            # Check if element is editable (worksharing ownership check)
            is_editable, edit_reason = is_element_editable(doc, element)
            if not is_editable:
                type_name, category_name = get_element_info(element)
                skipped_elements.append({
                    'element': element,
                    'reason': edit_reason,
                    'category': category_name,
                    'type_name': type_name
                })
                continue
            
            # Check if parameter is writable
            is_writable, write_reason, param = is_parameter_writable(element, weight_parameter)
            if not is_writable:
                type_name, category_name = get_element_info(element)
                skipped_elements.append({
                    'element': element,
                    'reason': write_reason,
                    'category': category_name,
                    'type_name': type_name
                })
                continue
            
            # Calculate weight
            weight, method = calculate_element_weight(element)
            
            if weight is not None:
                weight_formatted = round(weight, 1)
                results.append({
                    'element': element,
                    'weight': weight_formatted,
                    'method': method
                })
            else:
                type_name, category_name = get_element_info(element)
                failed_elements.append({
                    'element': element,
                    'reason': method,
                    'category': category_name,
                    'type_name': type_name
                })
    
    # Show results
    output_window = script.get_output()
    
    if failed_elements:
        print_failed_elements_table(failed_elements)
    
    if skipped_elements:
        print_skipped_elements_table(skipped_elements)
    
    if not results:
        if failed_elements or skipped_elements:
            output_window.print_md("---")
            output_window.print_md("## ⚠️ No elements were updated")
        return
    
    successful = 0
    write_failed = 0
    
    with revit.Transaction("Calculate Weights", doc):
        for result in results:
            element = result['element']
            weight = result['weight']
            
            if set_parameter_value(element, weight_parameter, float(weight)):
                successful += 1
            else:
                write_failed += 1
    
    # Print summary
    output_window.print_md("---")
    output_window.print_md("## ✅ Weight Calculation Complete")
    output_window.print_md("**Successfully updated:** {} elements".format(successful))
    
    if len(skipped_elements) > 0:
        output_window.print_md("**Skipped (not editable):** {} elements".format(len(skipped_elements)))
    if len(failed_elements) > 0:
        output_window.print_md("**Failed (no density):** {} elements".format(len(failed_elements)))
    if write_failed > 0:
        output_window.print_md("**Failed (write error):** {} elements".format(write_failed))
    
    output_window.print_md("**Parameter used:** {}".format(weight_parameter))


if __name__ == '__main__':
    main()
