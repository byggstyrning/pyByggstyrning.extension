# -*- coding: utf-8 -*-
"""Spatial element adapter for abstracting differences between Areas and Rooms."""

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInParameter, BuiltInCategory,
    Category, ElementId, Level, SpatialElementBoundaryOptions, StorageType,
    FamilyInstance
)
from pyrevit import script

logger = script.get_logger()


class SpatialElementAdapter(object):
    """Abstract adapter for spatial elements (Area/Room)."""
    
    def get_number(self, element):
        """Get element number string.
        
        Args:
            element: Area or Room element
            
        Returns:
            str: Element number or "?" if not found
        """
        raise NotImplementedError
    
    def get_name(self, element):
        """Get element name string.
        
        Args:
            element: Area or Room element
            
        Returns:
            str: Element name or "Unnamed" if not found
        """
        raise NotImplementedError
    
    def get_level_id(self, element):
        """Get level ID.
        
        Args:
            element: Area or Room element
            
        Returns:
            ElementId or None
        """
        if hasattr(element, 'LevelId') and element.LevelId:
            return element.LevelId
        return None
    
    def calculate_height(self, element, doc, levels_cache=None):
        """Calculate height. Returns float or None (use family default).
        
        Args:
            element: Area or Room element
            doc: Revit document
            levels_cache: Optional pre-collected and sorted list of Level elements
            
        Returns:
            float or None: Height in feet, or None to use family default
        """
        raise NotImplementedError
    
    def get_boundary_segments(self, element):
        """Get boundary segments.
        
        Args:
            element: Area or Room element
            
        Returns:
            List of boundary segment groups
        """
        boundary_options = SpatialElementBoundaryOptions()
        return element.GetBoundarySegments(boundary_options)
    
    def check_existing_zone(self, element, doc, zone_instances_cache=None):
        """Check if element already has a 3D Zone.
        
        Args:
            element: Area or Room element
            doc: Revit document
            zone_instances_cache: Optional pre-filtered list of 3DZone instances
            
        Returns:
            bool: True if element has a 3D Zone instance
        """
        raise NotImplementedError
    
    def get_family_name_prefix(self):
        """Return family name prefix: '3DZone_Area-' or '3DZone_Room-'.
        
        Returns:
            str: Family name prefix
        """
        raise NotImplementedError
    
    def sanitize_number(self, number):
        """Sanitize number for family naming. Returns sanitized string.
        
        Args:
            number: Number string to sanitize
            
        Returns:
            str: Sanitized number safe for family naming
        """
        raise NotImplementedError
    
    def copy_properties_to_instance(self, source_element, target_instance, doc):
        """Copy Name/Number/MMI from source to target.
        
        Args:
            source_element: Area or Room element
            target_instance: FamilyInstance to copy properties to
            doc: Revit document
        """
        raise NotImplementedError
    
    def get_phase_id(self, element):
        """Get phase ID if applicable. Returns ElementId or None.
        
        Args:
            element: Area or Room element
            
        Returns:
            ElementId or None
        """
        return None
    
    def set_phase_on_instance(self, instance, phase_id):
        """Set phase on instance if applicable.
        
        Args:
            instance: FamilyInstance
            phase_id: ElementId or None
        """
        pass  # Default: no-op
    
    def set_symbol_parameters_before_placement(self, symbol, height, template_info, doc, element_number_str):
        """Set parameters on symbol before instance creation (optional optimization).
        
        Args:
            symbol: FamilySymbol
            height: Height value
            template_info: Template info dict
            doc: Revit document
            element_number_str: Element number string for logging
        """
        pass  # Default: no-op


class AreaAdapter(SpatialElementAdapter):
    """Adapter for Area elements."""
    
    def get_number(self, element):
        """Get area number using LookupParameter."""
        param_names_to_try = ["Number", "Area Number", "AREA_NUMBER", "AREA_NUM"]
        for param_name in param_names_to_try:
            param = element.LookupParameter(param_name)
            if param and param.HasValue:
                return param.AsString()
        return "?"
    
    def get_name(self, element):
        """Get area name using LookupParameter."""
        param_names_to_try = ["Name", "Area Name", "AREA_NAME"]
        for param_name in param_names_to_try:
            param = element.LookupParameter(param_name)
            if param and param.HasValue:
                return param.AsString()
        return "Unnamed"
    
    def calculate_height(self, element, doc, levels_cache=None):
        """Calculate height using 'Storey Above' parameter logic."""
        level_id = self.get_level_id(element)
        if not level_id:
            logger.debug("Area {} has no level, using family default height".format(element.Id))
            return None  # Use family default
        
        level = doc.GetElement(level_id)
        if not level:
            return None
        
        try:
            # Get "Storey Above" parameter from level
            storey_above_param = None
            try:
                storey_above_param = level.get_Parameter(BuiltInParameter.LEVEL_STOREY_ABOVE)
            except:
                # Fallback: try LookupParameter
                try:
                    storey_above_param = level.LookupParameter("Storey Above")
                except:
                    pass
            
            if storey_above_param and storey_above_param.HasValue:
                storey_above_id = storey_above_param.AsElementId()
                
                # Check if it's set to "default" (InvalidElementId or None)
                if storey_above_id and storey_above_id != ElementId.InvalidElementId:
                    # Get the level above
                    level_above = doc.GetElement(storey_above_id)
                    if level_above:
                        base_z = level.Elevation
                        top_z = level_above.Elevation
                        return top_z - base_z
                    else:
                        logger.debug("Area {}: Storey Above level not found, using family default".format(element.Id))
                else:
                    # Storey Above is set to "default" - use family default height
                    logger.debug("Area {}: Storey Above is set to 'default', using family default height".format(element.Id))
                    return None
            
            # Storey Above parameter not found or not set - fallback to next level
            logger.debug("Area {}: Storey Above parameter not found, falling back to next level".format(element.Id))
            base_z = level.Elevation
            
            # Use cached levels if provided, otherwise collect them
            if levels_cache is None:
                all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                levels_sorted = sorted(all_levels, key=lambda l: l.Elevation)
            else:
                levels_sorted = levels_cache
            
            top_z = base_z + 10.0  # Default 10 feet
            for lvl in levels_sorted:
                if lvl.Elevation > base_z:
                    top_z = lvl.Elevation
                    break
            return top_z - base_z
            
        except Exception as storey_error:
            logger.debug("Error getting Storey Above parameter: {}, falling back to next level".format(storey_error))
            # Fallback to next level calculation
            base_z = level.Elevation
            if levels_cache is None:
                all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                levels_sorted = sorted(all_levels, key=lambda l: l.Elevation)
            else:
                levels_sorted = levels_cache
            top_z = base_z + 10.0
            for lvl in levels_sorted:
                if lvl.Elevation > base_z:
                    top_z = lvl.Elevation
                    break
            return top_z - base_z
    
    def check_existing_zone(self, element, doc, zone_instances_cache=None):
        """Check if area has existing zone using pattern matching."""
        try:
            # Get area number
            area_number_str = self.get_number(element)
            if not area_number_str or area_number_str == "?":
                return False
            
            # Find all Generic Model instances if cache not provided
            if zone_instances_cache is None:
                generic_instances = FilteredElementCollector(doc)\
                    .OfClass(FamilyInstance)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
            else:
                generic_instances = zone_instances_cache
            
            # Check if any instance matches this area by family name pattern
            area_id_str = str(element.Id.IntegerValue)
            expected_family_name_pattern = "3DZone_Area-{}_".format(area_number_str.replace(" ", "-"))
            
            for instance in generic_instances:
                if instance.Category and instance.Category.Id == Category.GetCategory(doc, BuiltInCategory.OST_GenericModel).Id:
                    try:
                        symbol = instance.Symbol
                        if symbol:
                            family = symbol.Family
                            if family:
                                family_name = family.Name
                                # Check if family name starts with the expected pattern
                                if family_name.startswith(expected_family_name_pattern):
                                    # Also verify it contains the area ID to be more precise
                                    if area_id_str in family_name:
                                        return True
                    except Exception:
                        pass
            
            return False
        except Exception as e:
            logger.debug("Error checking for existing 3D zone: {}".format(e))
            return False
    
    def get_family_name_prefix(self):
        """Return family name prefix for Areas."""
        return "3DZone_Area-"
    
    def sanitize_number(self, number):
        """No sanitization needed for Areas - return as-is."""
        return number
    
    def copy_properties_to_instance(self, source_element, target_instance, doc):
        """Copy properties using LookupParameter."""
        try:
            # Import MMI utilities
            try:
                from mmi.core import get_mmi_parameter_name
            except ImportError:
                def get_mmi_parameter_name(doc):
                    return "MMI"
            
            # Copy Name
            area_name = source_element.LookupParameter("Name")
            if not area_name:
                area_name = source_element.LookupParameter("Area Name")
            if area_name and area_name.HasValue:
                inst_name = target_instance.LookupParameter("Name")
                if inst_name and not inst_name.IsReadOnly:
                    inst_name.Set(area_name.AsString())
            
            # Copy Number
            area_number = source_element.LookupParameter("Number")
            if not area_number:
                area_number = source_element.LookupParameter("Area Number")
            if area_number and area_number.HasValue:
                inst_number = target_instance.LookupParameter("Number")
                if inst_number and not inst_number.IsReadOnly:
                    inst_number.Set(area_number.AsString())
            
            # Copy MMI
            configured_mmi_param = get_mmi_parameter_name(doc)
            mmi_copied = False
            if configured_mmi_param:
                source_mmi = source_element.LookupParameter(configured_mmi_param)
                if source_mmi and source_mmi.HasValue:
                    target_mmi = target_instance.LookupParameter(configured_mmi_param)
                    if target_mmi and not target_mmi.IsReadOnly:
                        if source_mmi.StorageType == StorageType.String:
                            value = source_mmi.AsString()
                            if value:
                                target_mmi.Set(value)
                                mmi_copied = True
            
            if not mmi_copied:
                source_mmi = source_element.LookupParameter("MMI")
                if source_mmi and source_mmi.HasValue:
                    target_mmi = target_instance.LookupParameter("MMI")
                    if target_mmi and not target_mmi.IsReadOnly:
                        if source_mmi.StorageType == StorageType.String:
                            value = source_mmi.AsString()
                            if value:
                                target_mmi.Set(value)
        except Exception as prop_error:
            logger.debug("Error copying properties: {}".format(prop_error))


class RoomAdapter(SpatialElementAdapter):
    """Adapter for Room elements."""
    
    def sanitize_family_name(self, name):
        """Sanitize a string to be safe for use as a Revit family name."""
        if not name:
            return "Unknown"
        
        # Replace problematic characters
        replacements = {
            '.': '_',
            '/': '-',
            '\\': '-',
            '*': '',
            '?': '',
            '"': '',
            '<': '',
            '>': '',
            '|': '',
            ':': '-',
            ' ': '-',
        }
        
        result = name
        for char, replacement in replacements.items():
            result = result.replace(char, replacement)
        
        # Remove any double dashes or underscores
        while '--' in result:
            result = result.replace('--', '-')
        while '__' in result:
            result = result.replace('__', '_')
        
        # Strip leading/trailing dashes and underscores
        result = result.strip('-_')
        
        # Ensure we have something left
        if not result:
            return "Unknown"
        
        return result
    
    def get_number(self, element):
        """Get room number using BuiltInParameter."""
        param = element.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        if param and param.HasValue:
            return param.AsString()
        return "?"
    
    def get_name(self, element):
        """Get room name using BuiltInParameter."""
        param = element.get_Parameter(BuiltInParameter.ROOM_NAME)
        if param and param.HasValue:
            return param.AsString()
        return "Unnamed"
    
    def calculate_height(self, element, doc, levels_cache=None):
        """Calculate height using UnboundedHeight or level-to-level."""
        try:
            # Try UnboundedHeight first
            if hasattr(element, 'UnboundedHeight') and element.UnboundedHeight > 0:
                return element.UnboundedHeight
            
            # Fallback: level to level
            level_id = self.get_level_id(element)
            if not level_id:
                return 10.0  # Default fallback
            
            level = doc.GetElement(level_id)
            if not level:
                return 10.0  # Default fallback
            
            base_z = level.Elevation
            
            # Use cached levels if provided, otherwise collect them
            if levels_cache is None:
                all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                levels_sorted = sorted(all_levels, key=lambda l: l.Elevation)
            else:
                levels_sorted = levels_cache
            
            top_z = base_z + 10.0  # Default 10 feet
            for lvl in levels_sorted:
                if lvl.Elevation > base_z:
                    top_z = lvl.Elevation
                    break
            return top_z - base_z
            
        except Exception as e:
            logger.debug("Error calculating height: {}".format(e))
            return 10.0  # Default fallback
    
    def check_existing_zone(self, element, doc, zone_instances_cache=None):
        """Check if room has existing zone using cache-optimized lookup."""
        room_id_str = str(element.Id.IntegerValue)
        
        # Iterate through pre-filtered 3DZone instances only
        if zone_instances_cache is None:
            # Fallback: collect all Generic Model instances
            all_generic_instances = FilteredElementCollector(doc)\
                .OfClass(FamilyInstance)\
                .OfCategory(BuiltInCategory.OST_GenericModel)\
                .WhereElementIsNotElementType()\
                .ToElements()
            
            zone_instances_cache = []
            for instance in all_generic_instances:
                try:
                    family_name = instance.Symbol.Family.Name
                    if family_name.startswith("3DZone_Room-"):
                        zone_instances_cache.append(instance)
                except:
                    pass
        
        for instance in zone_instances_cache:
            try:
                family_name = instance.Symbol.Family.Name
                # Check if this 3DZone family matches this room (room ID is unique)
                if room_id_str in family_name:
                    return True
            except:
                pass
        
        return False
    
    def get_family_name_prefix(self):
        """Return family name prefix for Rooms."""
        return "3DZone_Room-"
    
    def sanitize_number(self, number):
        """Sanitize number using sanitize_family_name."""
        return self.sanitize_family_name(number)
    
    def copy_properties_to_instance(self, source_element, target_instance, doc):
        """Copy properties using BuiltInParameter."""
        try:
            # Import MMI utilities
            try:
                from mmi.core import get_mmi_parameter_name
            except ImportError:
                def get_mmi_parameter_name(doc):
                    return "MMI"
            
            # Copy Name
            room_name = source_element.get_Parameter(BuiltInParameter.ROOM_NAME)
            if room_name and room_name.HasValue:
                inst_name = target_instance.LookupParameter("Name")
                if inst_name and not inst_name.IsReadOnly:
                    inst_name.Set(room_name.AsString())
            
            # Copy Number
            room_number = source_element.get_Parameter(BuiltInParameter.ROOM_NUMBER)
            if room_number and room_number.HasValue:
                inst_number = target_instance.LookupParameter("Number")
                if inst_number and not inst_number.IsReadOnly:
                    inst_number.Set(room_number.AsString())
            
            # Copy MMI
            configured_mmi_param = get_mmi_parameter_name(doc)
            mmi_copied = False
            if configured_mmi_param:
                source_mmi = source_element.LookupParameter(configured_mmi_param)
                if source_mmi and source_mmi.HasValue:
                    target_mmi = target_instance.LookupParameter(configured_mmi_param)
                    if target_mmi and not target_mmi.IsReadOnly:
                        if source_mmi.StorageType == StorageType.String:
                            value = source_mmi.AsString()
                            if value:
                                target_mmi.Set(value)
                                mmi_copied = True
            
            if not mmi_copied:
                source_mmi = source_element.LookupParameter("MMI")
                if source_mmi and source_mmi.HasValue:
                    target_mmi = target_instance.LookupParameter("MMI")
                    if target_mmi and not target_mmi.IsReadOnly:
                        if source_mmi.StorageType == StorageType.String:
                            value = source_mmi.AsString()
                            if value:
                                target_mmi.Set(value)
        except Exception as prop_error:
            logger.debug("Error copying properties: {}".format(prop_error))
    
    def get_phase_id(self, element):
        """Get room phase ID."""
        try:
            room_phase_param = element.get_Parameter(BuiltInParameter.ROOM_PHASE)
            if room_phase_param and room_phase_param.HasValue:
                return room_phase_param.AsElementId()
        except Exception:
            pass  # Room may not have phase
        return None
    
    def set_phase_on_instance(self, instance, phase_id):
        """Set phase on instance."""
        if instance and phase_id and phase_id != ElementId.InvalidElementId:
            try:
                instance.CreatedPhaseId = phase_id
            except Exception:
                pass  # Instance may not support phase setting
    
    def set_symbol_parameters_before_placement(self, symbol, height, template_info, doc, element_number_str):
        """Set Top Offset parameter on symbol before instance creation (Rooms-specific optimization).
        
        Args:
            symbol: FamilySymbol
            height: Height value
            template_info: Template info dict
            doc: Revit document
            element_number_str: Element number string for logging
        """
        if height and template_info and template_info.get('top_offset_param'):
            try:
                top_offset_name = template_info['top_offset_param'].Definition.Name
                # Try various name variations
                for param_name in [top_offset_name, "Top Offset", "TopOffset", "Top Offset (default)"]:
                    type_param = symbol.LookupParameter(param_name)
                    if type_param and not type_param.IsReadOnly:
                        type_param.Set(height)
                        logger.debug("Set top offset parameter {} to {} on loaded symbol".format(param_name, height))
                        doc.Regenerate()
                        return
                
                # Fallback: if Top Offset not found, try ExtrusionEnd
                if template_info.get('extrusion_end_param'):
                    end_param_name = template_info['extrusion_end_param'].Definition.Name
                    for param_name in [end_param_name, "ExtrusionEnd", "NVExtrusionEnd"]:
                        type_param = symbol.LookupParameter(param_name)
                        if type_param and not type_param.IsReadOnly:
                            type_param.Set(height)
                            logger.debug("Fallback: Set type parameter {} to {} on loaded symbol".format(param_name, height))
                            doc.Regenerate()
                            return
            except Exception as symbol_set_error:
                logger.debug("Could not set type parameter on symbol: {}".format(symbol_set_error))

