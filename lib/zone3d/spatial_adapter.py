# -*- coding: utf-8 -*-
"""Spatial element adapter for abstracting differences between Areas and Rooms."""

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInParameter, BuiltInCategory,
    Category, ElementId, Level, SpatialElementBoundaryOptions, StorageType,
    FamilyInstance, FilledRegion, View, Options, GeometryElement
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
    
    def get_existing_instance(self, element, doc, zone_instances_cache=None):
        """Get existing 3D Zone instance for element if it exists.
        
        Args:
            element: Area or Room element
            doc: Revit document
            zone_instances_cache: Optional pre-filtered list of 3DZone instances
            
        Returns:
            FamilyInstance or None: Existing instance if found, None otherwise
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
        """Get area number using BuiltInParameter (same as rooms)."""
        # Areas use the same ROOM_NUMBER built-in parameter as Rooms
        param = element.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        if param and param.HasValue:
            return param.AsString()
        # Fallback to LookupParameter
        param_names_to_try = ["Number", "Area Number", "AREA_NUMBER", "AREA_NUM"]
        for param_name in param_names_to_try:
            param = element.LookupParameter(param_name)
            if param and param.HasValue:
                return param.AsString()
        return "?"
    
    def get_name(self, element):
        """Get area name using BuiltInParameter (same as rooms)."""
        # Areas use the same ROOM_NAME built-in parameter as Rooms
        param = element.get_Parameter(BuiltInParameter.ROOM_NAME)
        if param and param.HasValue:
            return param.AsString()
        # Fallback to LookupParameter
        param_names_to_try = ["Name", "Area Name", "AREA_NAME"]
        for param_name in param_names_to_try:
            param = element.LookupParameter(param_name)
            if param and param.HasValue:
                return param.AsString()
        return "Unnamed"
    
    def get_area_type(self, element, doc):
        """Get area type/scheme name for the Area.
        
        Note: In Revit, Areas have two related concepts:
        - Area Scheme (e.g., "Gross Building", "Rentable") - accessible via element.AreaScheme
        - Area Type (e.g., "Building Common Area", "Floor Area") - stored in AREA_TYPE parameter
          but the element it points to has no Name property (Revit API limitation)
        
        This method returns the Area Scheme name since Area Type elements have no accessible name.
        The AREA_TYPE_TEXT parameter shows the Area Type text but is read-only.
        
        Args:
            element: Area element
            doc: Revit document
            
        Returns:
            str: Area scheme name or "?" if not found
        """
        try:
            # Primary approach: use AreaScheme property directly
            # This gives us the Area Scheme name (Gross Building, Rentable, etc.)
            if hasattr(element, 'AreaScheme') and element.AreaScheme:
                scheme_name = element.AreaScheme.Name
                logger.debug("Area {}: AreaScheme.Name = '{}'".format(element.Id, scheme_name))
                return scheme_name
            else:
                logger.debug("Area {}: AreaScheme not available".format(element.Id))
            
            # Fallback: try AREA_TYPE_TEXT parameter (read-only but gives the display text)
            try:
                param = element.get_Parameter(BuiltInParameter.AREA_TYPE_TEXT)
                if param and param.HasValue:
                    value = param.AsString()
                    if value:
                        logger.debug("Area {}: AREA_TYPE_TEXT = '{}'".format(element.Id, value))
                        return value
            except Exception as e:
                logger.debug("Area {}: AREA_TYPE_TEXT error: {}".format(element.Id, e))
            
            # Fallback: try LookupParameter for "Area Type" as string
            param = element.LookupParameter("Area Type")
            if param and param.HasValue:
                if param.StorageType == StorageType.String:
                    value = param.AsString()
                    if value:
                        logger.debug("Area {}: LookupParameter('Area Type') = '{}'".format(element.Id, value))
                        return value
                        
        except Exception as e:
            logger.debug("Error getting area type for Area {}: {}".format(element.Id, e))
        
        logger.debug("Area {}: Could not get area type, returning '?'".format(element.Id))
        return "?"
    
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
        return self.get_existing_instance(element, doc, zone_instances_cache) is not None
    
    def get_existing_instance(self, element, doc, zone_instances_cache=None):
        """Get existing 3D Zone instance for area if it exists."""
        try:
            # Get area number
            area_number_str = self.get_number(element)
            if not area_number_str or area_number_str == "?":
                return None
            
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
                                        return instance
                    except Exception:
                        pass
            
            return None
        except Exception as e:
            logger.debug("Error getting existing 3D zone instance: {}".format(e))
            return None
    
    def get_family_name_prefix(self):
        """Return family name prefix for Areas."""
        return "3DZone_Area-"
    
    def sanitize_number(self, number):
        """No sanitization needed for Areas - return as-is."""
        return number
    
    def copy_properties_to_instance(self, source_element, target_instance, doc):
        """Copy all matching parameters from source element to target instance.
        
        Iterates through all parameters on the source element and copies values
        to matching parameters on the target instance if they exist and are writable.
        
        Args:
            source_element: Area element
            target_instance: FamilyInstance to copy properties to
            doc: Revit document
        """
        try:
            # Parameters to skip (system/built-in parameters that shouldn't be copied)
            # Build skip list defensively - only include constants that exist
            skip_params = set()
            skip_param_names = [
                'INVALID',
                'ELEM_TYPE_PARAM',
                'ELEM_CATEGORY_PARAM',
                'ELEM_FAMILY_PARAM',
                'ELEM_FAMILY_AND_TYPE_PARAM',
                'ELEM_TYPE_NAME_PARAM',
                'ELEM_TYPE_ID_PARAM',
                'ELEM_ID_PARAM',
                'ELEM_LEVEL_PARAM',
                'ELEM_LEVEL_ID_PARAM',
                'ELEM_PHASE_CREATED_PARAM',
                'ELEM_PHASE_DEMOLISHED_PARAM',
            ]
            # Add parameters only if they exist (not all constants available in all Revit versions)
            for param_name in skip_param_names:
                try:
                    if hasattr(BuiltInParameter, param_name):
                        skip_params.add(getattr(BuiltInParameter, param_name))
                except Exception:
                    # Skip this parameter if it causes an error
                    pass
            
            # Get all parameters from source element
            copied_count = 0
            skipped_count = 0
            
            for source_param in source_element.Parameters:
                try:
                    # Skip if no value
                    if not source_param or not source_param.HasValue:
                        continue
                    
                    # Get parameter definition
                    defn = source_param.Definition
                    if not defn:
                        continue
                    
                    # Skip if it's a built-in parameter in our skip list
                    if hasattr(defn, 'BuiltInParameter'):
                        try:
                            bip = defn.BuiltInParameter
                            # Check if INVALID exists before comparing
                            invalid_bip = getattr(BuiltInParameter, 'INVALID', None)
                            if invalid_bip is not None and bip != invalid_bip and bip in skip_params:
                                continue
                            elif invalid_bip is None and bip in skip_params:
                                # If INVALID doesn't exist, just check if bip is in skip_params
                                continue
                        except Exception:
                            # If accessing BuiltInParameter causes an error, skip this check
                            pass
                    
                    # Get parameter name
                    param_name = defn.Name
                    if not param_name:
                        continue
                    
                    # Skip if read-only on source
                    if source_param.IsReadOnly:
                        continue
                    
                    # Find matching parameter on target (try instance first, then type)
                    target_param = target_instance.LookupParameter(param_name)
                    if not target_param and hasattr(target_instance, 'Symbol'):
                        symbol = target_instance.Symbol
                        if symbol:
                            target_param = symbol.LookupParameter(param_name)
                    
                    if not target_param:
                        continue  # No matching parameter on target
                    
                    # Skip if target parameter is read-only
                    if target_param.IsReadOnly:
                        skipped_count += 1
                        continue
                    
                    # Check storage types match
                    if source_param.StorageType != target_param.StorageType:
                        skipped_count += 1
                        continue
                    
                    # Copy the value based on storage type
                    storage_type = source_param.StorageType
                    copy_success = False
                    
                    if storage_type == StorageType.String:
                        value = source_param.AsString()
                        if value:  # Only copy non-empty strings
                            # Skip if value hasn't changed
                            if target_param.HasValue:
                                current_value = target_param.AsString()
                                if current_value == value:
                                    continue
                            target_param.Set(value)
                            copy_success = True
                            
                    elif storage_type == StorageType.Integer:
                        value = source_param.AsInteger()
                        # Skip if value hasn't changed
                        if target_param.HasValue:
                            current_value = target_param.AsInteger()
                            if current_value == value:
                                continue
                        target_param.Set(value)
                        copy_success = True
                        
                    elif storage_type == StorageType.Double:
                        value = source_param.AsDouble()
                        # Skip if value hasn't changed (with tolerance)
                        if target_param.HasValue:
                            current_value = target_param.AsDouble()
                            if abs(current_value - value) < 1e-9:
                                continue
                        target_param.Set(value)
                        copy_success = True
                        
                    elif storage_type == StorageType.ElementId:
                        value = source_param.AsElementId()
                        if value and value != ElementId.InvalidElementId:
                            # Skip if value hasn't changed
                            if target_param.HasValue:
                                current_value = target_param.AsElementId()
                                if current_value == value:
                                    continue
                            target_param.Set(value)
                            copy_success = True
                    
                    if copy_success:
                        copied_count += 1
                        logger.debug("Copied parameter '{}' from Area {} to 3D Zone instance".format(
                            param_name, source_element.Id))
                        
                except Exception as param_error:
                    # Log but continue with next parameter
                    param_name_local = defn.Name if 'defn' in locals() and defn else 'unknown'
                    logger.debug("Error copying parameter '{}': {}".format(param_name_local, param_error))
                    continue
            
            logger.debug("Copied {} parameters from Area {} to 3D Zone instance ({} skipped)".format(
                copied_count, source_element.Id, skipped_count))
                
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
        return self.get_existing_instance(element, doc, zone_instances_cache) is not None
    
    def get_existing_instance(self, element, doc, zone_instances_cache=None):
        """Get existing 3D Zone instance for room if it exists."""
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
                    return instance
            except:
                pass
        
        return None
    
    def get_family_name_prefix(self):
        """Return family name prefix for Rooms."""
        return "3DZone_Room-"
    
    def sanitize_number(self, number):
        """Sanitize number using sanitize_family_name."""
        return self.sanitize_family_name(number)
    
    def copy_properties_to_instance(self, source_element, target_instance, doc):
        """Copy all matching parameters from source element to target instance.
        
        Iterates through all parameters on the source element and copies values
        to matching parameters on the target instance if they exist and are writable.
        
        Args:
            source_element: Room element
            target_instance: FamilyInstance to copy properties to
            doc: Revit document
        """
        try:
            # Parameters to skip (system/built-in parameters that shouldn't be copied)
            # Build skip list defensively - only include constants that exist
            skip_params = set()
            skip_param_names = [
                'INVALID',
                'ELEM_TYPE_PARAM',
                'ELEM_CATEGORY_PARAM',
                'ELEM_FAMILY_PARAM',
                'ELEM_FAMILY_AND_TYPE_PARAM',
                'ELEM_TYPE_NAME_PARAM',
                'ELEM_TYPE_ID_PARAM',
                'ELEM_ID_PARAM',
                'ELEM_LEVEL_PARAM',
                'ELEM_LEVEL_ID_PARAM',
                'ELEM_PHASE_CREATED_PARAM',
                'ELEM_PHASE_DEMOLISHED_PARAM',
                'ROOM_NUMBER',
                'ROOM_NAME',
            ]
            # Add parameters only if they exist (not all constants available in all Revit versions)
            for param_name in skip_param_names:
                try:
                    if hasattr(BuiltInParameter, param_name):
                        skip_params.add(getattr(BuiltInParameter, param_name))
                except Exception:
                    # Skip this parameter if it causes an error
                    pass
            
            # Get all parameters from source element
            copied_count = 0
            skipped_count = 0
            
            for source_param in source_element.Parameters:
                try:
                    # Skip if no value
                    if not source_param or not source_param.HasValue:
                        continue
                    
                    # Get parameter definition
                    defn = source_param.Definition
                    if not defn:
                        continue
                    
                    # Skip if it's a built-in parameter in our skip list
                    if hasattr(defn, 'BuiltInParameter'):
                        try:
                            bip = defn.BuiltInParameter
                            # Check if INVALID exists before comparing
                            invalid_bip = getattr(BuiltInParameter, 'INVALID', None)
                            if invalid_bip is not None and bip != invalid_bip and bip in skip_params:
                                continue
                            elif invalid_bip is None and bip in skip_params:
                                # If INVALID doesn't exist, just check if bip is in skip_params
                                continue
                        except Exception:
                            # If accessing BuiltInParameter causes an error, skip this check
                            pass
                    
                    # Get parameter name
                    param_name = defn.Name
                    if not param_name:
                        continue
                    
                    # Skip if read-only on source
                    if source_param.IsReadOnly:
                        continue
                    
                    # Find matching parameter on target (try instance first, then type)
                    target_param = target_instance.LookupParameter(param_name)
                    if not target_param and hasattr(target_instance, 'Symbol'):
                        symbol = target_instance.Symbol
                        if symbol:
                            target_param = symbol.LookupParameter(param_name)
                    
                    if not target_param:
                        continue  # No matching parameter on target
                    
                    # Skip if target parameter is read-only
                    if target_param.IsReadOnly:
                        skipped_count += 1
                        continue
                    
                    # Check storage types match
                    if source_param.StorageType != target_param.StorageType:
                        skipped_count += 1
                        continue
                    
                    # Copy the value based on storage type
                    storage_type = source_param.StorageType
                    copy_success = False
                    
                    if storage_type == StorageType.String:
                        value = source_param.AsString()
                        if value:  # Only copy non-empty strings
                            # Skip if value hasn't changed
                            if target_param.HasValue:
                                current_value = target_param.AsString()
                                if current_value == value:
                                    continue
                            target_param.Set(value)
                            copy_success = True
                            
                    elif storage_type == StorageType.Integer:
                        value = source_param.AsInteger()
                        # Skip if value hasn't changed
                        if target_param.HasValue:
                            current_value = target_param.AsInteger()
                            if current_value == value:
                                continue
                        target_param.Set(value)
                        copy_success = True
                        
                    elif storage_type == StorageType.Double:
                        value = source_param.AsDouble()
                        # Skip if value hasn't changed (with tolerance)
                        if target_param.HasValue:
                            current_value = target_param.AsDouble()
                            if abs(current_value - value) < 1e-9:
                                continue
                        target_param.Set(value)
                        copy_success = True
                        
                    elif storage_type == StorageType.ElementId:
                        value = source_param.AsElementId()
                        if value and value != ElementId.InvalidElementId:
                            # Skip if value hasn't changed
                            if target_param.HasValue:
                                current_value = target_param.AsElementId()
                                if current_value == value:
                                    continue
                            target_param.Set(value)
                            copy_success = True
                    
                    if copy_success:
                        copied_count += 1
                        logger.debug("Copied parameter '{}' from Room {} to 3D Zone instance".format(
                            param_name, source_element.Id))
                        
                except Exception as param_error:
                    # Log but continue with next parameter
                    param_name_local = defn.Name if 'defn' in locals() and defn else 'unknown'
                    logger.debug("Error copying parameter '{}': {}".format(param_name_local, param_error))
                    continue
            
            logger.debug("Copied {} parameters from Room {} to 3D Zone instance ({} skipped)".format(
                copied_count, source_element.Id, skipped_count))
                
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


class CurveSegmentWrapper(object):
    """Wrapper to make Curve objects work with boundary segment interface."""
    def __init__(self, curve):
        self._curve = curve
    
    def GetCurve(self):
        return self._curve


class RegionAdapter(SpatialElementAdapter):
    """Adapter for FilledRegion elements."""
    
    def __init__(self, active_view=None):
        """Initialize adapter with active view for phase handling.
        
        Args:
            active_view: View object (optional, will be retrieved from doc if needed)
        """
        self._active_view = active_view
        self._view_phase_id = None
    
    def set_active_view(self, view):
        """Set the active view for phase handling.
        
        Args:
            view: View object
        """
        self._active_view = view
        self._view_phase_id = None  # Reset cached phase
    
    def get_number(self, element):
        """Get region element ID as string."""
        return str(element.Id.IntegerValue)
    
    def get_name(self, element):
        """Get region type name or 'Unnamed Region'."""
        try:
            # Try to get FilledRegionType name
            if hasattr(element, 'GetTypeId'):
                type_id = element.GetTypeId()
                if type_id and type_id != ElementId.InvalidElementId:
                    region_type = element.Document.GetElement(type_id)
                    if region_type and hasattr(region_type, 'Name'):
                        return region_type.Name
        except Exception:
            pass
        return "Unnamed Region"
    
    def get_level_id(self, element):
        """Get level ID from the view the FilledRegion is in.
        
        FilledRegions are view-specific. If the view is a plan view,
        we can get its associated level. Otherwise, try to find the closest level.
        
        Args:
            element: FilledRegion element
            
        Returns:
            ElementId or None
        """
        try:
            # Get the view the region is in
            view_id = element.OwnerViewId
            if not view_id or view_id == ElementId.InvalidElementId:
                # Try to use active view if available
                if self._active_view:
                    view_id = self._active_view.Id
                else:
                    return None
            
            if not view_id or view_id == ElementId.InvalidElementId:
                return None
            
            doc = element.Document
            view = doc.GetElement(view_id)
            if not view:
                return None
            
            # Try to get level from view (works for plan views)
            try:
                if hasattr(view, 'GenLevel') and view.GenLevel:
                    level_id = view.GenLevel.Id
                    if level_id and level_id != ElementId.InvalidElementId:
                        logger.debug("Got level {} from view {}".format(level_id, view_id))
                        return level_id
            except Exception:
                pass
            
            # If no level from view, try to find closest level to region's Z
            try:
                bbox = element.get_BoundingBox(view)
                if bbox:
                    # Get center Z coordinate
                    center_z = (bbox.Min.Z + bbox.Max.Z) / 2.0
                    
                    # Find closest level
                    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
                    if levels:
                        closest_level = None
                        min_distance = float('inf')
                        for level in levels:
                            distance = abs(level.Elevation - center_z)
                            if distance < min_distance:
                                min_distance = distance
                                closest_level = level
                        
                        if closest_level:
                            logger.debug("Found closest level {} (elevation {:.2f}) for region Z {:.2f}".format(
                                closest_level.Id, closest_level.Elevation, center_z))
                            return closest_level.Id
            except Exception as e:
                logger.debug("Error finding closest level: {}".format(e))
            
            return None
        except Exception as e:
            logger.debug("Error getting level for FilledRegion: {}".format(e))
            return None
    
    def calculate_height(self, element, doc, levels_cache=None):
        """Return default height (10.0 feet) for regions."""
        return 10.0
    
    def get_boundary_segments(self, element):
        """Extract boundary curves from FilledRegion.
        
        FilledRegion.GetBoundaries() returns IList[CurveLoop].
        We need to convert this to the format expected by extract_boundary_loops:
        List of boundary segment groups, where each group is a list of segments
        with GetCurve() method.
        
        Args:
            element: FilledRegion element
            
        Returns:
            List of boundary segment groups (list of lists of CurveSegmentWrapper)
        """
        try:
            # Get boundaries from FilledRegion
            boundaries = element.GetBoundaries()
            
            if not boundaries or len(boundaries) == 0:
                logger.debug("FilledRegion {} has no boundaries".format(element.Id))
                return []
            
            # Convert CurveLoop objects to boundary segment format
            boundary_segments = []
            for curve_loop in boundaries:
                # Create a group (list) of segments for this loop
                segment_group = []
                for curve in curve_loop:
                    # Wrap each curve in a wrapper that provides GetCurve() method
                    segment_group.append(CurveSegmentWrapper(curve))
                
                if segment_group:
                    boundary_segments.append(segment_group)
            
            logger.debug("FilledRegion {}: Found {} boundary loops".format(element.Id, len(boundary_segments)))
            return boundary_segments
            
        except Exception as e:
            logger.error("Error getting boundaries from FilledRegion {}: {}".format(element.Id, e))
            import traceback
            logger.error(traceback.format_exc())
            return []
    
    def check_existing_zone(self, element, doc, zone_instances_cache=None):
        """Check if region has existing zone using pattern matching."""
        return self.get_existing_instance(element, doc, zone_instances_cache) is not None
    
    def get_existing_instance(self, element, doc, zone_instances_cache=None):
        """Get existing 3D Zone instance for region if it exists."""
        try:
            region_id_str = str(element.Id.IntegerValue)
            
            # Find all Generic Model instances if cache not provided
            if zone_instances_cache is None:
                generic_instances = FilteredElementCollector(doc)\
                    .OfClass(FamilyInstance)\
                    .OfCategory(BuiltInCategory.OST_GenericModel)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
            else:
                generic_instances = zone_instances_cache
            
            # Check if any instance matches this region by family name pattern
            expected_family_name_pattern = "3DZone_Region_{}_".format(region_id_str)
            
            for instance in generic_instances:
                if instance.Category and instance.Category.Id == Category.GetCategory(doc, BuiltInCategory.OST_GenericModel).Id:
                    try:
                        symbol = instance.Symbol
                        if symbol:
                            family = symbol.Family
                            if family:
                                family_name = family.Name
                                # Check if family name starts with the expected pattern
                                if family_name.startswith(expected_family_name_pattern) or region_id_str in family_name:
                                    return instance
                    except Exception:
                        pass
            
            return None
        except Exception as e:
            logger.debug("Error getting existing 3D zone instance: {}".format(e))
            return None
    
    def get_family_name_prefix(self):
        """Return family name prefix for Regions."""
        return "3DZone_Region_"
    
    def sanitize_number(self, number):
        """Return element ID string as-is (already sanitized)."""
        return number
    
    def copy_properties_to_instance(self, source_element, target_instance, doc):
        """Copy matching parameters from FilledRegion to target instance.
        
        Args:
            source_element: FilledRegion element
            target_instance: FamilyInstance to copy properties to
            doc: Revit document
        """
        try:
            # Parameters to skip (system/built-in parameters)
            skip_params = set()
            skip_param_names = [
                'INVALID',
                'ELEM_TYPE_PARAM',
                'ELEM_CATEGORY_PARAM',
                'ELEM_FAMILY_PARAM',
                'ELEM_FAMILY_AND_TYPE_PARAM',
                'ELEM_TYPE_NAME_PARAM',
                'ELEM_TYPE_ID_PARAM',
                'ELEM_ID_PARAM',
                'ELEM_LEVEL_PARAM',
                'ELEM_LEVEL_ID_PARAM',
                'ELEM_PHASE_CREATED_PARAM',
                'ELEM_PHASE_DEMOLISHED_PARAM',
            ]
            for param_name in skip_param_names:
                try:
                    if hasattr(BuiltInParameter, param_name):
                        skip_params.add(getattr(BuiltInParameter, param_name))
                except Exception:
                    pass
            
            copied_count = 0
            skipped_count = 0
            
            for source_param in source_element.Parameters:
                try:
                    if not source_param or not source_param.HasValue:
                        continue
                    
                    defn = source_param.Definition
                    if not defn:
                        continue
                    
                    # Skip if it's a built-in parameter in our skip list
                    if hasattr(defn, 'BuiltInParameter'):
                        try:
                            bip = defn.BuiltInParameter
                            invalid_bip = getattr(BuiltInParameter, 'INVALID', None)
                            if invalid_bip is not None and bip != invalid_bip and bip in skip_params:
                                continue
                            elif invalid_bip is None and bip in skip_params:
                                continue
                        except Exception:
                            pass
                    
                    param_name = defn.Name
                    if not param_name:
                        continue
                    
                    if source_param.IsReadOnly:
                        continue
                    
                    # Find matching parameter on target
                    target_param = target_instance.LookupParameter(param_name)
                    if not target_param and hasattr(target_instance, 'Symbol'):
                        symbol = target_instance.Symbol
                        if symbol:
                            target_param = symbol.LookupParameter(param_name)
                    
                    if not target_param:
                        continue
                    
                    if target_param.IsReadOnly:
                        skipped_count += 1
                        continue
                    
                    if source_param.StorageType != target_param.StorageType:
                        skipped_count += 1
                        continue
                    
                    storage_type = source_param.StorageType
                    copy_success = False
                    
                    if storage_type == StorageType.String:
                        value = source_param.AsString()
                        if value:
                            if target_param.HasValue:
                                current_value = target_param.AsString()
                                if current_value == value:
                                    continue
                            target_param.Set(value)
                            copy_success = True
                    elif storage_type == StorageType.Integer:
                        value = source_param.AsInteger()
                        if target_param.HasValue:
                            current_value = target_param.AsInteger()
                            if current_value == value:
                                continue
                        target_param.Set(value)
                        copy_success = True
                    elif storage_type == StorageType.Double:
                        value = source_param.AsDouble()
                        if target_param.HasValue:
                            current_value = target_param.AsDouble()
                            if abs(current_value - value) < 1e-9:
                                continue
                        target_param.Set(value)
                        copy_success = True
                    elif storage_type == StorageType.ElementId:
                        value = source_param.AsElementId()
                        if value and value != ElementId.InvalidElementId:
                            if target_param.HasValue:
                                current_value = target_param.AsElementId()
                                if current_value == value:
                                    continue
                            target_param.Set(value)
                            copy_success = True
                    
                    if copy_success:
                        copied_count += 1
                        logger.debug("Copied parameter '{}' from FilledRegion {} to 3D Zone instance".format(
                            param_name, source_element.Id))
                        
                except Exception as param_error:
                    param_name_local = defn.Name if 'defn' in locals() and defn else 'unknown'
                    logger.debug("Error copying parameter '{}': {}".format(param_name_local, param_error))
                    continue
            
            logger.debug("Copied {} parameters from FilledRegion {} to 3D Zone instance ({} skipped)".format(
                copied_count, source_element.Id, skipped_count))
                
        except Exception as prop_error:
            logger.debug("Error copying properties: {}".format(prop_error))
    
    def get_phase_id(self, element):
        """Get phase ID from the view the FilledRegion is in.
        
        Args:
            element: FilledRegion element
            
        Returns:
            ElementId or None
        """
        try:
            # Get the view the region is in
            view_id = element.OwnerViewId
            if not view_id or view_id == ElementId.InvalidElementId:
                # Fallback to active view if available
                if self._active_view:
                    view_id = self._active_view.Id
                else:
                    return None
            
            if not view_id or view_id == ElementId.InvalidElementId:
                return None
            
            doc = element.Document
            view = doc.GetElement(view_id)
            if not view:
                # Fallback to active view
                if self._active_view:
                    view = self._active_view
                else:
                    return None
            
            # Get phase from view
            try:
                phase_param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
                if phase_param and phase_param.HasValue:
                    phase_id = phase_param.AsElementId()
                    if phase_id and phase_id != ElementId.InvalidElementId:
                        logger.debug("Got phase {} from view {}".format(phase_id, view_id))
                        return phase_id
            except Exception as e:
                logger.debug("Error getting phase from view: {}".format(e))
            
            return None
        except Exception as e:
            logger.debug("Error getting phase for FilledRegion: {}".format(e))
            return None
    
    def set_phase_on_instance(self, instance, phase_id):
        """Set phase on instance.
        
        Args:
            instance: FamilyInstance
            phase_id: ElementId or None
        """
        if instance and phase_id and phase_id != ElementId.InvalidElementId:
            try:
                instance.CreatedPhaseId = phase_id
                logger.debug("Set phase {} on instance {}".format(phase_id, instance.Id))
            except Exception as e:
                logger.debug("Error setting phase on instance: {}".format(e))

