# -*- coding: utf-8 -*-
"""Spatial containment detection for 3D Zone parameter mapping."""

from collections import defaultdict
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, XYZ, 
    SpatialElementBoundaryOptions, SpatialElementBoundaryLocation,
    BoundingBoxIntersectsFilter, BoundingBoxContainsPointFilter, 
    Outline, ElementIntersectsSolidFilter, Options, SpatialElement, 
    Line, SolidCurveIntersectionOptions, SolidCurveIntersectionMode, 
    Area, Level, CurveLoop, Solid, GeometryCreationUtilities, Transform
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.DB.Mechanical import Space
from pyrevit import script

# Initialize logger
logger = script.get_logger()

# Cache for geometry calculations (Mass/Generic Model)
# Structure: element_id -> {"solid": Solid, "bbox": BoundingBox}
_geometry_cache = {}

# Shared Options instance for geometry calculations (reused for performance)
_geometry_options = None

def _get_geometry_options(doc):
    """Get or create shared geometry options instance."""
    global _geometry_options
    if _geometry_options is None:
        _geometry_options = Options()
        _geometry_options.ComputeReferences = False
        if doc and doc.ActiveView:
            _geometry_options.DetailLevel = doc.ActiveView.DetailLevel
    return _geometry_options

def detect_containment_strategy(source_categories):
    """Detect optimal containment strategy based on source categories.
    
    Args:
        source_categories: List of BuiltInCategory values
        
    Returns:
        str: Strategy name ("room", "space", "area", "element", or None)
    """
    if not source_categories:
        return None
    
    # Check for Rooms (fastest)
    if BuiltInCategory.OST_Rooms in source_categories:
        return "room"
    
    # Check for Spaces (fastest)
    if BuiltInCategory.OST_MEPSpaces in source_categories:
        return "space"
    
    # Check for Areas (slower)
    if BuiltInCategory.OST_Areas in source_categories:
        return "area"
    
    # Check for Mass/Generic Model (slowest)
    if BuiltInCategory.OST_Mass in source_categories or \
       BuiltInCategory.OST_GenericModel in source_categories:
        return "element"
    
    return None

def get_element_representative_point(element):
    """Get a representative point from an element for containment testing.
    
    Areas are 2D planar elements and are only used as sources (not targets),
    but we handle them here for robustness.
    
    Args:
        element: Revit element
        
    Returns:
        XYZ: Point location or None if cannot be determined
    """
    try:
        # Special handling for Areas - use bounding box center
        # Areas are 2D planar and don't have reliable LocationCurve
        if isinstance(element, Area):
            bbox = element.get_BoundingBox(None)
            if bbox:
                return (bbox.Min + bbox.Max) / 2.0
            return None
        
        loc = element.Location
        if loc is None:
            # Try bounding box center
            bbox = element.get_BoundingBox(None)
            if bbox:
                return (bbox.Min + bbox.Max) / 2.0
            return None
        
        if hasattr(loc, "Point"):
            # LocationPoint
            return loc.Point
        elif hasattr(loc, "Curve"):
            # LocationCurve - get midpoint
            # Use normalized=False for unbound curves (like Area boundaries)
            curve = loc.Curve
            if curve:
                try:
                    # Try normalized first (for bound curves)
                    return curve.Evaluate(0.5, True)
                except:
                    # Fall back to non-normalized for unbound curves
                    try:
                        start = curve.GetEndPoint(0)
                        end = curve.GetEndPoint(1)
                        return (start + end) / 2.0
                    except:
                        pass
        
        # Fallback to bounding box
        bbox = element.get_BoundingBox(None)
        if bbox:
            return (bbox.Min + bbox.Max) / 2.0
        
        return None
    except Exception as e:
        logger.info("Error getting element point: {}".format(str(e)))
        return None

def get_element_test_points(element):
    """Get multiple test points for an element to improve containment detection.
    
    For walls and linear elements, returns multiple points along the element.
    For point-based elements, returns a single point.
    
    Areas are 2D planar elements and are only used as sources (not targets),
    but we handle them here for robustness.
    
    Args:
        element: Revit element
        
    Returns:
        list: List of XYZ points to test
    """
    points = []
    try:
        # Special handling for Areas - use bounding box center
        if isinstance(element, Area):
            bbox = element.get_BoundingBox(None)
            if bbox:
                points.append((bbox.Min + bbox.Max) / 2.0)
            return points
        
        loc = element.Location
        if loc is None:
            bbox = element.get_BoundingBox(None)
            if bbox:
                points.append((bbox.Min + bbox.Max) / 2.0)
            return points
        
        if hasattr(loc, "Point"):
            # LocationPoint - single point
            points.append(loc.Point)
        elif hasattr(loc, "Curve"):
            # LocationCurve - get multiple points along the curve
            curve = loc.Curve
            if curve:
                # Get points at 0.25, 0.5, and 0.75 along the curve
                for t in [0.25, 0.5, 0.75]:
                    try:
                        # Try normalized first
                        pt = curve.Evaluate(t, True)
                        points.append(pt)
                    except:
                        # Fall back to non-normalized calculation
                        try:
                            start = curve.GetEndPoint(0)
                            end = curve.GetEndPoint(1)
                            pt = start + (end - start) * t
                            points.append(pt)
                        except:
                            pass
                
                # Also try midpoint with slight offset perpendicular to wall
                try:
                    start_pt = curve.GetEndPoint(0)
                    end_pt = curve.GetEndPoint(1)
                    midpoint = (start_pt + end_pt) / 2.0
                    direction = (end_pt - start_pt).Normalize()
                    # Get perpendicular direction (rotate 90 degrees in XY plane)
                    perp_direction = XYZ(-direction.Y, direction.X, 0).Normalize()
                    # Try both directions with small offset (0.1 feet)
                    points.append(midpoint + perp_direction * 0.1)
                    points.append(midpoint - perp_direction * 0.1)
                except:
                    pass
        
        # Fallback to bounding box center
        if not points:
            bbox = element.get_BoundingBox(None)
            if bbox:
                points.append((bbox.Min + bbox.Max) / 2.0)
        
        return points
    except Exception as e:
        logger.info("Error getting element test points: {}".format(str(e)))
        return points

def is_point_in_room(room, point):
    """Check if a point is inside a room (fastest method).
    
    Args:
        room: Room element
        point: XYZ point
        
    Returns:
        bool: True if point is in room
    """
    try:
        if not isinstance(room, Room):
            return False
        return room.IsPointInRoom(point)
    except Exception as e:
        logger.info("Error checking point in room: {}".format(str(e)))
        return False

def is_point_in_space(space, point):
    """Check if a point is inside a space (fastest method).
    
    Args:
        space: Space element
        point: XYZ point
        
    Returns:
        bool: True if point is in space
    """
    try:
        if not isinstance(space, Space):
            return False
        return space.IsPointInSpace(point)
    except Exception as e:
        logger.info("Error checking point in space: {}".format(str(e)))
        return False

def _create_solid_from_area(area, doc):
    """Create a 3D solid from an Area by extruding its boundary between levels.
    
    Areas are 2D planar elements. This function creates a 3D solid by:
    1. Getting the Area's boundary segments
    2. Creating a CurveLoop from the segments
    3. Getting the Area's level
    4. Finding the level above (or using default height)
    5. Extruding the curve loop vertically to create a solid
    
    Args:
        area: Area element
        doc: Revit document
        
    Returns:
        Solid: 3D solid representing the Area volume, or None if creation fails
    """
    area_id = area.Id.IntegerValue
    try:
        logger.debug("Creating solid for area {} (ID: {})".format(area.Id, area_id))
        
        # Get boundary segments
        # Use default boundary location first (most reliable)
        boundary_options = SpatialElementBoundaryOptions()
        boundary_segments = area.GetBoundarySegments(boundary_options)
        
        if not boundary_segments or len(boundary_segments) == 0:
            logger.debug("Area {}: No boundary segments found".format(area_id))
            return None
        
        logger.debug("Area {}: Found {} boundary segment groups".format(area_id, len(boundary_segments)))
        
        # Use the first boundary loop (Areas typically have one main loop)
        first_loop = boundary_segments[0]
        if not first_loop or len(first_loop) == 0:
            logger.debug("Area {}: First boundary loop is empty".format(area_id))
            return None
        
        logger.debug("Area {}: First loop has {} segments".format(area_id, len(first_loop)))
        
        # Get Area's level first (needed for Z elevation)
        area_level = None
        if hasattr(area, "LevelId") and area.LevelId:
            area_level = doc.GetElement(area.LevelId)
        
        if not area_level:
            # Try to get level from parameter
            level_param = area.get_Parameter("Level")
            if level_param:
                level_id = level_param.AsElementId()
                if level_id:
                    area_level = doc.GetElement(level_id)
        
        if not area_level:
            logger.debug("Area {}: Could not find level".format(area_id))
            return None
        
        logger.debug("Area {}: Level found: {} (Elevation: {})".format(
            area_id, area_level.Name, area_level.Elevation))
        
        # Get level elevation
        base_z = area_level.Elevation
        
        # Find level above (next building storey level)
        # Get all levels sorted by elevation
        all_levels = FilteredElementCollector(doc)\
            .OfClass(Level)\
            .ToElements()
        
        levels_sorted = sorted(all_levels, key=lambda l: l.Elevation)
        
        # Find level above
        top_z = base_z + 10.0  # Default 10 feet if no level above
        for level in levels_sorted:
            if level.Elevation > base_z:
                top_z = level.Elevation
                break
        
        # Calculate extrusion height
        height = top_z - base_z
        if height <= 0:
            height = 10.0  # Fallback to 10 feet
        
        logger.debug("Area {}: Extrusion height: {} (base_z: {}, top_z: {})".format(
            area_id, height, base_z, top_z))
        
        # Create CurveLoop from boundary segments
        # CRITICAL: Build loop WITHOUT transformation first to maintain curve connectivity
        # Transforming individual curves can break the end-to-end connection
        curve_loop = CurveLoop()
        curve_z = None
        curves_appended = 0
        tolerance = 0.001  # Tolerance for curve connection (1mm in feet)
        
        for idx, segment in enumerate(first_loop):
            curve = segment.GetCurve()
            if not curve:
                logger.debug("Area {}: Segment {} has no curve".format(area_id, idx))
                continue
            
            # Get Z elevation from first curve (all curves should be at same Z)
            if curve_z is None:
                try:
                    curve_z = curve.GetEndPoint(0).Z
                    logger.debug("Area {}: Curve Z elevation: {}".format(area_id, curve_z))
                except Exception as e:
                    logger.debug("Area {}: Error getting curve Z: {}".format(area_id, str(e)))
                    curve_z = base_z
            
            # Append curve as-is - boundaries should already be properly connected
            try:
                curve_loop.Append(curve)
                curves_appended += 1
            except Exception as e:
                # If Append fails, curves don't connect - this causes "discontinuous loop" error
                logger.info("Area {}: Error appending curve {} to loop: {}".format(area_id, idx, str(e)))
                logger.debug("Area {}: Successfully appended {} curves before failure".format(area_id, curves_appended))
                return None
        
        logger.debug("Area {}: Successfully appended {} curves to loop".format(area_id, curves_appended))
        
        # Check if loop is valid and closed
        if curve_loop.IsOpen():
            logger.debug("Area {}: Curve loop is open (not closed)".format(area_id))
            return None
        
        # Verify loop has curves (check by counting appended curves, not using len())
        if curves_appended == 0:
            logger.debug("Area {}: No curves were appended to loop".format(area_id))
            return None
        
        logger.debug("Area {}: Curve loop is closed with {} curves".format(area_id, curves_appended))
        
        # Create extrusion direction (vertical, unit vector)
        extrusion_direction = XYZ(0, 0, 1)
        
        # Create solid by extruding the curve loop vertically
        # CreateExtrusionGeometry creates solid at the curve's current Z elevation
        try:
            logger.debug("Area {}: Creating extrusion geometry...".format(area_id))
            solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                [curve_loop],
                extrusion_direction,
                height
            )
            
            if not solid:
                logger.debug("Area {}: Extrusion geometry creation returned None".format(area_id))
                return None
            
            logger.debug("Area {}: Extrusion geometry created successfully".format(area_id))
            
            # Transform entire solid to correct Z position if needed
            # This preserves curve connectivity better than transforming individual curves
            if curve_z is not None and abs(curve_z - base_z) > 0.01:
                logger.debug("Area {}: Transforming solid from Z {} to Z {}".format(area_id, curve_z, base_z))
                translation = Transform.CreateTranslation(XYZ(0, 0, base_z - curve_z))
                solid = Solid.CreateTransformed(solid, translation)
            
            logger.debug("Area {}: Solid creation completed successfully".format(area_id))
            return solid
        except Exception as e:
            logger.info("Area {}: Error creating extrusion: {}".format(area_id, str(e)))
            logger.debug("Area {}: Extrusion error details - curves_appended: {}, height: {}, curve_z: {}, base_z: {}".format(
                area_id, curves_appended, height, curve_z, base_z))
            return None
        
    except Exception as e:
        logger.info("Area {}: Error creating solid from area: {}".format(area_id, str(e)))
        import traceback
        logger.debug("Area {}: Traceback: {}".format(area_id, traceback.format_exc()))
        return None

def is_point_in_area(area, point, doc):
    """Check if a point is inside an area using a 3D solid created from boundary segments.
    
    Areas are 2D planar elements. This function creates a 3D solid by extruding
    the Area's boundary between building storey levels, then checks if the point
    is inside that solid.
    
    Args:
        area: Area element
        point: XYZ point
        doc: Revit document
        
    Returns:
        bool: True if point is in area
    """
    try:
        element_id = area.Id.IntegerValue
        
        # Use cached solid if available
        if element_id in _geometry_cache:
            cached = _geometry_cache[element_id]
            solid = cached.get("solid")
            bbox = cached.get("bbox")
            
            # Check if this area failed solid creation (cached as None)
            if solid is None and "failed" in cached:
                return False  # Area failed solid creation, skip it
            
            # Fast bounding box rejection
            if bbox and not bbox.Contains(point):
                return False
            
            # Use optimized point-in-solid check
            if solid:
                return is_point_inside_solid_optimized(point, solid)
            return False
        
        # Fast bounding box pre-check before expensive solid creation
        bbox = area.get_BoundingBox(None)
        if bbox and not bbox.Contains(point):
            return False
        
        # Create solid from Area boundary
        solid = _create_solid_from_area(area, doc)
        if not solid:
            # Cache failure to avoid retrying for every target element
            _geometry_cache[element_id] = {"solid": None, "bbox": bbox, "failed": True}
            return False
        
        # Cache both solid and bounding box for future checks
        _geometry_cache[element_id] = {"solid": solid, "bbox": bbox}
        
        # Use optimized point-in-solid check
        return is_point_inside_solid_optimized(point, solid)
        
    except Exception as e:
        logger.info("Error checking point in area: {}".format(str(e)))
        return False

def is_point_inside_solid_optimized(point, solid):
    """Optimized point-in-solid check using learnrevitapi pattern.
    
    Uses SolidCurveIntersectionOptions instead of solid.IsInside() for better performance.
    
    Args:
        point: XYZ point to check
        solid: Solid geometry
        
    Returns:
        bool: True if point is inside solid
    """
    try:
        # Create tiny line from point (learnrevitapi best practice)
        line = Line.CreateBound(point, XYZ(point.X, point.Y, point.Z + 0.01))
        
        # Create intersection options
        opts = SolidCurveIntersectionOptions()
        opts.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside
        
        # Intersect line with solid
        sci = solid.IntersectWithCurve(line, opts)
        
        if sci:
            return True
        return False
    except Exception as e:
        logger.info("Error in optimized point-in-solid check: {}".format(str(e)))
        return False

def is_point_in_element(element, point, doc):
    """Check if a point is inside a Mass/Generic Model element using geometry.
    
    Optimized with bounding box pre-filtering and SolidCurveIntersectionOptions pattern.
    
    Args:
        element: Mass or Generic Model element
        point: XYZ point
        doc: Revit document
        
    Returns:
        bool: True if point is in element
    """
    try:
        element_id = element.Id.IntegerValue
        
        # Use cached geometry if available
        if element_id in _geometry_cache:
            cached = _geometry_cache[element_id]
            solid = cached.get("solid")
            bbox = cached.get("bbox")
            
            # Fast bounding box rejection
            if bbox and not bbox.Contains(point):
                return False
            
            # Use optimized point-in-solid check
            if solid:
                return is_point_inside_solid_optimized(point, solid)
            return False
        
        # Fast bounding box pre-check before expensive geometry calculation
        bbox = element.get_BoundingBox(None)
        if bbox and not bbox.Contains(point):
            return False
        
        # Calculate geometry (only if bounding box check passes)
        options = _get_geometry_options(doc)
        if doc and doc.ActiveView:
            options.DetailLevel = doc.ActiveView.DetailLevel
        
        geometry = element.get_Geometry(options)
        if not geometry:
            return False
        
        # Get first solid
        solid = None
        for geom_obj in geometry:
            if hasattr(geom_obj, "GetInstanceGeometry"):
                instance_geom = geom_obj.GetInstanceGeometry()
                for inst_obj in instance_geom:
                    if hasattr(inst_obj, "Volume") and inst_obj.Volume > 0:
                        solid = inst_obj
                        break
            elif hasattr(geom_obj, "Volume") and geom_obj.Volume > 0:
                solid = geom_obj
                break
        
        if not solid:
            return False
        
        # Cache both solid and bounding box for future checks
        _geometry_cache[element_id] = {"solid": solid, "bbox": bbox}
        
        # Use optimized point-in-solid check
        return is_point_inside_solid_optimized(point, solid)
        
    except Exception as e:
        logger.info("Error checking point in element: {}".format(str(e)))
        return False

def get_containing_room(element, doc, rooms_by_level=None):
    """Find the room containing an element.
    
    Uses multiple test points for better detection of walls and doors.
    Optimized with early exit - stops checking once containment is found.
    
    Args:
        element: Target element
        doc: Revit document
        rooms_by_level: Optional pre-grouped rooms dict by LevelId
        
    Returns:
        Room: Containing room or None
    """
    try:
        # Get multiple test points for better detection
        test_points = get_element_test_points(element)
        if not test_points:
            return None
        
        # Get rooms
        if rooms_by_level is None:
            rooms = FilteredElementCollector(doc)\
                .OfClass(Room)\
                .WhereElementIsNotElementType()\
                .ToElements()
            rooms_by_level = defaultdict(list)
            for room in rooms:
                if room.LevelId:
                    rooms_by_level[room.LevelId].append(room)
        
        # Check rooms on same level first
        element_level_id = None
        if hasattr(element, "LevelId"):
            element_level_id = element.LevelId
        elif hasattr(element, "get_Parameter"):
            level_param = element.get_Parameter("Level")
            if level_param:
                element_level_id = level_param.AsElementId()
        
        # Try each test point with early exit
        if element_level_id:
            for point in test_points:
                for room in rooms_by_level.get(element_level_id, []):
                    if is_point_in_room(room, point):
                        return room  # Early exit - found containment
        
        # Fallback: check all rooms with all test points (with early exit)
        for point in test_points:
            for room_list in rooms_by_level.values():
                for room in room_list:
                    if is_point_in_room(room, point):
                        return room  # Early exit - found containment
        
        return None
    except Exception as e:
        logger.info("Error getting containing room: {}".format(str(e)))
        return None

def get_containing_space(element, doc, spaces_by_level=None):
    """Find the space containing an element.
    
    Uses multiple test points for better detection of walls and doors.
    Optimized with early exit - stops checking once containment is found.
    
    Args:
        element: Target element
        doc: Revit document
        spaces_by_level: Optional pre-grouped spaces dict by LevelId
        
    Returns:
        Space: Containing space or None
    """
    try:
        # Get multiple test points for better detection
        test_points = get_element_test_points(element)
        if not test_points:
            return None
        
        # Get spaces
        if spaces_by_level is None:
            spaces = FilteredElementCollector(doc)\
                .OfClass(Space)\
                .WhereElementIsNotElementType()\
                .ToElements()
            spaces_by_level = defaultdict(list)
            for space in spaces:
                if space.LevelId:
                    spaces_by_level[space.LevelId].append(space)
        
        # Check spaces on same level first
        element_level_id = None
        if hasattr(element, "LevelId"):
            element_level_id = element.LevelId
        elif hasattr(element, "get_Parameter"):
            level_param = element.get_Parameter("Level")
            if level_param:
                element_level_id = level_param.AsElementId()
        
        # Try each test point with early exit
        if element_level_id:
            for point in test_points:
                for space in spaces_by_level.get(element_level_id, []):
                    if is_point_in_space(space, point):
                        return space  # Early exit - found containment
        
        # Fallback: check all spaces with all test points (with early exit)
        for point in test_points:
            for space_list in spaces_by_level.values():
                for space in space_list:
                    if is_point_in_space(space, point):
                        return space  # Early exit - found containment
        
        return None
    except Exception as e:
        logger.info("Error getting containing space: {}".format(str(e)))
        return None

def get_containing_area(element, doc, areas_by_level=None):
    """Find the area containing an element.
    
    Uses multiple test points for better detection of walls and doors.
    Optimized with early exit - stops checking once containment is found.
    
    Args:
        element: Target element
        doc: Revit document
        areas_by_level: Optional pre-grouped areas dict by LevelId
        
    Returns:
        Area: Containing area or None
    """
    try:
        # Get multiple test points for better detection
        test_points = get_element_test_points(element)
        if not test_points:
            logger.debug("[DEBUG] get_containing_area: No test points for element {} (ID: {})".format(
                element.GetType().Name, element.Id))
            return None
        
        # Get areas
        if areas_by_level is None:
            # Fallback: get all areas (less efficient)
            spatial_elements = FilteredElementCollector(doc)\
                .OfClass(SpatialElement)\
                .WhereElementIsNotElementType()\
                .ToElements()
            
            # Filter for Area instances
            areas = [elem for elem in spatial_elements if isinstance(elem, Area)]
            areas_by_level = defaultdict(list)
            for area in areas:
                if hasattr(area, "LevelId") and area.LevelId:
                    areas_by_level[area.LevelId].append(area)
        
        # Check areas on same level first
        element_level_id = None
        if hasattr(element, "LevelId"):
            element_level_id = element.LevelId
        elif hasattr(element, "get_Parameter"):
            level_param = element.get_Parameter("Level")
            if level_param:
                element_level_id = level_param.AsElementId()
        
        # Try each test point with early exit
        areas_checked = 0
        failed_areas_skipped = 0
        
        if element_level_id:
            level_areas = areas_by_level.get(element_level_id, [])
            logger.debug("[DEBUG] get_containing_area: Checking {} areas on same level for element {} (ID: {})".format(
                len(level_areas), element.GetType().Name, element.Id))
            
            for point in test_points:
                for area in level_areas:
                    areas_checked += 1
                    # Skip areas that failed solid creation during precomputation
                    area_id = area.Id.IntegerValue
                    if area_id in _geometry_cache:
                        cached = _geometry_cache[area_id]
                        if cached.get("failed", False):
                            failed_areas_skipped += 1
                            continue  # Skip failed areas
                    if is_point_in_area(area, point, doc):
                        logger.debug("[DEBUG] get_containing_area: Found containing area {} (ID: {}) for element {} (ID: {})".format(
                            area.Id, area_id, element.GetType().Name, element.Id))
                        return area  # Early exit - found containment
        
        # Fallback: check all areas with all test points (with early exit)
        if not element_level_id or areas_checked == 0:
            total_areas = sum(len(area_list) for area_list in areas_by_level.values())
            logger.debug("[DEBUG] get_containing_area: Checking all {} areas (no level match) for element {} (ID: {})".format(
                total_areas, element.GetType().Name, element.Id))
            
            for point in test_points:
                for area_list in areas_by_level.values():
                    for area in area_list:
                        areas_checked += 1
                        # Skip areas that failed solid creation during precomputation
                        area_id = area.Id.IntegerValue
                        if area_id in _geometry_cache:
                            cached = _geometry_cache[area_id]
                            if cached.get("failed", False):
                                failed_areas_skipped += 1
                                continue  # Skip failed areas
                        if is_point_in_area(area, point, doc):
                            logger.debug("[DEBUG] get_containing_area: Found containing area {} (ID: {}) for element {} (ID: {})".format(
                                area.Id, area_id, element.GetType().Name, element.Id))
                            return area  # Early exit - found containment
        
        if failed_areas_skipped > 0:
            logger.debug("[DEBUG] get_containing_area: Skipped {} areas that failed solid creation".format(failed_areas_skipped))
        
        logger.debug("[DEBUG] get_containing_area: No containing area found for element {} (ID: {}), checked {} areas".format(
            element.GetType().Name, element.Id, areas_checked))
        return None
    except Exception as e:
        logger.info("Error getting containing area: {}".format(str(e)))
        return None

def get_containing_element(element, doc, source_categories):
    """Find the containing element (Mass/Generic Model) for an element.
    
    Optimized with Quick Filter pre-filtering using BoundingBoxContainsPointFilter
    (learnrevitapi best practice - filters at database level before expanding elements).
    
    Args:
        element: Target element
        doc: Revit document
        source_categories: List of BuiltInCategory values to search
        
    Returns:
        Element: Containing element or None
    """
    try:
        point = get_element_representative_point(element)
        if not point:
            return None
        
        # CRITICAL OPTIMIZATION: Use Quick Filter to pre-filter candidates
        # BoundingBoxContainsPointFilter operates on ElementRecord (database level)
        # This is 10-100x faster than expanding all elements and checking geometry
        point_filter = BoundingBoxContainsPointFilter(point)
        
        # Build collector with category filters
        collector = FilteredElementCollector(doc)\
            .WhereElementIsNotElementType()
        
        # Filter by categories
        for category in source_categories:
            collector = collector.OfCategory(category)
        
        # Apply Quick Filter BEFORE ToElements() - filters at database level
        candidate_elements = collector.WherePasses(point_filter).ToElements()
        
        # Now check geometry only on pre-filtered candidates (much fewer elements)
        for source_el in candidate_elements:
            if is_point_in_element(source_el, point, doc):
                return source_el  # Early exit - found containment
        
        return None
    except Exception as e:
        logger.info("Error getting containing element: {}".format(str(e)))
        return None

def get_containing_element_by_strategy(element, doc, strategy, source_categories=None, 
                                     rooms_by_level=None, spaces_by_level=None, areas_by_level=None):
    """Unified function that routes to appropriate containment method.
    
    Args:
        element: Target element
        doc: Revit document
        strategy: Containment strategy ("room", "space", "area", "element")
        source_categories: List of BuiltInCategory values (for element strategy)
        rooms_by_level: Optional pre-grouped rooms dict
        spaces_by_level: Optional pre-grouped spaces dict
        areas_by_level: Optional pre-grouped areas dict
        
    Returns:
        Element: Containing element or None
    """
    if strategy == "room":
        return get_containing_room(element, doc, rooms_by_level)
    elif strategy == "space":
        return get_containing_space(element, doc, spaces_by_level)
    elif strategy == "area":
        return get_containing_area(element, doc, areas_by_level)
    elif strategy == "element":
        return get_containing_element(element, doc, source_categories)
    else:
        return None

def clear_geometry_cache():
    """Clear the geometry cache."""
    global _geometry_cache, _geometry_options
    _geometry_cache = {}
    _geometry_options = None

def precompute_geometries(elements, doc):
    """Pre-compute geometries for a batch of elements.
    
    This allows batch geometry extraction in a single pass, improving performance
    when processing many elements. Handles both Mass/Generic Model elements
    (which have native geometry) and Areas (which need solid creation from boundaries).
    
    Args:
        elements: List of elements to pre-compute geometries for
        doc: Revit document
    """
    global _geometry_cache
    
    logger.info("[DEBUG] Precomputing geometries for {} elements".format(len(elements)))
    
    options = _get_geometry_options(doc)
    if doc and doc.ActiveView:
        options.DetailLevel = doc.ActiveView.DetailLevel
    
    area_count = 0
    area_success_count = 0
    area_fail_count = 0
    
    for element in elements:
        element_id = element.Id.IntegerValue
        
        # Skip if already cached
        if element_id in _geometry_cache:
            continue
        
        try:
            # Get bounding box first (fast)
            bbox = element.get_BoundingBox(None)
            
            # Handle Areas specially - create solid from boundary
            if isinstance(element, Area):
                area_count += 1
                solid = _create_solid_from_area(element, doc)
                if solid:
                    # Cache both solid and bounding box
                    _geometry_cache[element_id] = {"solid": solid, "bbox": bbox}
                    area_success_count += 1
                else:
                    # Cache failure to avoid retrying for every target element
                    _geometry_cache[element_id] = {"solid": None, "bbox": bbox, "failed": True}
                    area_fail_count += 1
                    logger.debug("[DEBUG] Failed to create solid for area {} (ID: {})".format(element.Id, element_id))
                continue
            
            # For Mass/Generic Model elements, extract geometry
            geometry = element.get_Geometry(options)
            if not geometry:
                continue
            
            # Extract solid
            solid = None
            for geom_obj in geometry:
                if hasattr(geom_obj, "GetInstanceGeometry"):
                    instance_geom = geom_obj.GetInstanceGeometry()
                    for inst_obj in instance_geom:
                        if hasattr(inst_obj, "Volume") and inst_obj.Volume > 0:
                            solid = inst_obj
                            break
                elif hasattr(geom_obj, "Volume") and geom_obj.Volume > 0:
                    solid = geom_obj
                    break
            
            if solid:
                # Cache both solid and bounding box
                _geometry_cache[element_id] = {"solid": solid, "bbox": bbox}
        except Exception as e:
            logger.info("Error precomputing geometry for element {}: {}".format(element_id, str(e)))
            continue
    
    if area_count > 0:
        logger.info("[DEBUG] Area geometry precomputation: {} total, {} succeeded, {} failed".format(
            area_count, area_success_count, area_fail_count))

