# -*- coding: utf-8 -*-
"""Spatial containment detection for 3D Zone parameter mapping."""

from collections import defaultdict
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, XYZ, UV,
    SpatialElementBoundaryOptions, SpatialElementBoundaryLocation,
    BoundingBoxIntersectsFilter, BoundingBoxContainsPointFilter, 
    Outline, ElementIntersectsSolidFilter, Options, SpatialElement, 
    Line, SolidCurveIntersectionOptions, SolidCurveIntersectionMode, 
    Area, Level, CurveLoop, Solid, GeometryCreationUtilities, Transform,
    BoundingBoxXYZ, BuiltInParameter, Phase, ElementId, FamilyInstance,
    HostObject, HostObjectUtils, ShellLayerType, PlanarFace
)
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.DB.Mechanical import Space
from pyrevit import script

# Initialize logger
logger = script.get_logger()


def get_phase_map_for_link(host_doc, link_instance):
    """Get phase mapping from host document to linked document using Revit's API.
    
    Uses RevitLinkType.GetPhaseMap() which respects user-configured phase mappings
    set in Revit UI (Type Properties > Phase Mapping > Edit).
    
    Args:
        host_doc: The host Revit document
        link_instance: RevitLinkInstance object
        
    Returns:
        dict: Maps host phase ElementId -> linked phase ElementId, or None if unavailable
    """
    if not link_instance:
        return None
    
    try:
        link_type_id = link_instance.GetTypeId()
        if not link_type_id or link_type_id.IntegerValue < 0:
            return None
            
        link_type = host_doc.GetElement(link_type_id)
        if link_type and hasattr(link_type, 'GetPhaseMap'):
            phase_map = link_type.GetPhaseMap()
            logger.debug("Got phase map from RevitLinkType with {} mappings".format(
                len(phase_map) if phase_map else 0))
            return phase_map
    except Exception as e:
        logger.debug("Error getting phase map for link: {}".format(str(e)))
    
    return None


# IMPORTANT: Revit Unit System
# Revit's internal units for length are ALWAYS feet, regardless of project display units.
# Level.Elevation, XYZ coordinates, and all geometric calculations use internal units (feet).
# This means:
# - Level.Elevation returns elevation in feet
# - XYZ coordinates are in feet
# - Height calculations are in feet
# - No unit conversion is needed for internal calculations

# Default storey height in Revit internal units (feet)
# Used when no level above is found for Area solid creation
# 10 feet ≈ 3.05 meters (typical storey height)
DEFAULT_STOREY_HEIGHT_FEET = 10.0

# Cache for geometry calculations (Mass/Generic Model)
# Structure: element_id -> {"solids": [Solid, ...], "bbox": BoundingBox}
# NOTE: We store a list of solids to handle in-place families with multiple solid forms
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

def is_point_in_bbox(point, bbox):
    """Check if a point is inside a BoundingBoxXYZ.
    
    Args:
        point: XYZ point
        bbox: BoundingBoxXYZ
        
    Returns:
        bool: True if point is inside bounding box
    """
    if not bbox or not point:
        return False
    return (bbox.Min.X <= point.X <= bbox.Max.X and
            bbox.Min.Y <= point.Y <= bbox.Max.Y and
            bbox.Min.Z <= point.Z <= bbox.Max.Z)

def is_inplace_family(element):
    """Check if an element is an in-place family instance.
    
    In-place families are created directly in a project (not loaded from .rfa files).
    They typically have:
    - Location = None or Location.Point at project origin (0,0,0)
    - Multiple solid forms in their geometry
    - Family.IsInPlace = True
    
    Args:
        element: Revit element to check
        
    Returns:
        bool: True if element is an in-place family instance
    """
    try:
        if not isinstance(element, FamilyInstance):
            return False
        # Access the Family through the Symbol (FamilySymbol)
        family = element.Symbol.Family
        return family.IsInPlace
    except Exception:
        return False

def _get_inplace_family_centroid(element, doc=None):
    """Get centroid from in-place family geometry.
    
    In-place families often have Location = None or Location.Point at origin.
    This function extracts geometry and computes the centroid of the first
    solid found, providing a more accurate representative point.
    
    Args:
        element: FamilyInstance (in-place family)
        doc: Revit document (optional, for geometry options)
        
    Returns:
        XYZ: Centroid point or None if cannot be determined
    """
    try:
        options = Options()
        options.ComputeReferences = False
        
        geometry = element.get_Geometry(options)
        if not geometry:
            return None
        
        # Find the first solid and compute its centroid
        for geom_obj in geometry:
            if hasattr(geom_obj, "GetInstanceGeometry"):
                instance_geom = geom_obj.GetInstanceGeometry()
                if instance_geom:
                    for inst_obj in instance_geom:
                        if hasattr(inst_obj, "ComputeCentroid"):
                            try:
                                centroid = inst_obj.ComputeCentroid()
                                if centroid:
                                    return centroid
                            except:
                                pass
            elif hasattr(geom_obj, "ComputeCentroid"):
                try:
                    centroid = geom_obj.ComputeCentroid()
                    if centroid:
                        return centroid
                except:
                    pass
        
        return None
    except Exception:
        return None

def get_element_representative_point(element, doc=None):
    """Get a representative point from an element for containment testing.
    
    Areas are 2D planar elements and are only used as sources (not targets),
    but we handle them here for robustness.
    
    IMPORTANT: In-place families require special handling because:
    - Their Location property is often None
    - Or Location.Point is at project origin (0,0,0), not the actual element location
    For these elements, we compute the centroid from the solid geometry.
    
    Args:
        element: Revit element
        doc: Revit document (optional, used for in-place family geometry extraction)
        
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
        
        # Special handling for in-place families
        # These often have Location = None or Location.Point at origin
        if is_inplace_family(element):
            # Try to get centroid from solid geometry (most accurate)
            centroid = _get_inplace_family_centroid(element, doc)
            if centroid:
                return centroid
            # Fall back to bounding box center
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
            # LocationPoint - but check if it's at origin (may be unreliable)
            point = loc.Point
            # If point is very close to origin (0,0,0), it may be unreliable
            # Fall back to bounding box center in that case
            if abs(point.X) < 0.001 and abs(point.Y) < 0.001 and abs(point.Z) < 0.001:
                bbox = element.get_BoundingBox(None)
                if bbox:
                    bbox_center = (bbox.Min + bbox.Max) / 2.0
                    # Only use bbox center if it's significantly different from origin
                    if abs(bbox_center.X) > 0.1 or abs(bbox_center.Y) > 0.1 or abs(bbox_center.Z) > 0.1:
                        return bbox_center
            return point
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
        logger.debug("Error getting element point: {}".format(str(e)))
        return None

def _get_inplace_family_test_points(element, doc=None):
    """Get multiple test points from in-place family geometry.
    
    For in-place families with multiple solids, we generate test points
    from each solid's centroid. This ensures containment is detected
    correctly even for complex multi-form in-place families.
    
    Args:
        element: FamilyInstance (in-place family)
        doc: Revit document (optional)
        
    Returns:
        list: List of XYZ points (centroids of all solids)
    """
    points = []
    try:
        options = Options()
        options.ComputeReferences = False
        
        geometry = element.get_Geometry(options)
        if not geometry:
            return points
        
        # Collect centroids from ALL solids
        for geom_obj in geometry:
            if hasattr(geom_obj, "GetInstanceGeometry"):
                instance_geom = geom_obj.GetInstanceGeometry()
                if instance_geom:
                    for inst_obj in instance_geom:
                        if hasattr(inst_obj, "ComputeCentroid"):
                            try:
                                centroid = inst_obj.ComputeCentroid()
                                if centroid:
                                    points.append(centroid)
                            except:
                                pass
            elif hasattr(geom_obj, "ComputeCentroid"):
                try:
                    centroid = geom_obj.ComputeCentroid()
                    if centroid:
                        points.append(centroid)
                except:
                    pass
        
        # Also add bounding box center as a fallback point
        if not points:
            bbox = element.get_BoundingBox(None)
            if bbox:
                points.append((bbox.Min + bbox.Max) / 2.0)
        
        return points
    except Exception:
        return points

def _generate_grid_points_on_face(face, grid_size=5):
    """Generate a grid of test points on a planar face.
    
    Uses UV parameterization to create evenly distributed points across the face.
    Only includes points that are actually inside the face boundary (handles
    irregular face shapes like L-shaped floors).
    
    Args:
        face: PlanarFace object from Revit geometry
        grid_size: Number of points per axis (default 5 = 25 points total)
        
    Returns:
        list: List of XYZ points on the face surface
    """
    points = []
    try:
        bbox = face.GetBoundingBox()
        if not bbox:
            return points
        
        u_min, u_max = bbox.Min.U, bbox.Max.U
        v_min, v_max = bbox.Min.V, bbox.Max.V
        
        # Generate grid points using UV parameterization
        for i in range(grid_size):
            for j in range(grid_size):
                # Calculate UV coordinates at cell centers (offset by 0.5)
                u = u_min + (u_max - u_min) * (i + 0.5) / grid_size
                v = v_min + (v_max - v_min) * (j + 0.5) / grid_size
                uv = UV(u, v)
                
                # Only include points that are inside the face boundary
                # This handles irregular face shapes (L-shaped, with holes, etc.)
                try:
                    # IsInside returns IntersectionResult, check if point is inside
                    result = face.IsInside(uv)
                    # In IronPython, IsInside returns a tuple (bool, IntersectionResult)
                    # or just bool depending on overload
                    is_inside = result[0] if isinstance(result, tuple) else result
                    if is_inside:
                        xyz = face.Evaluate(uv)
                        if xyz:
                            points.append(xyz)
                except:
                    # Fallback: try to evaluate anyway
                    try:
                        xyz = face.Evaluate(uv)
                        if xyz:
                            points.append(xyz)
                    except:
                        pass
        
        return points
    except Exception as e:
        logger.debug("Error generating grid points on face: {}".format(str(e)))
        return points

def _get_host_object_test_points(element, doc=None, grid_size=5):
    """Get grid-based test points for HostObject elements (floors, ceilings, roofs, walls).
    
    Uses HostObjectUtils to extract actual face geometry:
    - Floors/Ceilings/Roofs: Uses GetTopFaces() to get horizontal surfaces
    - Walls: Uses GetSideFaces() for both exterior and interior faces
    
    This provides much better containment detection than bounding box center,
    especially for large elements that span multiple zones.
    
    Args:
        element: HostObject (Floor, Ceiling, Roof, or Wall)
        doc: Revit document (optional)
        grid_size: Number of points per axis for grid (default 5 = 25 points per face)
        
    Returns:
        list: List of XYZ points distributed across the element's faces
    """
    points = []
    
    if not isinstance(element, HostObject):
        return points
    
    try:
        # Check if element is a wall (has WallType property)
        is_wall = hasattr(element, 'WallType')
        
        if is_wall:
            # For walls: get both exterior and interior side faces
            for side in [ShellLayerType.Exterior, ShellLayerType.Interior]:
                try:
                    face_refs = HostObjectUtils.GetSideFaces(element, side)
                    if face_refs:
                        for ref in face_refs:
                            try:
                                face = element.GetGeometryObjectFromReference(ref)
                                if face and isinstance(face, PlanarFace):
                                    face_points = _generate_grid_points_on_face(face, grid_size)
                                    points.extend(face_points)
                            except Exception as face_error:
                                logger.debug("Error getting wall face geometry: {}".format(str(face_error)))
                                continue
                except Exception as side_error:
                    logger.debug("Error getting wall side faces: {}".format(str(side_error)))
                    continue
        else:
            # For floors/ceilings/roofs: get top faces
            try:
                face_refs = HostObjectUtils.GetTopFaces(element)
                if face_refs:
                    for ref in face_refs:
                        try:
                            face = element.GetGeometryObjectFromReference(ref)
                            if face and isinstance(face, PlanarFace):
                                face_points = _generate_grid_points_on_face(face, grid_size)
                                points.extend(face_points)
                        except Exception as face_error:
                            logger.debug("Error getting top face geometry: {}".format(str(face_error)))
                            continue
            except Exception as top_error:
                logger.debug("Error getting top faces: {}".format(str(top_error)))
        
        return points
    except Exception as e:
        logger.debug("Error getting host object test points: {}".format(str(e)))
        return points

def get_element_test_points(element, doc=None):
    """Get multiple test points for an element to improve containment detection.
    
    For walls and linear elements, returns multiple points along the element.
    For point-based elements, returns a single point.
    
    Areas are 2D planar elements and are only used as sources (not targets),
    but we handle them here for robustness.
    
    IMPORTANT: In-place families require special handling because:
    - Their Location property is often None or at project origin
    - They may have multiple solid forms
    For these elements, we generate test points from solid centroids.
    
    Args:
        element: Revit element
        doc: Revit document (optional, used for in-place family handling)
        
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
        
        # Special handling for in-place families
        # These often have Location = None or unreliable Location.Point
        if is_inplace_family(element):
            inplace_points = _get_inplace_family_test_points(element, doc)
            if inplace_points:
                return inplace_points
            # Fall back to bounding box center
            bbox = element.get_BoundingBox(None)
            if bbox:
                points.append((bbox.Min + bbox.Max) / 2.0)
            return points
        
        # Special handling for HostObjects (floors, ceilings, roofs, walls)
        # These elements benefit from grid-based sampling across their faces
        # for better containment detection when spanning multiple zones
        if isinstance(element, HostObject):
            host_points = _get_host_object_test_points(element, doc)
            if host_points:
                return host_points
            # Fall through to Location/bbox fallback if face extraction fails
        
        loc = element.Location
        if loc is None:
            bbox = element.get_BoundingBox(None)
            if bbox:
                points.append((bbox.Min + bbox.Max) / 2.0)
            return points
        
        if hasattr(loc, "Point"):
            # LocationPoint - get vertical test points for columns/vertical elements
            base_point = loc.Point
            bbox = element.get_BoundingBox(None)
            
            if bbox and abs(bbox.Max.Z - bbox.Min.Z) > 0.1:
                # Element has vertical extent - add points at 25%, 50%, 75% of height
                min_z = bbox.Min.Z
                max_z = bbox.Max.Z
                for t in [0.25, 0.5, 0.75]:
                    z = min_z + (max_z - min_z) * t
                    points.append(XYZ(base_point.X, base_point.Y, z))
            else:
                # No significant vertical extent - use base point only
                points.append(base_point)
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
                    # Try both directions with offset (2.0 feet in Revit internal units)
                    # This offset helps detect containment for walls/linear elements, especially wide walls
                    points.append(midpoint + perp_direction * 2.0)
                    points.append(midpoint - perp_direction * 2.0)
                except:
                    pass
        
        # Fallback to bounding box center
        if not points:
            bbox = element.get_BoundingBox(None)
            if bbox:
                points.append((bbox.Min + bbox.Max) / 2.0)
        
        return points
    except Exception as e:
        logger.debug("Error getting element test points: {}".format(str(e)))
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
        logger.debug("Error checking point in room: {}".format(str(e)))
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
        logger.debug("Error checking point in space: {}".format(str(e)))
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
        
        # Get level elevation (in Revit internal units: feet)
        base_z = area_level.Elevation
        
        # Find level above (next building storey level)
        # Get all levels sorted by elevation
        all_levels = FilteredElementCollector(doc)\
            .OfClass(Level)\
            .ToElements()
        
        levels_sorted = sorted(all_levels, key=lambda l: l.Elevation)
        
        # Find level above (all elevations in Revit internal units: feet)
        top_z = base_z + DEFAULT_STOREY_HEIGHT_FEET  # Default if no level above
        for level in levels_sorted:
            if level.Elevation > base_z:
                top_z = level.Elevation
                break
        
        # Calculate extrusion height (in Revit internal units: feet)
        height = top_z - base_z
        if height <= 0:
            height = DEFAULT_STOREY_HEIGHT_FEET  # Fallback to default storey height
        
        logger.debug("Area {}: Extrusion height: {} (base_z: {}, top_z: {})".format(
            area_id, height, base_z, top_z))
        
        # Create CurveLoop from boundary segments
        # CRITICAL: Build loop WITHOUT transformation first to maintain curve connectivity
        # Transforming individual curves can break the end-to-end connection
        curve_loop = CurveLoop()
        curve_z = None
        curves_appended = 0
        # Tolerance for curve connection (0.001 feet ≈ 0.3mm in Revit internal units)
        tolerance = 0.001
        
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
                logger.debug("Area {}: Error appending curve {} to loop: {}".format(area_id, idx, str(e)))
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
            logger.debug("Area {}: Error creating extrusion: {}".format(area_id, str(e)))
            logger.debug("Area {}: Extrusion error details - curves_appended: {}, height: {}, curve_z: {}, base_z: {}".format(
                area_id, curves_appended, height, curve_z, base_z))
            return None
        
    except Exception as e:
        logger.debug("Area {}: Error creating solid from area: {}".format(area_id, str(e)))
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
            if bbox and not is_point_in_bbox(point, bbox):
                return False
            
            # Use optimized point-in-solid check
            if solid:
                return is_point_inside_solid_optimized(point, solid)
            return False
        
        # Fast bounding box pre-check before expensive solid creation
        bbox = area.get_BoundingBox(None)
        if bbox and not is_point_in_bbox(point, bbox):
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
        logger.debug("Error checking point in area: {}".format(str(e)))
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
        # Offset of 0.01 feet (≈3mm) in Revit internal units for point-in-solid check
        line = Line.CreateBound(point, XYZ(point.X, point.Y, point.Z + 0.01))
        
        # Create intersection options
        opts = SolidCurveIntersectionOptions()
        opts.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside
        
        # Intersect line with solid
        sci = solid.IntersectWithCurve(line, opts)
        
        # Check if there are actual segments inside the solid
        # sci can be non-None but have SegmentCount of 0 if point is outside
        if sci and sci.SegmentCount > 0:
            return True
        return False
    except Exception as e:
        logger.debug("Error in optimized point-in-solid check: {}".format(str(e)))
        return False

def _collect_all_solids_from_geometry(geometry):
    """Extract ALL solids from geometry, including nested GeometryInstance objects.
    
    This is critical for in-place families which often have multiple solid forms.
    Standard families may also have multiple solids that need to be checked.
    
    Args:
        geometry: GeometryElement from element.get_Geometry()
        
    Returns:
        list: List of Solid objects with Volume > 0
    """
    solids = []
    if not geometry:
        return solids
    
    for geom_obj in geometry:
        # Check if it's a GeometryInstance (wrapper around family geometry)
        if hasattr(geom_obj, "GetInstanceGeometry"):
            # GetInstanceGeometry() returns geometry in project coordinates
            # This is correct for containment testing
            instance_geom = geom_obj.GetInstanceGeometry()
            if instance_geom:
                for inst_obj in instance_geom:
                    if hasattr(inst_obj, "Volume") and inst_obj.Volume > 0:
                        solids.append(inst_obj)
        # Direct solid in geometry
        elif hasattr(geom_obj, "Volume") and geom_obj.Volume > 0:
            solids.append(geom_obj)
    
    return solids

def is_point_in_element(element, point, doc):
    """Check if a point is inside a Mass/Generic Model element using geometry.
    
    Optimized with bounding box pre-filtering and SolidCurveIntersectionOptions pattern.
    
    IMPORTANT: This function now collects ALL solids from the element geometry,
    not just the first one. This is critical for in-place families which often
    have multiple solid forms (e.g., multiple extrusions, blends, etc.).
    The point is considered "inside" if it's inside ANY of the solids.
    
    Args:
        element: Mass or Generic Model element (including in-place families)
        point: XYZ point
        doc: Revit document
        
    Returns:
        bool: True if point is in element (inside any of its solids)
    """
    try:
        element_id = element.Id.IntegerValue
        
        # Use cached geometry if available
        if element_id in _geometry_cache:
            cached = _geometry_cache[element_id]
            solids = cached.get("solids", [])
            bbox = cached.get("bbox")
            
            # Fast bounding box rejection
            if bbox and not is_point_in_bbox(point, bbox):
                return False
            
            # Check if point is inside ANY of the cached solids
            for solid in solids:
                if is_point_inside_solid_optimized(point, solid):
                    return True
            return False
        
        # Fast bounding box pre-check before expensive geometry calculation
        bbox = element.get_BoundingBox(None)
        if bbox and not is_point_in_bbox(point, bbox):
            return False
        
        # Calculate geometry (only if bounding box check passes)
        options = _get_geometry_options(doc)
        if doc and doc.ActiveView:
            options.DetailLevel = doc.ActiveView.DetailLevel
        
        geometry = element.get_Geometry(options)
        if not geometry:
            return False
        
        # Collect ALL solids from geometry (critical for in-place families)
        solids = _collect_all_solids_from_geometry(geometry)
        
        if not solids:
            return False
        
        # Cache all solids and bounding box for future checks
        _geometry_cache[element_id] = {"solids": solids, "bbox": bbox}
        
        # Log if this is an in-place family with multiple solids (useful for debugging)
        if len(solids) > 1 and is_inplace_family(element):
            logger.debug("In-place family element {} has {} solids".format(element_id, len(solids)))
        
        # Check if point is inside ANY of the solids
        for solid in solids:
            if is_point_inside_solid_optimized(point, solid):
                return True
        return False
        
    except Exception as e:
        logger.debug("Error checking point in element: {}".format(str(e)))
        return False

def get_room_phase_id(room):
    """Get the phase ID for a room.
    
    Args:
        room: Room element
        
    Returns:
        ElementId: Phase ID or None if not found
    """
    try:
        if not isinstance(room, Room):
            return None
        
        # Try BuiltInParameter.ROOM_PHASE first (most reliable)
        phase_param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
        if phase_param and phase_param.HasValue:
            phase_id = phase_param.AsElementId()
            if phase_id and phase_id != ElementId.InvalidElementId:
                return phase_id
        
        # Fallback to PhaseId property if available
        if hasattr(room, "PhaseId") and room.PhaseId:
            if room.PhaseId != ElementId.InvalidElementId:
                return room.PhaseId
        
        return None
    except Exception as e:
        logger.debug("Error getting room phase ID: {}".format(str(e)))
        return None

def get_ordered_phases(doc):
    """Get phases in model order (ascending sequence).
    
    Args:
        doc: Revit document
        
    Returns:
        list: List of tuples (phase, order_index) sorted by sequence
    """
    try:
        phases = FilteredElementCollector(doc)\
            .OfClass(Phase)\
            .ToElements()
        
        # Sort by SequenceNumber if available, otherwise by enumeration order
        phases_list = list(phases)
        try:
            # Try sorting by SequenceNumber (most reliable)
            phases_list.sort(key=lambda p: p.SequenceNumber if hasattr(p, "SequenceNumber") else 0)
        except:
            # Fallback: keep original order (phases are typically already in order)
            pass
        
        # Return list of (phase, index) tuples
        ordered = [(phase, idx) for idx, phase in enumerate(phases_list)]
        return ordered
    except Exception as e:
        logger.debug("Error getting ordered phases: {}".format(str(e)))
        return []

def get_element_phase_range(element, ordered_phases):
    """Get the phase range where an element exists.
    
    An element exists in a phase if:
    - CreatedPhaseId <= phase AND
    - (DemolishedPhaseId is invalid OR phase < DemolishedPhaseId)
    
    Args:
        element: Element to check
        ordered_phases: List of (phase, order_index) tuples from get_ordered_phases()
        
    Returns:
        tuple: (start_idx, end_idx) indices into ordered_phases, or (0, len(ordered_phases)) if no phase data
    """
    try:
        if not ordered_phases:
            return (0, 0)
        
        # Get CreatedPhaseId
        created_phase_id = None
        if hasattr(element, "CreatedPhaseId") and element.CreatedPhaseId:
            if element.CreatedPhaseId != ElementId.InvalidElementId:
                created_phase_id = element.CreatedPhaseId
        elif hasattr(element, "get_Parameter"):
            created_param = element.get_Parameter(BuiltInParameter.PHASE_CREATED)
            if created_param and created_param.HasValue:
                created_phase_id = created_param.AsElementId()
                if not created_phase_id or created_phase_id == ElementId.InvalidElementId:
                    created_phase_id = None
        
        # Get DemolishedPhaseId
        demolished_phase_id = None
        if hasattr(element, "DemolishedPhaseId") and element.DemolishedPhaseId:
            if element.DemolishedPhaseId != ElementId.InvalidElementId:
                demolished_phase_id = element.DemolishedPhaseId
        elif hasattr(element, "get_Parameter"):
            demolished_param = element.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED)
            if demolished_param and demolished_param.HasValue:
                demolished_phase_id = demolished_param.AsElementId()
                if not demolished_phase_id or demolished_phase_id == ElementId.InvalidElementId:
                    demolished_phase_id = None
        
        # If no phase data, assume element exists in all phases
        if not created_phase_id:
            return (0, len(ordered_phases))
        
        # Get created phase sequence number for comparison
        created_phase_seq = None
        for phase, _ in ordered_phases:
            if phase.Id == created_phase_id:
                created_phase_seq = phase.SequenceNumber if hasattr(phase, "SequenceNumber") else None
                break
        
        # Get demolished phase sequence number for comparison
        demolished_phase_seq = None
        if demolished_phase_id:
            for phase, _ in ordered_phases:
                if phase.Id == demolished_phase_id:
                    demolished_phase_seq = phase.SequenceNumber if hasattr(phase, "SequenceNumber") else None
                    break
        
        # Find start index (first phase >= CreatedPhaseId by sequence)
        start_idx = 0
        if created_phase_seq is not None:
            for idx, (phase, _) in enumerate(ordered_phases):
                phase_seq = phase.SequenceNumber if hasattr(phase, "SequenceNumber") else idx
                if phase.Id == created_phase_id:
                    start_idx = idx
                    break
                elif phase_seq >= created_phase_seq:
                    # Phase is at or after creation, start here
                    start_idx = idx
                    break
        
        # Find end index (first phase >= DemolishedPhaseId, or end of list)
        end_idx = len(ordered_phases)
        if demolished_phase_id and demolished_phase_seq is not None:
            for idx, (phase, _) in enumerate(ordered_phases):
                phase_seq = phase.SequenceNumber if hasattr(phase, "SequenceNumber") else idx
                if phase.Id == demolished_phase_id:
                    end_idx = idx  # Element is demolished at this phase, so it doesn't exist in this phase
                    break
                elif phase_seq >= demolished_phase_seq:
                    # Phase is at or after demolition, element doesn't exist here
                    end_idx = idx
                    break
        
        return (start_idx, end_idx)
    except Exception as e:
        logger.debug("Error getting element phase range: {}".format(str(e)))
        # Fallback: assume element exists in all phases
        return (0, len(ordered_phases))

def build_rooms_by_phase_and_level(rooms):
    """Build a spatial index of rooms by phase and level.
    
    Args:
        rooms: List of Room elements
        
    Returns:
        dict: {phase_id_int: {level_id: [rooms...]}}
    """
    rooms_by_phase_by_level = defaultdict(lambda: defaultdict(list))
    rooms_indexed = 0
    rooms_skipped_no_phase = 0
    
    for room in rooms:
        if not isinstance(room, Room):
            continue
        
        # Get room phase
        phase_id = get_room_phase_id(room)
        if not phase_id:
            # Skip rooms without phase (shouldn't happen, but be safe)
            rooms_skipped_no_phase += 1
            continue
        
        phase_id_int = phase_id.IntegerValue
        
        # Get room level
        if hasattr(room, "LevelId") and room.LevelId:
            level_id = room.LevelId
            rooms_by_phase_by_level[phase_id_int][level_id].append(room)
            rooms_indexed += 1
    return rooms_by_phase_by_level

def element_exists_in_phase(element, phase_id):
    """Check if an element exists in a specific phase.
    
    An element exists in a phase if:
    - CreatedPhaseId <= phase AND
    - (DemolishedPhaseId is invalid OR phase < DemolishedPhaseId)
    
    Args:
        element: Element to check
        phase_id: ElementId of the phase to check
        
    Returns:
        bool: True if element exists in the phase, False otherwise
    """
    try:
        # Get CreatedPhaseId
        created_phase_id = None
        if hasattr(element, "CreatedPhaseId") and element.CreatedPhaseId:
            if element.CreatedPhaseId != ElementId.InvalidElementId:
                created_phase_id = element.CreatedPhaseId
        elif hasattr(element, "get_Parameter"):
            created_param = element.get_Parameter(BuiltInParameter.PHASE_CREATED)
            if created_param and created_param.HasValue:
                created_phase_id = created_param.AsElementId()
                if not created_phase_id or created_phase_id == ElementId.InvalidElementId:
                    created_phase_id = None
        
        # Get DemolishedPhaseId
        demolished_phase_id = None
        if hasattr(element, "DemolishedPhaseId") and element.DemolishedPhaseId:
            if element.DemolishedPhaseId != ElementId.InvalidElementId:
                demolished_phase_id = element.DemolishedPhaseId
        elif hasattr(element, "get_Parameter"):
            demolished_param = element.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED)
            if demolished_param and demolished_param.HasValue:
                demolished_phase_id = demolished_param.AsElementId()
                if not demolished_phase_id or demolished_phase_id == ElementId.InvalidElementId:
                    demolished_phase_id = None
        
        # If no phase data, assume element exists in all phases
        if not created_phase_id:
            return True
        
        # Element must be created in or before this phase
        if created_phase_id.IntegerValue > phase_id.IntegerValue:
            return False
        
        # Element must not be demolished in or before this phase
        if demolished_phase_id:
            # If demolished phase <= current phase, element doesn't exist
            if demolished_phase_id.IntegerValue <= phase_id.IntegerValue:
                return False
        
        return True
    except Exception as e:
        logger.debug("Error checking element existence in phase: {}".format(str(e)))
        return True  # Fallback: assume exists

def get_containing_room_phase_aware(element, doc, rooms_by_phase_by_level, ordered_phases, element_phases_for_checking=None, link_instance=None, host_doc=None, phase_map=None):
    """Find the room containing an element, considering phase relationships.
    
    Iterates through phases where the element exists, finding the latest phase
    where containment is found. Does NOT clear if a later phase has no room.
    Skips phases where the element is demolished.
    
    Args:
        element: Target element
        doc: Revit document (source document where rooms are)
        rooms_by_phase_by_level: Dict from build_rooms_by_phase_and_level()
        ordered_phases: List of (phase, order_index) tuples from get_ordered_phases() for source doc
        element_phases_for_checking: Optional list of (phase, order_index) tuples for element's document
                                    (needed when element is in different doc than rooms)
        link_instance: Optional RevitLinkInstance for coordinate transformation
        host_doc: Optional host document (needed for phase map lookup)
        phase_map: Optional phase map from RevitLinkType.GetPhaseMap() - maps host phase IDs to linked phase IDs
        
    Returns:
        Room: Containing room from latest applicable phase, or None
    """
    try:
        # Get element's document for in-place family handling
        element_doc = element.Document if hasattr(element, 'Document') else doc
        
        # Get multiple test points for better detection
        # Pass doc for in-place family handling (uses geometry extraction)
        test_points = get_element_test_points(element, element_doc)
        if not test_points:
            return None
        
        # CRITICAL: Transform points if element is in different document than rooms
        # When element is in main doc and rooms are in linked doc, coordinates need transformation
        if element_doc and element_doc != doc and link_instance:
            try:
                # Get the transform from the link instance
                link_transform = link_instance.GetTotalTransform()
                
                # Check if transform is identity (no transformation needed)
                is_identity = (link_transform.Origin.X == 0 and link_transform.Origin.Y == 0 and link_transform.Origin.Z == 0 and
                              link_transform.BasisX.X == 1 and link_transform.BasisX.Y == 0 and link_transform.BasisX.Z == 0 and
                              link_transform.BasisY.X == 0 and link_transform.BasisY.Y == 1 and link_transform.BasisY.Z == 0 and
                              link_transform.BasisZ.X == 0 and link_transform.BasisZ.Y == 0 and link_transform.BasisZ.Z == 1)
                
                # Transform test points from element doc to room doc coordinate system
                # Use Inverse transform: points in host -> points in linked doc
                transformed_points = []
                for point in test_points:
                    if is_identity:
                        transformed_point = point  # No transformation needed
                    else:
                        transformed_point = link_transform.Inverse.OfPoint(point)
                    transformed_points.append(transformed_point)
                test_points = transformed_points
            except Exception as e:
                pass  # Continue with original points if transform fails
        
        # Get element phase range
        # Use element_phases_for_checking if provided (for cross-document phase matching)
        # Otherwise use ordered_phases (same document case)
        phases_for_element_check = element_phases_for_checking if element_phases_for_checking is not None else ordered_phases
        elem_start_idx, elem_end_idx = get_element_phase_range(element, phases_for_element_check)
        
        # Map element phase indices to source phase indices
        if element_phases_for_checking is not None and elem_start_idx < elem_end_idx:
            # Create mapping: element phase index -> source phase index
            element_phase_map = {}
            
            # PREFERRED: Use Revit's phase map from RevitLinkType.GetPhaseMap() if available
            # This respects user-configured phase mappings set in Revit UI
            if phase_map:
                # Build lookup from source phase ElementId to index
                src_phase_id_to_idx = {}
                for src_idx, (src_phase, _) in enumerate(ordered_phases):
                    src_phase_id_to_idx[src_phase.Id.IntegerValue] = src_idx
                
                # Map each element phase to source phase using Revit's phase map
                for elem_idx, (elem_phase, _) in enumerate(element_phases_for_checking):
                    host_phase_id = elem_phase.Id
                    # phase_map maps host phase ID -> linked (source) phase ID
                    if host_phase_id in phase_map:
                        linked_phase_id = phase_map[host_phase_id]
                        linked_phase_int = linked_phase_id.IntegerValue
                        if linked_phase_int in src_phase_id_to_idx:
                            element_phase_map[elem_idx] = src_phase_id_to_idx[linked_phase_int]
                
                if element_phase_map:
                    logger.debug("Phase mapping via RevitLinkType.GetPhaseMap(): {} mappings".format(len(element_phase_map)))
            
            # FALLBACK: Try matching by name if phase_map is not available or didn't produce mappings
            if not element_phase_map:
                for elem_idx, (elem_phase, _) in enumerate(element_phases_for_checking):
                    # First try matching by name
                    matched = False
                    for src_idx, (src_phase, _) in enumerate(ordered_phases):
                        if elem_phase.Name == src_phase.Name:
                            element_phase_map[elem_idx] = src_idx
                            matched = True
                            break
                    
                    # If no name match, try matching by sequence number
                    if not matched:
                        try:
                            elem_seq = elem_phase.SequenceNumber if hasattr(elem_phase, 'SequenceNumber') else None
                            if elem_seq is not None:
                                for src_idx, (src_phase, _) in enumerate(ordered_phases):
                                    src_seq = src_phase.SequenceNumber if hasattr(src_phase, 'SequenceNumber') else None
                                    if src_seq == elem_seq:
                                        element_phase_map[elem_idx] = src_idx
                                        matched = True
                                        break
                        except:
                            pass
                
                if element_phase_map:
                    logger.debug("Phase mapping via name/sequence fallback: {} mappings".format(len(element_phase_map)))
            
            # Adjust indices to source phase space
            is_fallback_mode = False
            if element_phase_map:
                mapped_indices = [element_phase_map.get(i) for i in range(elem_start_idx, elem_end_idx) if i in element_phase_map]
                if mapped_indices:
                    mapped_start = min(mapped_indices)
                    mapped_end = max(mapped_indices) + 1
                    start_idx, end_idx = mapped_start, mapped_end
                else:
                    # No valid mapping found for element's phase range
                    # FALLBACK: Check all source phases (element exists but phases don't match)
                    start_idx, end_idx = 0, len(ordered_phases)
                    is_fallback_mode = True
                    logger.debug("Phase mapping fallback: checking all {} source phases".format(len(ordered_phases)))
            else:
                # No mapping possible - phases don't match at all
                # FALLBACK: Check all source phases (element exists but phases don't match)
                start_idx, end_idx = 0, len(ordered_phases)
                is_fallback_mode = True
                logger.debug("No phase mapping possible: checking all {} source phases".format(len(ordered_phases)))
        else:
            # Same document case - use indices directly
            start_idx, end_idx = elem_start_idx, elem_end_idx
            is_fallback_mode = False
        
        if start_idx >= end_idx:
            # Element doesn't exist in any phase
            return None
        
        # Get element level for optimization
        element_level_id = None
        if hasattr(element, "LevelId"):
            element_level_id = element.LevelId
        elif hasattr(element, "get_Parameter"):
            level_param = element.get_Parameter("Level")
            if level_param:
                element_level_id = level_param.AsElementId()
        
        # Get element bounding box for pre-filtering
        element_bbox = element.get_BoundingBox(None)
        
        # Track all matching rooms (will select by lowest ElementId)
        matching_rooms = []
        
        # Iterate through phases where element exists (ascending order)
        for phase_idx in range(start_idx, end_idx):
            phase, _ = ordered_phases[phase_idx]
            phase_id_int = phase.Id.IntegerValue
            
            # CRITICAL: Check if element exists in this specific phase
            # Skip if element is demolished in this phase
            # NOTE: When using fallback mode (checking all source phases), we can't reliably check
            # element phase existence because phase IDs are document-specific. So we skip this check
            # in fallback mode and check all phases anyway.
            if not is_fallback_mode and not element_exists_in_phase(element, phase.Id):
                continue  # Element is demolished in this phase, skip it
            
            
            # Get rooms for this phase
            phase_rooms_by_level = rooms_by_phase_by_level.get(phase_id_int, {})
            if not phase_rooms_by_level:
                continue  # No rooms in this phase
            
            # Check rooms on same level first (optimization)
            rooms_to_check = []
            if element_level_id:
                level_rooms = phase_rooms_by_level.get(element_level_id, [])
                rooms_to_check.extend(level_rooms)
            
            # Pre-filter by bounding box if we have many rooms
            if element_bbox and len(rooms_to_check) > 50:
                expanded_min = XYZ(element_bbox.Min.X - 0.5, element_bbox.Min.Y - 0.5, element_bbox.Min.Z - 0.5)
                expanded_max = XYZ(element_bbox.Max.X + 0.5, element_bbox.Max.Y + 0.5, element_bbox.Max.Z + 0.5)
                
                def bboxes_overlap(bbox1_min, bbox1_max, bbox2_min, bbox2_max):
                    """Check if two bounding boxes overlap."""
                    return (bbox1_min.X <= bbox2_max.X and bbox1_max.X >= bbox2_min.X and
                            bbox1_min.Y <= bbox2_max.Y and bbox1_max.Y >= bbox2_min.Y and
                            bbox1_min.Z <= bbox2_max.Z and bbox1_max.Z >= bbox2_min.Z)
                
                filtered_rooms = []
                for room in rooms_to_check:
                    room_bbox = room.get_BoundingBox(None)
                    if room_bbox and bboxes_overlap(expanded_min, expanded_max, room_bbox.Min, room_bbox.Max):
                        filtered_rooms.append(room)
                
                if len(filtered_rooms) < len(rooms_to_check) * 0.8:
                    rooms_to_check = filtered_rooms
            
            # Check containment in this phase
            for point in test_points:
                for room in rooms_to_check:
                    if is_point_in_room(room, point):
                        # Add to matching rooms if not already added
                        if room not in matching_rooms:
                            matching_rooms.append(room)
                        break  # Found containment for this point, move to next point
                # Continue checking all points even if we found a match
            
            # If not found on same level, check other levels (fallback)
            if not matching_rooms:
                # Collect rooms from other levels
                other_level_rooms = []
                for level_id, level_rooms in phase_rooms_by_level.items():
                    if element_level_id and level_id == element_level_id:
                        continue  # Already checked
                    other_level_rooms.extend(level_rooms)
                
                # Pre-filter by bounding box
                if element_bbox and other_level_rooms:
                    expanded_min = XYZ(element_bbox.Min.X - 0.5, element_bbox.Min.Y - 0.5, element_bbox.Min.Z - 0.5)
                    expanded_max = XYZ(element_bbox.Max.X + 0.5, element_bbox.Max.Y + 0.5, element_bbox.Max.Z + 0.5)
                    
                    def bboxes_overlap(bbox1_min, bbox1_max, bbox2_min, bbox2_max):
                        """Check if two bounding boxes overlap."""
                        return (bbox1_min.X <= bbox2_max.X and bbox1_max.X >= bbox2_min.X and
                                bbox1_min.Y <= bbox2_max.Y and bbox1_max.Y >= bbox2_min.Y and
                                bbox1_min.Z <= bbox2_max.Z and bbox1_max.Z >= bbox2_min.Z)
                    
                    filtered_other = []
                    for room in other_level_rooms:
                        room_bbox = room.get_BoundingBox(None)
                        if room_bbox and bboxes_overlap(expanded_min, expanded_max, room_bbox.Min, room_bbox.Max):
                            filtered_other.append(room)
                    
                    other_level_rooms = filtered_other
                
                # Check containment in other levels
                for point in test_points:
                    for room in other_level_rooms:
                        if is_point_in_room(room, point):
                            # Add to matching rooms if not already added
                            if room not in matching_rooms:
                                matching_rooms.append(room)
                            break  # Found containment for this point, move to next point
        
        # Return room with lowest ElementId if multiple matches found
        if matching_rooms:
            # Sort by ElementId (lowest first) and return the first one
            matching_rooms.sort(key=lambda r: r.Id.IntegerValue)
            return matching_rooms[0]
        
        return None
    except Exception as e:
        logger.debug("Error getting containing room (phase-aware): {}".format(str(e)))
        return None

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
        # Pass doc for in-place family handling (uses geometry extraction)
        test_points = get_element_test_points(element, doc)
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
        # OPTIMIZATION: Pre-filter rooms by bounding box even for level-based search
        element_bbox = element.get_BoundingBox(None)
        if element_level_id:
            level_rooms = rooms_by_level.get(element_level_id, [])
            
            # Pre-filter level rooms by bounding box to reduce checks
            if element_bbox and len(level_rooms) > 50:  # Only filter if many rooms on level
                expanded_min = XYZ(element_bbox.Min.X - 0.5, element_bbox.Min.Y - 0.5, element_bbox.Min.Z - 0.5)
                expanded_max = XYZ(element_bbox.Max.X + 0.5, element_bbox.Max.Y + 0.5, element_bbox.Max.Z + 0.5)
                
                def bboxes_overlap(bbox1_min, bbox1_max, bbox2_min, bbox2_max):
                    """Check if two bounding boxes overlap."""
                    return (bbox1_min.X <= bbox2_max.X and bbox1_max.X >= bbox2_min.X and
                            bbox1_min.Y <= bbox2_max.Y and bbox1_max.Y >= bbox2_min.Y and
                            bbox1_min.Z <= bbox2_max.Z and bbox1_max.Z >= bbox2_min.Z)
                
                filtered_level_rooms = []
                for room in level_rooms:
                    room_bbox = room.get_BoundingBox(None)
                    if room_bbox and bboxes_overlap(expanded_min, expanded_max, room_bbox.Min, room_bbox.Max):
                        filtered_level_rooms.append(room)
                
                # Use filtered rooms if we filtered significantly
                if len(filtered_level_rooms) < len(level_rooms) * 0.8:  # If filtered to <80% of original
                    level_rooms = filtered_level_rooms
            
            for point in test_points:
                for room in level_rooms:
                    if is_point_in_room(room, point):
                        return room  # Early exit - found containment
        
        # Fallback: only check if level-based search didn't find containment
        # OPTIMIZATION: Use bounding box pre-filtering to reduce rooms to check
        # Get element bounding box to filter rooms spatially
        element_bbox = element.get_BoundingBox(None)
        if element_bbox:
            # Use tighter bounding box expansion (0.5 feet) for better filtering
            # Only expand enough to catch rooms that might contain the element
            expanded_min = XYZ(element_bbox.Min.X - 0.5, element_bbox.Min.Y - 0.5, element_bbox.Min.Z - 0.5)
            expanded_max = XYZ(element_bbox.Max.X + 0.5, element_bbox.Max.Y + 0.5, element_bbox.Max.Z + 0.5)
            
            # Pre-filter rooms by bounding box intersection (much faster than checking all rooms)
            # Check if bounding boxes overlap manually
            def bboxes_overlap(bbox1_min, bbox1_max, bbox2_min, bbox2_max):
                """Check if two bounding boxes overlap."""
                return (bbox1_min.X <= bbox2_max.X and bbox1_max.X >= bbox2_min.X and
                        bbox1_min.Y <= bbox2_max.Y and bbox1_max.Y >= bbox2_min.Y and
                        bbox1_min.Z <= bbox2_max.Z and bbox1_max.Z >= bbox2_min.Z)
            
            filtered_rooms = []
            total_rooms = sum(len(room_list) for room_list in rooms_by_level.values())
            # Limit fallback to max 200 rooms to prevent expensive searches
            MAX_FALLBACK_ROOMS = 200
            for room_list in rooms_by_level.values():
                for room in room_list:
                    # Skip rooms already checked on the same level
                    if element_level_id and room.LevelId == element_level_id:
                        continue
                    room_bbox = room.get_BoundingBox(None)
                    if room_bbox and bboxes_overlap(expanded_min, expanded_max, room_bbox.Min, room_bbox.Max):
                        filtered_rooms.append(room)
                        # Early exit if we've found enough candidates
                        if len(filtered_rooms) >= MAX_FALLBACK_ROOMS:
                            break
                if len(filtered_rooms) >= MAX_FALLBACK_ROOMS:
                    break
            
            # Only check pre-filtered rooms (typically much fewer than all rooms)
            if filtered_rooms:
                for point in test_points:
                    for room in filtered_rooms:
                        if is_point_in_room(room, point):
                            return room  # Early exit - found containment
        else:
            # If no bounding box, fall back to checking all rooms (rare case)
            total_rooms = sum(len(room_list) for room_list in rooms_by_level.values())
            for point in test_points:
                for room_list in rooms_by_level.values():
                    for room in room_list:
                        if is_point_in_room(room, point):
                            return room  # Early exit - found containment
        
        return None
    except Exception as e:
        logger.debug("Error getting containing room: {}".format(str(e)))
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
        # Pass doc for in-place family handling (uses geometry extraction)
        test_points = get_element_test_points(element, doc)
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
        # OPTIMIZATION: Use bounding box pre-filtering to reduce spaces to check
        element_bbox = element.get_BoundingBox(None)
        if element_bbox:
            # Expand bounding box slightly (1 foot in Revit units) to account for edge cases
            expanded_min = XYZ(element_bbox.Min.X - 1.0, element_bbox.Min.Y - 1.0, element_bbox.Min.Z - 1.0)
            expanded_max = XYZ(element_bbox.Max.X + 1.0, element_bbox.Max.Y + 1.0, element_bbox.Max.Z + 1.0)
            
            # Pre-filter spaces by bounding box intersection
            def bboxes_overlap(bbox1_min, bbox1_max, bbox2_min, bbox2_max):
                """Check if two bounding boxes overlap."""
                return (bbox1_min.X <= bbox2_max.X and bbox1_max.X >= bbox2_min.X and
                        bbox1_min.Y <= bbox2_max.Y and bbox1_max.Y >= bbox2_min.Y and
                        bbox1_min.Z <= bbox2_max.Z and bbox1_max.Z >= bbox2_min.Z)
            
            filtered_spaces = []
            for space_list in spaces_by_level.values():
                for space in space_list:
                    space_bbox = space.get_BoundingBox(None)
                    if space_bbox and bboxes_overlap(expanded_min, expanded_max, space_bbox.Min, space_bbox.Max):
                        filtered_spaces.append(space)
            
            # Only check pre-filtered spaces
            if filtered_spaces:
                for point in test_points:
                    for space in filtered_spaces:
                        if is_point_in_space(space, point):
                            return space  # Early exit - found containment
        else:
            # If no bounding box, fall back to checking all spaces (rare case)
            for point in test_points:
                for space_list in spaces_by_level.values():
                    for space in space_list:
                        if is_point_in_space(space, point):
                            return space  # Early exit - found containment
        
        return None
    except Exception as e:
        logger.debug("Error getting containing space: {}".format(str(e)))
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
        # Pass doc for in-place family handling (uses geometry extraction)
        test_points = get_element_test_points(element, doc)
        if not test_points:
            # Reduced logging - only log if debug mode is very verbose
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
            # Removed excessive debug logging - only log summary at end if needed
            
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
                        # Removed excessive debug logging - found containment, return immediately
                        return area  # Early exit - found containment
        
        # Fallback: check all areas with all test points (with early exit)
        if not element_level_id or areas_checked == 0:
            total_areas = sum(len(area_list) for area_list in areas_by_level.values())
            # Removed excessive debug logging
            
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
                            # Removed excessive debug logging - found containment, return immediately
                            return area  # Early exit - found containment
        
        # Removed excessive debug logging - only log errors, not every check
        return None
    except Exception as e:
        logger.debug("Error getting containing area: {}".format(str(e)))
        return None

def build_source_element_spatial_index(source_elements, doc, cell_size_feet=50.0, sort_property="ElementId", sort_descending=False):
    """Build a spatial hash index for source elements to enable fast containment queries.
    
    Creates a 2D grid (XY plane) where each cell contains a list of source elements
    whose bounding boxes overlap that cell. Elements are sorted by the specified property
    (default: ElementId) for deterministic containment selection.
    
    Args:
        source_elements: List of source zone elements (Mass/Generic Model)
        doc: Revit document
        cell_size_feet: Size of each grid cell in Revit internal units (feet)
        sort_property: Property name to sort by (default: "ElementId")
        sort_descending: If True, sort in descending order (default: False)
        
    Returns:
        dict: Spatial index mapping (ix, iy) -> sorted list of source elements
    """
    spatial_index = {}
    
    if not source_elements:
        return spatial_index
    
    # Ensure geometries are precomputed (should already be done, but safe check)
    precompute_geometries(source_elements, doc)
    
    # Sort elements by specified property (elements should already be sorted, but ensure consistency)
    # If sort_property is ElementId, use simple integer value sort
    if sort_property == "ElementId":
        sorted_elements = sorted(source_elements, key=lambda el: el.Id.IntegerValue, reverse=sort_descending)
    else:
        # Import sort function from core
        try:
            from zone3d.core import sort_source_elements
            sorted_elements = sort_source_elements(source_elements, sort_property, descending=sort_descending)
        except ImportError:
            # Fallback to ElementId sorting if import fails
            sorted_elements = sorted(source_elements, key=lambda el: el.Id.IntegerValue, reverse=sort_descending)
    
    for element in sorted_elements:
        element_id = element.Id.IntegerValue
        
        # Get cached bounding box
        bbox = None
        if element_id in _geometry_cache:
            cached = _geometry_cache[element_id]
            bbox = cached.get("bbox")
        
        # Fallback to direct bbox if not cached
        if not bbox:
            bbox = element.get_BoundingBox(None)
        
        if not bbox:
            continue  # Skip elements without bounding boxes
        
        # Calculate grid cell range for this element's bounding box
        min_x, min_y = bbox.Min.X, bbox.Min.Y
        max_x, max_y = bbox.Max.X, bbox.Max.Y
        
        # Convert to grid coordinates
        min_ix = int(min_x / cell_size_feet)
        min_iy = int(min_y / cell_size_feet)
        max_ix = int(max_x / cell_size_feet)
        max_iy = int(max_y / cell_size_feet)
        
        # Add element to all overlapping cells
        for ix in range(min_ix, max_ix + 1):
            for iy in range(min_iy, max_iy + 1):
                cell_key = (ix, iy)
                if cell_key not in spatial_index:
                    spatial_index[cell_key] = []
                spatial_index[cell_key].append(element)
    
    logger.debug("[DEBUG] Built spatial index with {} cells for {} source elements".format(
        len(spatial_index), len(source_elements)))
    
    return spatial_index

def get_containing_element_indexed(target_el, doc, element_index, cell_size_feet=50.0, sort_property="ElementId", sort_descending=False):
    """Find containing element using pre-built spatial index (fast path).
    
    Uses spatial hash lookup instead of database queries. Checks a 3x3 cell
    neighborhood around the target point to handle boundary cases.
    Uses multiple test points for elements with vertical extent (columns).
    
    Args:
        target_el: Target element
        doc: Revit document
        element_index: Spatial index dict from build_source_element_spatial_index
        cell_size_feet: Size of each grid cell (must match index cell size)
        
    Returns:
        Element: First containing element matching user's sort order, or None
    """
    try:
        # Get multiple test points for better detection (columns get vertical points)
        # Pass doc for in-place family handling (uses geometry extraction)
        test_points = get_element_test_points(target_el, doc)
        if not test_points:
            return None
        
        # Use first point for spatial index lookup (X/Y position)
        primary_point = test_points[0]
        
        # Calculate grid cell for target point
        ix = int(primary_point.X / cell_size_feet)
        iy = int(primary_point.Y / cell_size_feet)
        
        # Check 3x3 neighborhood to handle boundary cases
        # Collect all candidates from all cells, allowing duplicates
        all_candidates = []
        for di in [-1, 0, 1]:
            for dj in [-1, 0, 1]:
                cell_key = (ix + di, iy + dj)
                if cell_key in element_index:
                    all_candidates.extend(element_index[cell_key])
        
        if not all_candidates:
            return None
        
        # Deduplicate candidates
        seen_ids = set()
        candidates_dict = {}
        for source_el in all_candidates:
            el_id = source_el.Id.IntegerValue
            if el_id not in seen_ids:
                seen_ids.add(el_id)
                candidates_dict[el_id] = source_el
        
        # Convert to list and sort by sort_property to preserve user's configured sort order
        candidates = list(candidates_dict.values())
        if sort_property == "ElementId":
            candidates.sort(key=lambda el: el.Id.IntegerValue, reverse=sort_descending)
        else:
            # Use sort_source_elements function to maintain consistency with spatial index sorting
            try:
                from zone3d.core import sort_source_elements
                candidates = sort_source_elements(candidates, sort_property, descending=sort_descending)
            except ImportError:
                # Fallback to ElementId sorting if import fails
                candidates.sort(key=lambda el: el.Id.IntegerValue, reverse=sort_descending)
        
        # Check candidates in user's configured sort order (preserved from spatial index)
        # Elements are already sorted in the index, and deduplication preserves sort order
        for source_el in candidates:
            # Check each test point against this candidate
            for point in test_points:
                # Fast bounding box rejection
                try:
                    element_id = source_el.Id.IntegerValue
                    if element_id in _geometry_cache:
                        cached = _geometry_cache[element_id]
                        bbox = cached.get("bbox")
                        if bbox and not is_point_in_bbox(point, bbox):
                            continue  # This point not in bbox, try next point
                except Exception as e:
                    continue
                
                # Final point-in-solid check
                if is_point_in_element(source_el, point, doc):
                    return source_el  # Found containment - return first match respecting user's sort order
        
        return None
    except Exception as e:
        logger.debug("Error getting containing element (indexed): {}".format(str(e)))
        return None

def get_containing_element(element, doc, source_categories):
    """Find the containing element (Mass/Generic Model) for an element.
    
    Fallback method using Quick Filter pre-filtering with BoundingBoxContainsPointFilter
    (learnrevitapi best practice - filters at database level before expanding elements).
    
    FIXED: Now uses ElementMulticategoryFilter for multiple categories (OR logic)
    instead of chaining OfCategory() which creates AND logic.
    
    Args:
        element: Target element
        doc: Revit document
        source_categories: List of BuiltInCategory values to search
        
    Returns:
        Element: Containing element or None
    """
    try:
        # Pass doc for in-place family handling (uses geometry extraction)
        point = get_element_representative_point(element, doc)
        if not point:
            return None
        
        # CRITICAL OPTIMIZATION: Use Quick Filter to pre-filter candidates
        # BoundingBoxContainsPointFilter operates on ElementRecord (database level)
        # This is 10-100x faster than expanding all elements and checking geometry
        point_filter = BoundingBoxContainsPointFilter(point)
        
        # Build collector with category filters
        # FIXED: Collect from each category separately and combine (OR logic)
        # Chaining OfCategory() creates AND logic which is wrong
        # Use the same pattern as target element collection (per-category + combine)
        candidate_elements = []
        element_ids = set()  # Track IDs to avoid duplicates
        
        if source_categories:
            for category in source_categories:
                category_elements = FilteredElementCollector(doc)\
                    .WhereElementIsNotElementType()\
                    .OfCategory(category)\
                    .WherePasses(point_filter)\
                    .ToElements()
                
                for el in category_elements:
                    el_id = el.Id.IntegerValue
                    if el_id not in element_ids:
                        element_ids.add(el_id)
                        candidate_elements.append(el)
        else:
            # No categories specified - get all elements (rare case)
            candidate_elements = FilteredElementCollector(doc)\
                .WhereElementIsNotElementType()\
                .WherePasses(point_filter)\
                .ToElements()
        
        # Sort candidates by ElementId for deterministic selection (lowest wins)
        candidate_elements = sorted(candidate_elements, key=lambda el: el.Id.IntegerValue)
        
        # Now check geometry only on pre-filtered candidates (much fewer elements)
        for source_el in candidate_elements:
            if is_point_in_element(source_el, point, doc):
                return source_el  # Early exit - found containment (lowest ElementId)
        
        return None
    except Exception as e:
        logger.debug("Error getting containing element: {}".format(str(e)))
        return None

def get_containing_element_by_strategy(element, doc, strategy, source_categories=None, 
                                     rooms_by_level=None, spaces_by_level=None, areas_by_level=None,
                                     element_index=None, element_index_cell_size=50.0, 
                                     sort_property="ElementId", sort_descending=False):
    """Unified function that routes to appropriate containment method.
    
    Args:
        element: Target element
        doc: Revit document
        strategy: Containment strategy ("room", "space", "area", "element")
        source_categories: List of BuiltInCategory values (for element strategy)
        rooms_by_level: Optional pre-grouped rooms dict
        spaces_by_level: Optional pre-grouped spaces dict
        areas_by_level: Optional pre-grouped areas dict
        element_index: Optional spatial index dict for element strategy (fast path)
        element_index_cell_size: Cell size for spatial index (feet, default 50.0)
        
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
        # Use indexed lookup if index is provided (fast path)
        if element_index is not None:
            return get_containing_element_indexed(element, doc, element_index, element_index_cell_size, sort_property=sort_property, sort_descending=sort_descending)
        else:
            # Fallback to database query method
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
    
    logger.debug("[DEBUG] Precomputing geometries for {} elements".format(len(elements)))
    
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
            logger.debug("Error precomputing geometry for element {}: {}".format(element_id, str(e)))
            continue
    
    if area_count > 0:
        logger.debug("[DEBUG] Area geometry precomputation: {} total, {} succeeded, {} failed".format(
            area_count, area_success_count, area_fail_count))

