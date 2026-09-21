# -*- coding: utf-8 -*-
"""Place and keep up to date the '3D View Reference' family for views.

One family instance marks where a view cuts the model and how far it extends.
Each instance stores the UniqueId of its view, so running the tools again
updates the existing instance instead of adding another one.

Geometry (checked in Revit 2026 on sections, a 30 degree skewed section, an
elevation, a detail callout and plan callouts): a view's ``CropBox.Transform``
has BasisX/Y/Z equal to the view's right, up and view direction, and local
``Max.Z`` is the cut plane of a section-like view. A work plane built from
(right, up) therefore gives an instance whose X/Y axes are the view's own,
with no per-direction special cases. Plans take their cut plane height from
the view range instead: a ceiling plan callout's crop box does not sit on it.

The family extrudes from its work plane towards the side its "View Name" text
faces, which is the viewer's side. To show the view depth the work plane is
therefore put on the far clip plane, so the box ends at the cut plane and the
text still faces the viewer.
"""
from Autodesk.Revit.DB import (
    BuiltInParameter,
    Element,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
    Level,
    Plane,
    PlanViewPlane,
    SketchPlane,
    View,
    ViewPlan,
    ViewSheet,
    ViewType,
    XYZ,
)
from Autodesk.Revit.DB.ExtensibleStorage import ExtensibleStorageFilter
from Autodesk.Revit.DB.Structure import StructuralType

from pyrevit import script

from extensible_storage import BaseSchema, simple_field
from revit.compat import get_element_id_value

logger = script.get_logger()

FAMILY_NAME = "3D View Reference"
PREFERRED_TYPE_NAME = "Standard Reference"

PARAM_WIDTH = "View Width"
PARAM_HEIGHT = "View Height"
PARAM_DEPTH = "View Depth"
PARAM_NAME = "View Name"

# Thickness of the plate when the view depth is not shown (the family default)
PLATE_THICKNESS = 10 / 304.8

# View kinds offered by the tools: (key, label shown in the UI)
KIND_SECTION = "section"
KIND_ELEVATION = "elevation"
KIND_DETAIL = "detail"
KIND_PLAN_CALLOUT = "plan_callout"
VIEW_KINDS = [
    (KIND_SECTION, "Sections"),
    (KIND_ELEVATION, "Elevations"),
    (KIND_DETAIL, "Detail Views"),
    (KIND_PLAN_CALLOUT, "Plan Callouts"),
]

_PLAN_TYPES = (
    ViewType.FloorPlan,
    ViewType.CeilingPlan,
    ViewType.EngineeringPlan,
    ViewType.AreaPlan,
)


class ViewReferenceSchema(BaseSchema):
    """Links a 3D View Reference instance to the view it represents."""

    guid = "30522bb7-5204-4224-a3f5-9e96d014006b"

    @simple_field(value_type="string")
    def view_unique_id():
        """UniqueId of the view this instance represents."""
        return None


class ViewFrame(object):
    """Where a view sits in model space: cut-plane centre, axes and size."""

    def __init__(self, origin, right, up, width, height, depth):
        self.origin = origin
        self.right = right
        self.up = up
        self.width = width
        self.height = height
        self.depth = depth  # None when the view has no usable far limit

    def placement_origin(self, show_depth):
        """Work plane origin: the cut plane, or the far clip when showing depth."""
        if show_depth and self.depth is not None:
            view_direction = self.right.CrossProduct(self.up)
            return self.origin - view_direction.Multiply(self.depth)
        return self.origin

    def thickness(self, show_depth):
        if show_depth and self.depth is not None:
            return self.depth
        return PLATE_THICKNESS


class SyncResult(object):
    """Outcome of sync_view_references."""

    def __init__(self):
        self.created = []   # ElementIds
        self.updated = []   # ElementIds
        self.skipped = []   # (view name, reason)

    @property
    def element_ids(self):
        return self.created + self.updated


def get_view_kind(view):
    """Return the VIEW_KINDS key for a view, or None if it is not supported."""
    if view.IsTemplate:
        return None
    view_type = view.ViewType
    if view_type == ViewType.Section:
        return KIND_SECTION
    if view_type == ViewType.Elevation:
        return KIND_ELEVATION
    if view_type == ViewType.Detail:
        return KIND_DETAIL
    if view_type in _PLAN_TYPES and view.IsCallout:
        return KIND_PLAN_CALLOUT
    return None


def get_view_kind_label(view):
    """Human readable kind for one view, e.g. 'Section (callout)'."""
    label = view.ViewType.ToString()
    if get_view_kind(view) == KIND_PLAN_CALLOUT:
        return "{} callout".format(label)
    if view.IsCallout:
        label += " (callout)"
    return label


def collect_views(doc, kinds=None):
    """All views the tools can place a reference for, optionally by kind."""
    views = []
    for view in FilteredElementCollector(doc).OfClass(View):
        kind = get_view_kind(view)
        if kind is None:
            continue
        if kinds is not None and kind not in kinds:
            continue
        views.append(view)
    return views


def build_sheet_lookup(doc):
    """Map view id value -> 'number - name' of the sheet the view is placed on."""
    lookup = {}
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        label = "{} - {}".format(sheet.SheetNumber, sheet.Name)
        for view_id in sheet.GetAllPlacedViews():
            lookup[get_element_id_value(view_id)] = label
    return lookup


def find_family_symbol(doc):
    """The 3D View Reference type to place, or None if the family is not loaded."""
    first = None
    for symbol in FilteredElementCollector(doc).OfClass(FamilySymbol):
        family = symbol.Family
        if family is None or family.Name != FAMILY_NAME:
            continue
        if Element.Name.__get__(symbol) == PREFERRED_TYPE_NAME:
            return symbol
        if first is None:
            first = symbol
    return first


def get_view_frame(view):
    """Return the ViewFrame of a view, or None if it has no usable crop box."""
    crop = view.CropBox
    if crop is None:
        return None
    width = crop.Max.X - crop.Min.X
    height = crop.Max.Y - crop.Min.Y
    if width <= 0 or height <= 0:
        return None

    transform = crop.Transform
    centre_on_cut_plane = XYZ(
        (crop.Min.X + crop.Max.X) / 2.0,
        (crop.Min.Y + crop.Max.Y) / 2.0,
        crop.Max.Z,
    )
    origin = transform.OfPoint(centre_on_cut_plane)
    if isinstance(view, ViewPlan):
        cut_z = _get_plan_cut_elevation(view)
        if cut_z is not None:
            origin = XYZ(origin.X, origin.Y, cut_z)
    depth = _get_view_depth(view, crop, origin)
    return ViewFrame(origin, transform.BasisX, transform.BasisY, width, height, depth)


def _get_plane_elevation(view, plane):
    view_range = view.GetViewRange()
    level = view.Document.GetElement(view_range.GetLevelId(plane))
    if isinstance(level, Level):
        return level.ProjectElevation + view_range.GetOffset(plane)
    return None


def _get_plan_cut_elevation(view):
    try:
        return _get_plane_elevation(view, PlanViewPlane.CutPlane)
    except Exception as ex:
        logger.debug("No cut plane for plan '{}': {}".format(view.Name, ex))
        return None


def _get_view_depth(view, crop, origin):
    """Distance from the cut plane to the far limit of the view, or None."""
    if isinstance(view, ViewPlan):
        # A ceiling plan looks up, the family can only show depth downwards
        if view.ViewType == ViewType.CeilingPlan:
            return None
        try:
            far_z = _get_plane_elevation(view, PlanViewPlane.ViewDepthPlane)
            if far_z is not None and origin.Z - far_z > 0:
                return origin.Z - far_z
        except Exception as ex:
            logger.debug("No view depth for plan '{}': {}".format(view.Name, ex))
        return None

    # The crop box keeps a far offset even when far clipping is switched off
    far_clip = view.get_Parameter(BuiltInParameter.VIEWER_BOUND_FAR_CLIPPING)
    if far_clip is not None and far_clip.AsInteger() == 0:
        return None
    depth = crop.Max.Z - crop.Min.Z
    return depth if depth > 0 else None


def find_existing_references(doc):
    """Map view UniqueId -> list of reference instances already in the model."""
    existing = {}
    schema = ViewReferenceSchema.schema
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance).WherePasses(
        ExtensibleStorageFilter(schema.GUID))
    for instance in collector:
        unique_id = ViewReferenceSchema(instance, update=False).get("view_unique_id")
        if unique_id:
            existing.setdefault(unique_id, []).append(instance)
    return existing


def _is_on_frame(instance, frame, show_depth):
    transform = instance.GetTransform()
    return (transform.Origin.IsAlmostEqualTo(frame.placement_origin(show_depth))
            and transform.BasisX.IsAlmostEqualTo(frame.right)
            and transform.BasisY.IsAlmostEqualTo(frame.up))


def _set_parameter(instance, name, value):
    param = instance.LookupParameter(name)
    if param is None:
        logger.warning("Parameter '{}' not found on the {} family".format(name, FAMILY_NAME))
        return
    if param.IsReadOnly:
        logger.warning("Parameter '{}' is read-only".format(name))
        return
    param.Set(value)


def _apply_frame(instance, frame, view, show_depth):
    _set_parameter(instance, PARAM_WIDTH, frame.width)
    _set_parameter(instance, PARAM_HEIGHT, frame.height)
    _set_parameter(instance, PARAM_DEPTH, frame.thickness(show_depth))
    _set_parameter(instance, PARAM_NAME, view.Name)


def _place(doc, symbol, frame, view, show_depth):
    origin = frame.placement_origin(show_depth)
    plane = Plane.CreateByOriginAndBasis(origin, frame.right, frame.up)
    sketch_plane = SketchPlane.Create(doc, plane)
    instance = doc.Create.NewFamilyInstance(
        origin, symbol, sketch_plane, StructuralType.NonStructural)

    link = ViewReferenceSchema(instance, update=False)
    link.set("view_unique_id", view.UniqueId)
    instance.SetEntity(link.unwrap())
    return instance


def sync_view_references(doc, views, symbol, show_depth=False):
    """Create or update one reference per view. Call inside an open transaction.

    With show_depth the reference becomes a box from the cut plane to the
    view's far limit; otherwise it is a thin plate on the cut plane.

    An instance that already sits on the right plane only gets its parameters
    refreshed; one that does not is replaced, because a work plane based
    instance cannot be moved to another plane.
    """
    result = SyncResult()
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()

    existing = find_existing_references(doc)

    for view in views:
        view_name = view.Name
        try:
            frame = get_view_frame(view)
            if frame is None:
                result.skipped.append((view_name, "no crop box"))
                continue

            instances = existing.get(view.UniqueId, [])
            keep = None
            for instance in instances:
                if keep is None and _is_on_frame(instance, frame, show_depth):
                    keep = instance
                else:
                    doc.Delete(instance.Id)

            if keep is not None:
                _apply_frame(keep, frame, view, show_depth)
                result.updated.append(keep.Id)
            else:
                instance = _place(doc, symbol, frame, view, show_depth)
                _apply_frame(instance, frame, view, show_depth)
                if instances:
                    result.updated.append(instance.Id)
                else:
                    result.created.append(instance.Id)
        except Exception as ex:
            logger.error("Could not place a reference for '{}': {}".format(view_name, ex))
            result.skipped.append((view_name, str(ex)))

    return result
