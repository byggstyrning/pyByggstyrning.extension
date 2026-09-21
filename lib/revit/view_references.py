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

Two family files are bundled with the Load Family button. The original is in
Revit 2024 format so every supported Revit can load it. "2026/" holds a copy
with a "Sheet Number" text of its own, built by build_2026_family.py, which
also explains why that text is a nested label family. load_reference_family
picks the newest file the running Revit can open. With a family that has the
Sheet Number parameter the sheet number goes there; with one that does not,
it becomes a second line of the "View Name" text.
"""
import os

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Element,
    FamilyInstance,
    FamilySymbol,
    FilteredElementCollector,
    IFamilyLoadOptions,
    Level,
    Plane,
    PlanViewPlane,
    SketchPlane,
    StorageType,
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
PARAM_SHEET = "Sheet Number"  # only in the 2026 family, see the module docstring

# Shown by the Sheet Number text of a view that is not on a sheet
NO_SHEET_TEXT = "-"

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

    def __init__(self, origin, right, up, width, height, depth, looks_up=False):
        self.origin = origin
        self.right = right
        self.up = up
        self.width = width
        self.height = height
        self.depth = depth  # None when the view has no usable far limit
        self.looks_up = looks_up  # a ceiling plan looks along +view direction
        self.box_depth = None  # how deep the reference is drawn, None = plate

    def set_box_depth(self, show_depth, manual_depth):
        """Pick the drawn depth: a typed value wins, then the view's own depth."""
        if manual_depth is not None:
            self.box_depth = manual_depth
        elif show_depth and self.depth is not None:
            self.box_depth = self.depth
        else:
            self.box_depth = None

    def placement_origin(self):
        """Work plane origin: the cut plane, or the far end of the box.

        The family extrudes towards the viewer, so a box that should reach away
        from the viewer starts at its far end. A ceiling plan looks the other
        way and its box starts on the cut plane.
        """
        if self.box_depth is None or self.looks_up:
            return self.origin
        view_direction = self.right.CrossProduct(self.up)
        return self.origin - view_direction.Multiply(self.box_depth)

    def thickness(self):
        return PLATE_THICKNESS if self.box_depth is None else self.box_depth


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
    """Map view id value -> the ViewSheet the view is placed on."""
    lookup = {}
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        for view_id in sheet.GetAllPlacedViews():
            lookup[get_element_id_value(view_id)] = sheet
    return lookup


def get_sheet_label(sheet):
    """'number - name' of a sheet, for lists."""
    return "{} - {}".format(sheet.SheetNumber, sheet.Name)


def get_sheet_parameter_names(doc):
    """Names of all parameters found on the project's sheets, sorted."""
    names = set()
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        for parameter in sheet.Parameters:
            names.add(parameter.Definition.Name)
    return sorted(names, key=lambda name: name.lower())


def get_parameter_text(element, name):
    """The value of a parameter as the user sees it, '' if missing or empty."""
    if element is None or not name:
        return ""
    parameter = element.LookupParameter(name)
    if parameter is None or not parameter.HasValue:
        return ""
    if parameter.StorageType == StorageType.String:
        return parameter.AsString() or ""
    return parameter.AsValueString() or ""


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


class _KeepValuesLoadOptions(IFamilyLoadOptions):
    """Load over an already loaded family without touching instance values."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True


def get_family_file(family_dir, revit_version):
    """The newest bundled family file that Revit version can open.

    family_dir holds the original file; subfolders named after a Revit version
    hold copies saved in that version.
    """
    file_name = FAMILY_NAME + ".rfa"
    best_version, best_path = 0, os.path.join(family_dir, file_name)
    for name in os.listdir(family_dir):
        path = os.path.join(family_dir, name, file_name)
        if name.isdigit() and best_version < int(name) <= revit_version and os.path.isfile(path):
            best_version, best_path = int(name), path
    return best_path


def load_reference_family(doc, family_dir):
    """Load the family, or upgrade the loaded one. Call inside an open transaction.

    Existing instances keep their values. Returns True if the project changed.
    """
    path = get_family_file(family_dir, int(doc.Application.VersionNumber))
    loaded = doc.LoadFamily(path, _KeepValuesLoadOptions())
    # IronPython returns (bool, Family) for the overload with an out parameter
    if isinstance(loaded, tuple):
        loaded = loaded[0]
    logger.debug("LoadFamily('{}') -> {}".format(path, loaded))
    return bool(loaded)


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
    return ViewFrame(origin, transform.BasisX, transform.BasisY, width, height, depth,
                     looks_up=view.ViewType == ViewType.CeilingPlan)


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


def _is_on_frame(instance, frame):
    transform = instance.GetTransform()
    return (transform.Origin.IsAlmostEqualTo(frame.placement_origin())
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


def _apply_frame(instance, frame, view, sheet):
    _set_parameter(instance, PARAM_WIDTH, frame.width)
    _set_parameter(instance, PARAM_HEIGHT, frame.height)
    _set_parameter(instance, PARAM_DEPTH, frame.thickness())

    sheet_number = sheet.SheetNumber if sheet is not None else ""
    label = view.Name
    if instance.LookupParameter(PARAM_SHEET) is not None:
        _set_parameter(instance, PARAM_SHEET, sheet_number or NO_SHEET_TEXT)
    elif sheet_number:
        # A line break in the value gives a two-line model text
        label = "{}\r\n{}".format(view.Name, sheet_number)
    _set_parameter(instance, PARAM_NAME, label)


def _place(doc, symbol, frame, view):
    origin = frame.placement_origin()
    plane = Plane.CreateByOriginAndBasis(origin, frame.right, frame.up)
    sketch_plane = SketchPlane.Create(doc, plane)
    instance = doc.Create.NewFamilyInstance(
        origin, symbol, sketch_plane, StructuralType.NonStructural)

    link = ViewReferenceSchema(instance, update=False)
    link.set("view_unique_id", view.UniqueId)
    instance.SetEntity(link.unwrap())
    return instance


def sync_view_references(doc, views, symbol, show_depth=False, manual_depth=None):
    """Create or update one reference per view. Call inside an open transaction.

    With show_depth the reference becomes a box from the cut plane to the
    view's far limit; otherwise it is a thin plate on the cut plane.
    manual_depth (feet) draws every reference that deep instead, whatever the
    view's own depth is.

    An instance that already sits on the right plane only gets its parameters
    refreshed; one that does not is replaced, because a work plane based
    instance cannot be moved to another plane.
    """
    result = SyncResult()
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()

    existing = find_existing_references(doc)
    sheet_lookup = build_sheet_lookup(doc)

    for view in views:
        view_name = view.Name
        sheet = sheet_lookup.get(get_element_id_value(view.Id))
        try:
            frame = get_view_frame(view)
            if frame is None:
                result.skipped.append((view_name, "no crop box"))
                continue
            frame.set_box_depth(show_depth, manual_depth)

            instances = existing.get(view.UniqueId, [])
            keep = None
            for instance in instances:
                if keep is None and _is_on_frame(instance, frame):
                    keep = instance
                else:
                    doc.Delete(instance.Id)

            if keep is not None:
                _apply_frame(keep, frame, view, sheet)
                result.updated.append(keep.Id)
            else:
                instance = _place(doc, symbol, frame, view)
                _apply_frame(instance, frame, view, sheet)
                if instances:
                    result.updated.append(instance.Id)
                else:
                    result.created.append(instance.Id)
        except Exception as ex:
            logger.error("Could not place a reference for '{}': {}".format(view_name, ex))
            result.skipped.append((view_name, str(ex)))

    return result
