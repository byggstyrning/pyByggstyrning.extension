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

One family file is bundled with the Load Family button, saved in Revit 2025
format, so Revit 2025 and newer can load it. With a family that has the
Sheet Number parameter the sheet number goes there; with one that does not,
it becomes a second line of the "View Name" text.

What is written to Sheet Number and to View Name is a formula stored in the
model, like the MMI parameter. It is read left to right. Each part is the
view's own name, the sheet's own number, another parameter of that view or
sheet, or a project information parameter, and the text between two parts is
set on that gap. Empty parts are skipped. Sheet Number defaults to the sheet
number alone. View Name defaults to the view's own name.

Every instance, including ones placed by hand, is set to export to IFC as
IfcVirtualElement. Those are the instance parameters, not the type ones.
"""
import json
import os
import System

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
from Autodesk.Revit.DB.ExtensibleStorage import DataStorage, ExtensibleStorageFilter, Schema
from Autodesk.Revit.DB.Structure import StructuralType

from pyrevit import revit, script

from extensible_storage import BaseSchema, simple_field
from revit.compat import get_element_id_value
from revit.view_reference_formula import (
    DEFAULT_SHEET_FORMULA,
    DEFAULT_VIEW_FORMULA,
    SHEET_NUMBER_LABEL,
    SOURCE_PROJECT,
    SOURCE_SHEET,
    SOURCE_SHEET_NUMBER,
    SOURCE_VIEW,
    SOURCE_VIEW_NAME,
    VIEW_NAME_LABEL,
    compose_parts,
    normalize_sheet_number_formula,
    normalize_view_name_formula,
    part_label,
)

logger = script.get_logger()

FAMILY_NAME = "3D View Reference"
PREFERRED_TYPE_NAME = "Standard Reference"

PARAM_WIDTH = "View Width"
PARAM_HEIGHT = "View Height"
PARAM_DEPTH = "View Depth"
PARAM_NAME = "View Name"
PARAM_SHEET = "Sheet Number"

# Instance overrides. The type has its own "Export Type to IFC" parameters.
# 1 is IFCExportElement.Yes.
IFC_EXPORT_YES = 1
IFC_EXPORT_AS = "IfcVirtualElement"

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


# The settings schema cannot gain a field in place. 2a67 held the sheet formula only.
PREVIOUS_SETTINGS_GUID = "7f4c2a91-8e15-4d6b-b3a0-1c9e5f8d2a67"


class ViewReferenceSettingsSchema(BaseSchema):
    """Sheet-number and view-name formulas for this model, shared by everyone."""

    guid = "7f4c2a91-8e15-4d6b-b3a0-1c9e5f8d2a68"

    @simple_field(value_type="string")
    def sheet_number_formula():
        """JSON object: parts in left-to-right order, with the text between them."""
        return None

    @simple_field(value_type="string")
    def view_name_formula():
        """JSON object: parts written to View Name, left to right."""
        return None


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


def get_project_parameter_names(doc):
    """Names of all parameters on Project Information, sorted."""
    info = doc.ProjectInformation
    if info is None:
        return []
    names = set()
    for parameter in info.Parameters:
        name = parameter.Definition.Name
        if name:
            names.add(name)
    return sorted(names, key=lambda name: name.lower())


def get_view_parameter_names(doc):
    """Names of parameters found on the views these tools can reference, sorted."""
    names = set()
    for view in collect_views(doc):
        for parameter in view.Parameters:
            name = parameter.Definition.Name
            if name:
                names.add(name)
    return sorted(names, key=lambda name: name.lower())


def sheet_number_part_label(part):
    """What a formula part is called in the dropdown."""
    return part_label(part)


def _choice_list(doc, formula, normalize, extra_parts):
    """Dropdown rows: (label, part). Saved parts stay even if the parameter is gone."""
    choices = []
    seen = set()

    def add(label, part):
        if label in seen:
            return
        seen.add(label)
        choices.append((label, part))

    for label, part in extra_parts:
        add(label, part)
    if formula:
        for part in normalize(formula)["parts"]:
            add(part_label(part), part)
    return choices


def _named_parts(source, names, prefix):
    parts = []
    for name in names:
        parts.append((u"{}: {}".format(prefix, name), {"source": source, "name": name}))
    return parts


def sheet_number_choice_list(doc, formula=None):
    """Dropdown rows for the sheet number: sheet parameters and project information."""
    extras = [(SHEET_NUMBER_LABEL, {"source": SOURCE_SHEET_NUMBER})]
    extras.extend(_named_parts(SOURCE_SHEET, get_sheet_parameter_names(doc), "Sheet"))
    extras.extend(_named_parts(SOURCE_PROJECT, get_project_parameter_names(doc), "Project"))
    return _choice_list(doc, formula, normalize_sheet_number_formula, extras)


def view_name_choice_list(doc, formula=None):
    """Dropdown rows for the view name, including the same sheet and project parts."""
    extras = [(VIEW_NAME_LABEL, {"source": SOURCE_VIEW_NAME})]
    extras.extend(_named_parts(SOURCE_VIEW, get_view_parameter_names(doc), "View"))
    extras.append((SHEET_NUMBER_LABEL, {"source": SOURCE_SHEET_NUMBER}))
    extras.extend(_named_parts(SOURCE_SHEET, get_sheet_parameter_names(doc), "Sheet"))
    extras.extend(_named_parts(SOURCE_PROJECT, get_project_parameter_names(doc), "Project"))
    return _choice_list(doc, formula, normalize_view_name_formula, extras)


def _part_text(doc, view, sheet, part):
    source = part.get("source")
    if source == SOURCE_VIEW_NAME:
        return view.Name if view is not None else u""
    if source == SOURCE_VIEW:
        return get_parameter_text(view, part.get("name"))
    if source == SOURCE_SHEET_NUMBER:
        return sheet.SheetNumber if sheet is not None else u""
    if source == SOURCE_SHEET:
        return get_parameter_text(sheet, part.get("name"))
    if source == SOURCE_PROJECT:
        return get_parameter_text(doc.ProjectInformation, part.get("name"))
    return u""


def compose_sheet_number(doc, sheet, formula):
    """The sheet-number text the formula writes. Empty parts are left out."""
    return compose_parts(
        formula,
        SOURCE_SHEET_NUMBER,
        lambda part: _part_text(doc, None, sheet, part),
    )


def compose_view_name(doc, view, sheet, formula):
    """The view-name text the formula writes. Empty parts are left out.

    A formula that resolves to nothing falls back to the view's own name.
    """
    text = compose_parts(
        formula,
        SOURCE_VIEW_NAME,
        lambda part: _part_text(doc, view, sheet, part),
    )
    if text:
        return text
    if view is None:
        return u""
    return view.Name or u""


def _storage_with_schema(doc, schema):
    if schema is None:
        return None
    for storage in FilteredElementCollector(doc).OfClass(DataStorage):
        try:
            entity = storage.GetEntity(schema)
        except Exception:
            continue
        if entity.IsValid():
            return storage
    return None


def _storage_with_guid(doc, guid_text):
    guid = System.Guid(guid_text)
    for storage in FilteredElementCollector(doc).OfClass(DataStorage):
        try:
            if guid in storage.GetEntitySchemaGuids():
                return storage
        except Exception:
            continue
    return None


def _read_schema_string(storage, field_name):
    try:
        value = ViewReferenceSettingsSchema(storage, update=False).get(field_name)
    except Exception:
        return ""
    if value is None:
        return ""
    return value


def _migrate_settings_storage(doc, old_storage):
    """Copy the sheet formula onto the current schema and default the view name."""
    old_guid = System.Guid(PREVIOUS_SETTINGS_GUID)
    raw_sheet = ""
    try:
        old_entity = old_storage.GetEntity(old_guid)
        if old_entity is not None and old_entity.IsValid():
            raw_sheet = old_entity.Get[str]("sheet_number_formula") or ""
    except Exception as ex:
        logger.debug("Could not read the previous view reference settings: {}".format(ex))
    view_raw = json.dumps(DEFAULT_VIEW_FORMULA, ensure_ascii=True, sort_keys=True)
    try:
        with revit.Transaction("Update view reference settings", doc):
            entity = ViewReferenceSettingsSchema(old_storage, update=False)
            if raw_sheet:
                entity.set("sheet_number_formula", raw_sheet)
            entity.set("view_name_formula", view_raw)
            old_storage.SetEntity(entity.unwrap())
            old_schema = Schema.Lookup(old_guid)
            if old_schema is not None:
                old_storage.DeleteEntity(old_schema)
        return old_storage
    except Exception as ex:
        logger.error("Could not update view reference settings: {}".format(ex))
        return None


def _formula_storage(doc):
    """The settings element, migrating the previous schema when that is all the model has."""
    schema = ViewReferenceSettingsSchema.schema
    storage = _storage_with_schema(doc, schema)
    if storage is not None:
        return storage
    old_storage = _storage_with_guid(doc, PREVIOUS_SETTINGS_GUID)
    if old_storage is None:
        return None
    return _migrate_settings_storage(doc, old_storage)


def _load_formula(doc, field_name, normalize, default_formula):
    storage = _formula_storage(doc)
    if storage is None:
        return normalize(default_formula)
    raw = _read_schema_string(storage, field_name)
    if not raw:
        return normalize(default_formula)
    try:
        return normalize(json.loads(raw))
    except Exception as ex:
        logger.debug("Formula '{}' is not readable: {}".format(field_name, ex))
        return normalize(default_formula)


def _save_formula(doc, field_name, formula, normalize, transaction_name):
    """Store one formula. No transaction is started when it is unchanged."""
    formula = normalize(formula)
    raw = json.dumps(formula, ensure_ascii=True, sort_keys=True)
    storage = _formula_storage(doc)
    if storage is not None and _read_schema_string(storage, field_name) == raw:
        return True
    try:
        with revit.Transaction(transaction_name, doc):
            if storage is None:
                storage = DataStorage.Create(doc)
            entity = ViewReferenceSettingsSchema(storage, update=False)
            entity.set(field_name, raw)
            storage.SetEntity(entity.unwrap())
        return True
    except Exception as ex:
        logger.error("Could not save {}: {}".format(field_name, ex))
        return False


def load_sheet_number_formula(doc):
    """The formula saved in this model, or the sheet number on its own."""
    return _load_formula(
        doc, "sheet_number_formula", normalize_sheet_number_formula, DEFAULT_SHEET_FORMULA)


def save_sheet_number_formula(doc, formula):
    """Store the sheet number formula. Returns True when the model holds it."""
    return _save_formula(
        doc, "sheet_number_formula", formula,
        normalize_sheet_number_formula, "Save sheet number formula")


def load_view_name_formula(doc):
    """The formula saved in this model, or the view's own name."""
    return _load_formula(
        doc, "view_name_formula", normalize_view_name_formula, DEFAULT_VIEW_FORMULA)


def save_view_name_formula(doc, formula):
    """Store the view name formula. Returns True when the model holds it."""
    return _save_formula(
        doc, "view_name_formula", formula,
        normalize_view_name_formula, "Save view name formula")


def get_view_area(view):
    """Crop box width x height in square metres, None without a usable crop box."""
    crop = view.CropBox
    if crop is None:
        return None
    width = crop.Max.X - crop.Min.X
    height = crop.Max.Y - crop.Min.Y
    if width <= 0 or height <= 0:
        return None
    return width * height * 0.09290304


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


def get_family_file(family_dir):
    """The bundled family file, saved in Revit 2025 format."""
    return os.path.join(family_dir, FAMILY_NAME + ".rfa")


def load_reference_family(doc, family_dir):
    """Load the family, or upgrade the loaded one. Call inside an open transaction.

    Existing instances keep their values. Returns True if the project changed.
    """
    path = get_family_file(family_dir)
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


_ifc_warned = set()


def _warn_ifc_once(name):
    if name in _ifc_warned:
        return
    _ifc_warned.add(name)
    logger.warning("Instance parameter '{}' could not be set".format(name))


def _set_ifc_export(instance):
    """Export this instance to IFC as IfcVirtualElement."""
    export_flag = instance.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT)
    if export_flag is None or export_flag.IsReadOnly:
        _warn_ifc_once("Export to IFC")
    elif export_flag.AsInteger() != IFC_EXPORT_YES:
        export_flag.Set(IFC_EXPORT_YES)

    export_as = instance.get_Parameter(BuiltInParameter.IFC_EXPORT_ELEMENT_AS)
    if export_as is None or export_as.IsReadOnly:
        _warn_ifc_once("Export to IFC As")
    elif export_as.AsString() != IFC_EXPORT_AS:
        export_as.Set(IFC_EXPORT_AS)


def _set_ifc_export_on_family(doc, symbol):
    """Set the IFC instance overrides on every instance of this family.

    Instances placed by hand have no view link, so the update pass would
    otherwise leave them alone.
    """
    family_id = get_element_id_value(symbol.Family.Id)
    for instance in FilteredElementCollector(doc).OfClass(FamilyInstance):
        try:
            instance_family_id = get_element_id_value(instance.Symbol.Family.Id)
        except Exception:
            continue
        if instance_family_id == family_id:
            _set_ifc_export(instance)


def _set_parameter(instance, name, value):
    param = instance.LookupParameter(name)
    if param is None:
        logger.warning("Parameter '{}' not found on the {} family".format(name, FAMILY_NAME))
        return
    if param.IsReadOnly:
        logger.warning("Parameter '{}' is read-only".format(name))
        return
    param.Set(value)


def _apply_frame(instance, frame, view, sheet, doc, formula, view_formula):
    _set_parameter(instance, PARAM_WIDTH, frame.width)
    _set_parameter(instance, PARAM_HEIGHT, frame.height)
    _set_parameter(instance, PARAM_DEPTH, frame.thickness())

    sheet_number = compose_sheet_number(doc, sheet, formula)
    label = compose_view_name(doc, view, sheet, view_formula)
    if instance.LookupParameter(PARAM_SHEET) is not None:
        _set_parameter(instance, PARAM_SHEET, sheet_number or NO_SHEET_TEXT)
    elif sheet_number:
        # A line break in the value gives a two-line model text
        label = "{}\r\n{}".format(label, sheet_number)
    _set_parameter(instance, PARAM_NAME, label)
    _set_ifc_export(instance)


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


def sync_view_references(doc, views, symbol, show_depth=False, manual_depth=None,
                         sheet_formula=None, view_formula=None):
    """Create or update one reference per view. Call inside an open transaction.

    With show_depth the reference becomes a box from the cut plane to the
    view's far limit; otherwise it is a thin plate on the cut plane.
    manual_depth (feet) draws every reference that deep instead, whatever the
    view's own depth is.
    sheet_formula and view_formula override the formulas stored in the model.

    An instance that already sits on the right plane only gets its parameters
    refreshed; one that does not is replaced, because a work plane based
    instance cannot be moved to another plane.
    """
    result = SyncResult()
    if sheet_formula is None:
        sheet_formula = load_sheet_number_formula(doc)
    else:
        sheet_formula = normalize_sheet_number_formula(sheet_formula)
    if view_formula is None:
        view_formula = load_view_name_formula(doc)
    else:
        view_formula = normalize_view_name_formula(view_formula)
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
                _apply_frame(keep, frame, view, sheet, doc, sheet_formula, view_formula)
                result.updated.append(keep.Id)
            else:
                instance = _place(doc, symbol, frame, view)
                _apply_frame(instance, frame, view, sheet, doc, sheet_formula, view_formula)
                if instances:
                    result.updated.append(instance.Id)
                else:
                    result.created.append(instance.Id)
        except Exception as ex:
            logger.error("Could not place a reference for '{}': {}".format(view_name, ex))
            result.skipped.append((view_name, str(ex)))

    _set_ifc_export_on_family(doc, symbol)
    return result
