# -*- coding: utf-8 -*-
"""Build '2026/3D View Reference.rfa' from '3D View Reference.rfa'.

Not a pyRevit button. Run it headlessly from a machine with Revit 2026:

    pyrevit run "<this file>" --revit=2026

The family next to this script is kept in Revit 2024 format so that every supported
Revit can load it. The 2026 copy adds a 'Sheet Number' text below the 'View Name' text.

That text is a nested label family rather than a second model text. A model text has
no references, so it cannot be dimensioned, and the way the View Name text is held to
the frame (grouped with an invisible model line that is locked to the left reference
plane) cannot be reproduced through the API: FamilyCreate.NewGroup refuses model lines.
A nested family instance does have references, so its Center (Left/Right) reference is
locked to the same left reference plane, at the distance the View Name text keeps.

A report, build_2026_family.json, is written to the TEMP folder of that Revit session
(pyrevit run gives every run a folder of its own under %TEMP%). pyrevit run exits 0
whatever happens, so read the report. The file is only saved if the new label follows
the frame when the family is flexed.
"""
import io
import json
import os
import shutil
import tempfile
import traceback

from Autodesk.Revit.DB import (
    BuiltInParameter, Dimension, ElementTransformUtils, FamilyInstanceReferenceType,
    FamilySymbol, FilteredElementCollector, GroupTypeId, HorizontalAlign, IFamilyLoadOptions,
    Level, Line, ModelLine, ModelText, ModelTextType, ReferenceArray, ReferencePlane,
    SaveAsOptions, SketchPlane, SpecTypeId, Transaction, XYZ,
)
from Autodesk.Revit.DB.Structure import StructuralType

HERE = os.path.dirname(os.path.abspath(__file__))
FAMILY_FILE = "3D View Reference.rfa"
SOURCE = os.path.join(HERE, FAMILY_FILE)
TARGET = os.path.join(HERE, "2026", FAMILY_FILE)
LABEL_FAMILY = "3D View Reference Label"
TEMPLATE_NAME = os.path.join("English", "Metric Generic Model.rft")
REPORT = os.path.join(tempfile.gettempdir(), "build_2026_family.json")

SHEET_PARAM = "Sheet Number"
LABEL_TEXT_PARAM = "Label Text"    # "Label" and "Depth" clash with built-in names
LABEL_DEPTH_PARAM = "Label Depth"  # on a nested instance and cannot be associated
LINE_SPACING = 1.5  # text heights between the View Name and Sheet Number baselines

report = {"errors": []}


class _Overwrite(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True


def _save_as(doc, path):
    options = SaveAsOptions()
    options.OverwriteExistingFile = True
    options.MaximumBackups = 1
    doc.SaveAs(path, options)


def _roll_back(transaction):
    if transaction is not None and transaction.HasStarted() and not transaction.HasEnded():
        transaction.RollBack()


def build_label(app, path, text_size, font):
    """A generic model holding one right-aligned model text that ends on the family origin."""
    doc = app.NewFamilyDocument(os.path.join(app.FamilyTemplatePath, TEMPLATE_NAME))
    transaction = None
    try:
        manager = doc.FamilyManager
        transaction = Transaction(doc, "Build label")
        transaction.Start()
        if manager.CurrentType is None:
            manager.NewType("Label")
        text_param = manager.AddParameter(
            LABEL_TEXT_PARAM, GroupTypeId.Text, SpecTypeId.String.Text, True)
        depth_param = manager.AddParameter(
            LABEL_DEPTH_PARAM, GroupTypeId.Geometry, SpecTypeId.Length, True)
        manager.Set(text_param, SHEET_PARAM)
        manager.Set(depth_param, 10 / 304.8)

        text_type = FilteredElementCollector(doc).OfClass(ModelTextType).FirstElement()
        text_type.get_Parameter(BuiltInParameter.MODEL_TEXT_SIZE).Set(text_size)
        text_type.get_Parameter(BuiltInParameter.TEXT_FONT).Set(font)
        level = FilteredElementCollector(doc).OfClass(Level).FirstElement()
        text = doc.FamilyCreate.NewModelText(
            SHEET_PARAM, text_type, SketchPlane.Create(doc, level.Id), XYZ.Zero,
            HorizontalAlign.Right, 10 / 304.8)
        doc.Regenerate()
        # NewModelText starts the text at the point whatever the alignment
        ElementTransformUtils.MoveElement(doc, text.Id, XYZ.Zero - text.Location.Point)
        manager.AssociateElementParameterToFamilyParameter(
            text.get_Parameter(BuiltInParameter.TEXT_TEXT), text_param)
        manager.AssociateElementParameterToFamilyParameter(
            text.get_Parameter(BuiltInParameter.EXTRUSION_LENGTH), depth_param)
        transaction.Commit()
        _save_as(doc, path)
    except Exception:
        _roll_back(transaction)
        raise
    finally:
        doc.Close(False)


def find_text_anchor(doc, name_text):
    """The reference plane the View Name text's group is locked to, and the plan view."""
    group = doc.GetElement(name_text.GroupId)
    member_ids = set(i.ToString() for i in group.GetMemberIds())
    for dimension in FilteredElementCollector(doc).OfClass(Dimension):
        try:
            if not dimension.IsLocked or dimension.View is None:
                continue
        except Exception:
            continue  # multi-segment dimensions have no single lock
        elements = [doc.GetElement(r.ElementId) for r in dimension.References]
        lines = [e for e in elements if isinstance(e, ModelLine) and e.Id.ToString() in member_ids]
        planes = [e for e in elements if isinstance(e, ReferencePlane)]
        if len(elements) == 2 and lines and planes:
            return planes[0], dimension.View
    raise Exception("No locked dimension holds the View Name text")


def _positions(name_text, label):
    return {"view_name_x": round(name_text.Location.Point.X, 4),
            "sheet_number_x": round(label.Location.Point.X, 4)}


def _flex(doc, width, name_text, label):
    manager = doc.FamilyManager
    transaction = Transaction(doc, "Flex")
    transaction.Start()
    manager.Set(manager.get_Parameter("View Width"), width)
    doc.Regenerate()
    transaction.Commit()
    return _positions(name_text, label)


def build(app):
    work_dir = tempfile.mkdtemp(prefix="view_reference_build_")
    work_file = os.path.join(work_dir, FAMILY_FILE)
    label_file = os.path.join(work_dir, LABEL_FAMILY + ".rfa")
    shutil.copyfile(SOURCE, work_file)

    doc = app.OpenDocumentFile(work_file)
    transaction = None
    try:
        manager = doc.FamilyManager
        if manager.get_Parameter(SHEET_PARAM) is not None:
            raise Exception("The source family already has '{}'".format(SHEET_PARAM))
        name_text = FilteredElementCollector(doc).OfClass(ModelText).FirstElement()
        text_type = doc.GetElement(name_text.GetTypeId())
        text_size = text_type.get_Parameter(BuiltInParameter.MODEL_TEXT_SIZE).AsDouble()
        font = text_type.get_Parameter(BuiltInParameter.TEXT_FONT).AsString()
        default_width = manager.CurrentType.AsDouble(manager.get_Parameter("View Width"))

        build_label(app, label_file, text_size, font)
        left_plane, plan_view = find_text_anchor(doc, name_text)
        name_param = manager.get_Parameter("View Name")

        transaction = Transaction(doc, "Add Sheet Number label")
        transaction.Start()
        doc.LoadFamily(label_file, _Overwrite())
        symbol = [s for s in FilteredElementCollector(doc).OfClass(FamilySymbol)
                  if s.Family.Name == LABEL_FAMILY][0]
        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()

        sheet_param = manager.AddParameter(
            SHEET_PARAM, name_param.Definition.GetGroupTypeId(),
            name_param.Definition.GetDataType(), True)
        manager.Set(sheet_param, SHEET_PARAM)

        anchor = name_text.Location.Point
        position = XYZ(anchor.X, anchor.Y - text_size * LINE_SPACING, 0)
        level = FilteredElementCollector(doc).OfClass(Level).FirstElement()
        label = doc.FamilyCreate.NewFamilyInstance(
            position, symbol, level, StructuralType.NonStructural)
        doc.Regenerate()

        manager.AssociateElementParameterToFamilyParameter(
            label.LookupParameter(LABEL_TEXT_PARAM), sheet_param)
        manager.AssociateElementParameterToFamilyParameter(
            label.LookupParameter(LABEL_DEPTH_PARAM), manager.get_Parameter("View Depth"))

        references = ReferenceArray()
        references.Append(left_plane.GetReference())
        references.Append(list(label.GetReferences(FamilyInstanceReferenceType.CenterLeftRight))[0])
        dimension = doc.FamilyCreate.NewDimension(
            plan_view, Line.CreateBound(position, position + XYZ.BasisX), references)
        dimension.IsLocked = True
        transaction.Commit()

        # The build is only good if both texts move together when the frame changes
        wide = _flex(doc, 30.0, name_text, label)
        narrow = _flex(doc, 2.0, name_text, label)
        _flex(doc, default_width, name_text, label)
        report["flex"] = {"wide": wide, "narrow": narrow}
        for state in (wide, narrow):
            if abs(state["view_name_x"] - state["sheet_number_x"]) > 1e-6:
                raise Exception("The Sheet Number label does not follow the frame: {}".format(state))

        if not os.path.isdir(os.path.dirname(TARGET)):
            os.makedirs(os.path.dirname(TARGET))
        _save_as(doc, TARGET)
        report["saved"] = TARGET
    except Exception:
        _roll_back(transaction)
        raise
    finally:
        doc.Close(False)
        shutil.rmtree(work_dir, ignore_errors=True)


try:
    build(__revit__.Application)  # noqa: F821
except Exception:
    report["errors"].append(traceback.format_exc())
finally:
    text = json.dumps(report, ensure_ascii=False, indent=1, sort_keys=True)
    if isinstance(text, bytes):
        text = text.decode("utf-8")
    with io.open(REPORT, "w", encoding="utf-8") as f:
        f.write(text)
