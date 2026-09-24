# -*- coding: utf-8 -*-
"""Revit Batch Processor task script: Autodesk file-size analysis.

Follows Autodesk's "How to keep the size of Revit files manageable":
https://www.autodesk.com/support/technical/article/caas/sfdcarticles/sfdcarticles/Revit-How-to-keep-size-of-Revit-files-manageable.html

Does NOT modify the production model. RBP should open a detached copy (or the
pyByggstyrning "File Size Analysis" button saves a new central first); this
script then Save As compact snapshots into an output folder, measuring file
size after each reduction so you can see where the megabytes live.

Steps (each Save As + size check):
  00  compact Save As (baseline)
  01  purge unused (until stable)
  02  remove Revit links
  03  remove CAD links / imports
  04  delete unplaced group types
  05  remove raster images, PDFs, point clouds
  06  remove all sheets
  07  remove all views except the dedicated keep 3D
  08  unused view templates + unused filters
  09  remove in-place families
  10  unused project / shared parameters
  11  remove unused loaded families + purge again
  12  purify placed loadable families (optional; detached-copy analysis)
  13  clear design options (non-primary, then option sets)
  14+ one compact Save As per phase (delete model created in that phase)
      leftover unphased model, then all but one level

Also writes a loaded-family inventory, health metrics, CSV/JSON/Markdown,
and a self-contained report.html.

Environment (optional):
  RBP_SIZE_OUT_DIR          Output folder. Default: <model dir>\\_filesize_analysis\\<name>_<stamp>
  RBP_KEEP_STEP_FILES       1 (default) keep every snapshot; 0 delete previous snapshot after the next save
  RBP_MEASURE_FAMILY_SIZES  1 (default unless purifying) pre-measure every family; 0 to skip
  RBP_FAMILY_EXPORT_DIR     Folder of already-exported .rfa files; used for size + name matching
  RBP_FAMILY_ANALYTICS_ONLY 1 skip the size-reduction steps; inventory + family report only
  RBP_PURGE_PASSES          Purge loop count (default 15; Autodesk purge is repeated until empty)
  RBP_PURIFY_FAMILIES       1 run Family Purify analysis profile; 0 (default) skip
  RBP_PURIFY_FAMILY_MAX     Maximum families to process; 0 (default) means all
  RBP_PURIFY_FAMILY_OFFSET  Skip this many ranked candidates (default 0)
  RBP_FAMILY_PURIFY_ONLY    1 run compact baseline + family purification only
  RBP_FAMILY_PURIFY_AGGRESSIVE 1 enable detached-only nested/type/view/purge cleanup
  RBP_FAMILY_RANKING_CSV    Prior *_families.csv used to rank largest families
  RBP_PURIFY_RELOAD_MIN_KB  Minimum measured saving before reload (aggressive: 5120)
  RBP_SIDECAR_PATH          Optional JSON sidecar path

RBP settings: open detached from central; do not save the original.
"""

from __future__ import print_function

import datetime
import csv
import json
import os
import sys
import traceback

import clr
import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    CategoryType,
    ElementId,
    Family,
    FamilyInstance,
    FamilySource,
    FilteredElementCollector,
    IFailuresPreprocessor,
    ImportInstance,
    IFamilyLoadOptions,
    FailureProcessingResult,
    FailureSeverity,
    Level,
    LinePatternElement,
    FillPatternElement,
    AppearanceAssetElement,
    RevitLinkInstance,
    RevitLinkType,
    SaveAsOptions,
    StorageType,
    Transaction,
    TransactionGroup,
    TransactionStatus,
    View,
    View3D,
    ViewFamily,
    ViewFamilyType,
    ViewSheet,
    ViewType,
    WorksharingSaveAsOptions,
)
from System.Collections.Generic import HashSet, List as NetList  # noqa: E402
from System.Reflection import BindingFlags  # noqa: E402

try:
    from Autodesk.Revit.DB import ImageType
except Exception:
    ImageType = None

try:
    from Autodesk.Revit.DB import CADLinkType
except Exception:
    CADLinkType = None

try:
    from Autodesk.Revit.DB.PointClouds import PointCloudInstance, PointCloudType
except Exception:
    PointCloudInstance = None
    PointCloudType = None

try:
    from Autodesk.Revit.DB import Grid
except Exception:
    Grid = None

try:
    from Autodesk.Revit.DB import Group, GroupType
except Exception:
    Group = None
    GroupType = None

try:
    from Autodesk.Revit.DB import ParameterFilterElement
except Exception:
    ParameterFilterElement = None

try:
    from Autodesk.Revit.DB import ParameterElement
except Exception:
    ParameterElement = None

try:
    from Autodesk.Revit.DB import SharedParameterElement
except Exception:
    SharedParameterElement = None

try:
    from Autodesk.Revit.DB import FilledRegion, FilledRegionType
except Exception:
    FilledRegion = None
    FilledRegionType = None

try:
    from Autodesk.Revit.DB import TextNote, TextNoteType, Dimension, DimensionType
except Exception:
    TextNote = None
    TextNoteType = None
    Dimension = None
    DimensionType = None

try:
    from Autodesk.Revit.DB import DesignOption
except Exception:
    DesignOption = None

try:
    from Autodesk.Revit.DB import Phase
except Exception:
    Phase = None

try:
    from Autodesk.Revit.DB.Architecture import Room
except Exception:
    Room = None

try:
    from Autodesk.Revit.DB import FilteredWorksetCollector, WorksetKind
except Exception:
    FilteredWorksetCollector = None
    WorksetKind = None

try:
    from Autodesk.Revit.DB import Viewport
except Exception:
    Viewport = None

try:
    from Autodesk.Revit.DB import BuiltInParameter
except Exception:
    BuiltInParameter = None

# Host detection: Revit Batch Processor provides revit_script_util; inside a
# pyRevit pushbutton it is missing and the caller wires the host up through
# configure_host() / set_output_sink() before calling run().
try:
    import revit_script_util
    from revit_script_util import Output as _rbp_output
    RUNNING_IN_RBP = True
except ImportError:
    revit_script_util = None
    _rbp_output = None
    RUNNING_IN_RBP = False

_output_sink = None


def set_output_sink(fn):
    """Route log lines to fn(msg) instead of RBP Output / print."""
    global _output_sink
    _output_sink = fn


def Output(msg=None):
    if _output_sink is not None:
        if msg is not None:
            _output_sink(msg)
        return
    if _rbp_output is not None:
        if msg is None:
            _rbp_output()
        else:
            _rbp_output(msg)
        return
    if msg is not None:
        print(msg)


if RUNNING_IN_RBP:
    sessionId = revit_script_util.GetSessionId()
    uiapp = revit_script_util.GetUIApplication()
    doc = revit_script_util.GetScriptDocument()
    revitFilePath = revit_script_util.GetRevitFilePath() or ""
else:
    sessionId = "pyrevit"
    uiapp = None
    doc = None
    revitFilePath = ""


def configure_host(ui_application, session_id=None):
    """Set the UIApplication (and report session id) when not running in RBP."""
    global uiapp, sessionId
    uiapp = ui_application
    if session_id:
        sessionId = session_id

LOG_TAG = "[RBP-SIZE]"

KEEP_VIEW_NAME = "RBP_SizeAnalysis_3D"
SOURCE_ARTICLE = (
    "https://www.autodesk.com/support/technical/article/caas/"
    "sfdcarticles/sfdcarticles/Revit-How-to-keep-size-of-Revit-files-manageable.html"
)

# 00 compact through 11 loaded_families, family purification, design options,
# N phases, leftover geometry, levels, final purge
STEPS_THROUGH_UNUSED_FAMILIES = 12


def _expected_snapshot_count(n_phases, include_family_purify=False):
    return (
        STEPS_THROUGH_UNUSED_FAMILIES
        + (1 if include_family_purify else 0)
        + 1
        + int(n_phases or 0)
        + 1
        + 1
        + 1
    )


# ---------------------------------------------------------------------------
# Logging / paths
# ---------------------------------------------------------------------------


def log(msg):
    Output("{} {}".format(LOG_TAG, msg))


def _env_flag(name, default=False):
    raw = os.environ.get(name, "")
    if not raw:
        return default
    return raw.strip().lower() in ("1", "true", "yes", "y", "on")


def _env_int(name, default):
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        return int(raw)
    except Exception:
        return default


def _safe_str(x):
    if x is None:
        return ""
    try:
        unicode_type = unicode  # noqa: F821
    except NameError:
        unicode_type = str
    if isinstance(x, unicode_type):
        try:
            return x.strip()
        except Exception:
            return x
    if isinstance(x, str):
        for enc in ("utf-8", "cp1252", "latin-1"):
            try:
                return x.decode(enc).strip()
            except Exception:
                continue
        try:
            return unicode_type(x, "latin-1", "replace").strip()
        except Exception:
            return repr(x)
    try:
        text = unicode_type(x).strip()
    except Exception:
        try:
            text = str(x)
        except Exception:
            return ""
    # IronPython can hand back a byte string (cp1252 Ä = 0xC4) instead of unicode.
    if sys.version_info[0] < 3 and isinstance(text, str) and not isinstance(text, unicode_type):
        for enc in ("utf-8", "cp1252", "latin-1"):
            try:
                return text.decode(enc).strip()
            except Exception:
                continue
    return text


def _safe_filename(name):
    text = _safe_str(name) or "unnamed"
    bad = '<>:"/\\|?*'
    chars = []
    for c in text:
        chars.append("_" if c in bad else c)
    cleaned = "".join(chars).strip(" .")
    if not cleaned:
        cleaned = "unnamed"
    return cleaned[:80]


def _id_value(element_id):
    """Numeric value of an ElementId.

    Revit 2026 removed ElementId.IntegerValue (deprecated since 2024). .Value is
    the long-typed replacement and exists from Revit 2024 on; the fallback keeps
    Revit 2023 and earlier working. Only AttributeError is caught, so callers
    relying on their own try/except keep identical behaviour.
    """
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def _eid(el_or_id):
    try:
        if el_or_id is None:
            return None
        if isinstance(el_or_id, ElementId):
            return _id_value(el_or_id)
        return _id_value(el_or_id.Id)
    except Exception:
        return None


def _file_bytes(path):
    try:
        if path and os.path.isfile(path):
            return int(os.path.getsize(path))
    except Exception:
        pass
    return None


def _mb(n):
    if n is None:
        return None
    return round(float(n) / (1024.0 * 1024.0), 3)


def _pct(part, whole):
    if part is None or not whole:
        return None
    return round(100.0 * float(part) / float(whole), 2)


def _decode_str_to_unicode(s):
    """Normalize byte/str payloads for JSON (IronPython may expose UTF-8 bytes as str)."""
    if not s:
        try:
            return unicode("")  # noqa: F821
        except NameError:
            return ""
    try:
        return s.decode("utf-8")
    except Exception:
        try:
            return unicode(s, "utf-8", errors="replace")  # noqa: F821
        except Exception:
            try:
                return unicode(s, "latin-1", errors="replace")  # noqa: F821
            except Exception:
                try:
                    return unicode("")
                except NameError:
                    return ""


def sanitize_for_json(obj):
    """Recursively convert dicts so json.dumps never hits UnicodeDecodeError on aao / m2.

    Copied from model_audit.py (do not import sibling modules from an RBP task script).
    """
    try:
        if sys.version_info[0] >= 3:
            if isinstance(obj, dict):
                return {
                    sanitize_for_json(k): sanitize_for_json(v) for k, v in obj.items()
                }
            if isinstance(obj, (list, tuple)):
                return [sanitize_for_json(x) for x in obj]
            if isinstance(obj, bytes):
                return obj.decode("utf-8", errors="replace")
            if isinstance(obj, (int, float, bool)) or obj is None:
                return obj
            if isinstance(obj, str):
                return obj
            return str(obj)
        try:
            _int_types = (int, long)  # noqa: F821
        except NameError:
            _int_types = (int,)
        if isinstance(obj, dict):
            return {
                sanitize_for_json(k): sanitize_for_json(v) for k, v in obj.items()
            }
        if isinstance(obj, (list, tuple)):
            return [sanitize_for_json(x) for x in obj]
        try:
            u = unicode  # noqa: F821
        except NameError:
            u = str
        if isinstance(obj, u):
            return obj
        if isinstance(obj, str):
            return _decode_str_to_unicode(obj)
        if isinstance(obj, _int_types + (float, bool)) or obj is None:
            return obj
        try:
            return u(obj)
        except Exception:
            try:
                return repr(obj)
            except Exception:
                return u"<unserializable>"
    except Exception:
        return None


def _json_safe(obj):
    cleaned = sanitize_for_json(obj)
    if cleaned is None:
        return _safe_str(obj)
    return cleaned


def _html_escape(text):
    s = _safe_str(text)
    s = s.replace("&", "&amp;")
    s = s.replace("<", "&lt;")
    s = s.replace(">", "&gt;")
    s = s.replace('"', "&quot;")
    return s


def _try_write(label, fn):
    try:
        path = fn()
        if path:
            log("{}: {}".format(label, path))
        else:
            log("{}: ok".format(label))
        return True
    except Exception as ex:
        log("WARNING: {} failed: {}".format(label, ex))
        return False


# ---------------------------------------------------------------------------
# Failure handling
# ---------------------------------------------------------------------------


class WarningSwallower(IFailuresPreprocessor):
    """Delete warnings and resolve errors without user interaction.

    Errors such as "Constraints are not satisfied" get their default
    resolution (e.g. Remove Constraints, Delete Element) - the same button
    the interactive Revit dialog would otherwise wait for.
    """

    def PreprocessFailures(self, accessor):
        resolved = False
        try:
            for msg in list(accessor.GetFailureMessages()):
                severity = msg.GetSeverity()
                if severity == FailureSeverity.Warning:
                    accessor.DeleteWarning(msg)
                elif severity == FailureSeverity.Error and msg.HasResolutions():
                    accessor.ResolveFailure(msg)
                    resolved = True
        except Exception:
            try:
                accessor.DeleteAllWarnings()
            except Exception:
                pass
        if resolved:
            return FailureProcessingResult.ProceedWithCommit
        return FailureProcessingResult.Continue


def _swallow_warnings(txn):
    try:
        opts = txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(WarningSwallower())
        opts.SetClearAfterRollback(True)
        txn.SetFailureHandlingOptions(opts)
    except Exception:
        pass


def _run_transaction(the_doc, name, fn):
    """Run fn(doc) inside a transaction. Returns (ok, result_or_error)."""
    txn = Transaction(the_doc, name[:79])
    _swallow_warnings(txn)
    txn.Start()
    try:
        result = fn(the_doc)
        if txn.GetStatus() == TransactionStatus.Started:
            status = txn.Commit()
            if status != TransactionStatus.Committed:
                return False, "commit returned {}".format(status)
        return True, result
    except Exception as ex:
        try:
            if txn.GetStatus() == TransactionStatus.Started:
                txn.RollBack()
        except Exception:
            pass
        return False, ex


class OverwriteFamilyLoadOptions(IFamilyLoadOptions):
    """Reload the edited family without overwriting project type-parameter values."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = FamilySource.Family
        overwriteParameterValues.Value = False
        return True


def _save_family_measure(fam_doc, path, compact):
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    if os.path.isfile(path):
        try:
            os.remove(path)
        except Exception:
            pass
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    try:
        opts.Compact = bool(compact)
    except Exception:
        pass
    try:
        opts.MaximumBackups = 1
    except Exception:
        pass
    fam_doc.SaveAs(path, opts)
    return _file_bytes(path)


def _family_txn(fam_doc, name, fn):
    """Run a best-effort Family Purify operation and swallow Revit warnings."""
    ok, result = _run_transaction(fam_doc, name, fn)
    if not ok:
        raise RuntimeError("{}: {}".format(name, result))
    return result


def _purify_imports_images(fam_doc):
    ids = [el.Id for el in _collect_class(fam_doc, ImportInstance)]
    if ImageType is not None:
        ids.extend(el.Id for el in _collect_class(fam_doc, ImageType, not_types=False))
    return _delete_ids(fam_doc, ids)


def _purify_unused_subcategories(fam_doc):
    """Port of Family Purify's category-usage test."""
    elems = list(
        FilteredElementCollector(fam_doc).WhereElementIsNotElementType().ToElements()
    )
    used = set()
    parents = {}
    for el in elems:
        try:
            cat = el.Category
            if cat is None:
                continue
            cid = _id_value(cat.Id)
            used.add(cid)
            if cid not in parents:
                parents[cid] = cat
        except Exception:
            continue

    ids = []
    for parent in parents.values():
        try:
            it = parent.SubCategories.GetEnumerator()
            while it.MoveNext():
                subcat = it.Current
                name = _safe_str(subcat.Name).lower()
                if name in ("<sketch>", "invisible lines"):
                    continue
                if _id_value(subcat.Id) not in used:
                    ids.append(subcat.Id)
        except Exception:
            continue
    return _delete_ids(fam_doc, ids)


def _scan_element_id_parameters(elements, id_sets, used_sets):
    for el in elements:
        try:
            params = el.Parameters
        except Exception:
            continue
        for param in params:
            try:
                if param.StorageType != StorageType.ElementId:
                    continue
                eid = param.AsElementId()
                if eid is None or eid == ElementId.InvalidElementId:
                    continue
                for i in range(len(id_sets)):
                    if eid in id_sets[i]:
                        used_sets[i].add(eid)
            except Exception:
                continue


def _purify_unused_patterns_assets(fam_doc):
    """Port Family Purify's ElementId-reference scan and delete unused styles."""
    line_patterns = _collect_class(fam_doc, LinePatternElement, not_types=False)
    fill_patterns = _collect_class(fam_doc, FillPatternElement, not_types=False)
    assets = _collect_class(fam_doc, AppearanceAssetElement, not_types=False)
    all_sets = [
        set(el.Id for el in line_patterns),
        set(el.Id for el in fill_patterns),
        set(el.Id for el in assets),
    ]
    used_sets = [set(), set(), set()]

    try:
        from Autodesk.Revit.DB import Material

        for material in _collect_class(fam_doc, Material, not_types=False):
            try:
                aid = material.AppearanceAssetId
                if aid is not None and aid != ElementId.InvalidElementId and aid in all_sets[2]:
                    used_sets[2].add(aid)
            except Exception:
                continue
    except Exception:
        pass

    instances = FilteredElementCollector(fam_doc).WhereElementIsNotElementType().ToElements()
    types = FilteredElementCollector(fam_doc).WhereElementIsElementType().ToElements()
    _scan_element_id_parameters(instances, all_sets, used_sets)
    _scan_element_id_parameters(types, all_sets, used_sets)

    # Match Family Purify's dependency-friendly order: assets, fills, then lines.
    ids = []
    ids.extend(list(all_sets[2] - used_sets[2]))
    ids.extend(list(all_sets[1] - used_sets[1]))
    ids.extend(list(all_sets[0] - used_sets[0]))
    return _delete_ids(fam_doc, ids)


def _purify_unused_materials(fam_doc):
    try:
        from Autodesk.Revit.DB import Material
    except Exception:
        return 0, 0
    materials = _collect_class(fam_doc, Material, not_types=False)
    all_ids = set(el.Id for el in materials)
    used = set()
    instances = FilteredElementCollector(fam_doc).WhereElementIsNotElementType().ToElements()
    types = FilteredElementCollector(fam_doc).WhereElementIsElementType().ToElements()
    _scan_element_id_parameters(instances, [all_ids], [used])
    _scan_element_id_parameters(types, [all_ids], [used])
    return _delete_ids(fam_doc, list(all_ids - used))


def _purify_non_core_views(fam_doc):
    keep_types = set(
        [
            ViewType.FloorPlan,
            ViewType.CeilingPlan,
            ViewType.Elevation,
            ViewType.Section,
            ViewType.ThreeD,
        ]
    )
    active_id = None
    try:
        active_id = fam_doc.ActiveView.Id
    except Exception:
        pass
    ids = []
    for view in _collect_class(fam_doc, View):
        try:
            if view.IsTemplate or (active_id is not None and view.Id == active_id):
                continue
            name = _safe_str(view.Name).lower()
            if view.ViewType in keep_types:
                continue
            if "ref" in name or "reference" in name or "level" in name:
                continue
            ids.append(view.Id)
        except Exception:
            continue
    return _delete_ids(fam_doc, ids)


def _purify_family_types(fam_doc, delete_type_names):
    manager = fam_doc.FamilyManager
    types = list(manager.Types)
    names = set(_safe_str(x).lower() for x in (delete_type_names or []))
    candidates = [t for t in types if _safe_str(t.Name).lower() in names]
    removed = 0
    failed = 0
    for family_type in candidates:
        if len(types) - removed <= 1:
            break
        try:
            manager.CurrentType = family_type
            manager.DeleteCurrentType()
            removed += 1
        except Exception:
            failed += 1
    return removed, failed


def _project_unused_type_names_by_family(the_doc):
    """Use Revit's Purge Unused result to identify safely removable family types."""
    unused_ids = set()
    try:
        unused = the_doc.GetUnusedElements(HashSet[ElementId]())
        unused_ids = set(_eid(x) for x in unused)
    except Exception as ex:
        log("  unused family-type scan failed: {}".format(ex))
        return {}
    result = {}
    for fam in _collect_class(the_doc, Family, not_types=False):
        names = []
        try:
            for symbol_id in fam.GetFamilySymbolIds():
                if _eid(symbol_id) not in unused_ids:
                    continue
                symbol = the_doc.GetElement(symbol_id)
                if symbol is not None:
                    names.append(_safe_str(symbol.Name))
        except Exception:
            continue
        if names:
            result[_eid(fam)] = names
    return result


def _purge_project_types_for_families(the_doc, family_ids):
    target_ids = set(int(x) for x in (family_ids or []) if x is not None)
    if not target_ids:
        return 0, 0
    try:
        unused = the_doc.GetUnusedElements(HashSet[ElementId]())
        unused_ids = set(_eid(x) for x in unused)
    except Exception:
        return 0, 1
    ids = []
    for fam in _collect_class(the_doc, Family, not_types=False):
        if _eid(fam) not in target_ids:
            continue
        try:
            ids.extend(
                sid for sid in fam.GetFamilySymbolIds() if _eid(sid) in unused_ids
            )
        except Exception:
            continue
    return _delete_ids(the_doc, ids)


def _purify_nested_families(fam_doc, depth, max_depth, ancestry):
    totals = {"deleted": 0, "failed": 0, "nested": 0, "nested_failed": 0}
    if depth >= max_depth:
        return totals
    nested = []
    for fam in _collect_class(fam_doc, Family, not_types=False):
        try:
            key = _safe_str(fam.Name).lower()
            if fam.IsInPlace or key in ancestry:
                continue
            if hasattr(fam, "IsEditable") and not fam.IsEditable:
                continue
            nested.append((key, fam.Id))
        except Exception:
            continue
    nested.sort(key=lambda item: item[0])
    for key, family_id in nested:
        child_doc = None
        try:
            child = fam_doc.GetElement(family_id)
            if child is None:
                continue
            child_doc = fam_doc.EditFamily(child)
            child_stats = purify_family_document(
                child_doc,
                aggressive=True,
                delete_type_names=None,
                depth=depth + 1,
                max_depth=max_depth,
                ancestry=set(ancestry) | set([key]),
            )
            loaded = child_doc.LoadFamily(fam_doc, OverwriteFamilyLoadOptions())
            if loaded is None:
                raise RuntimeError("nested LoadFamily returned no family")
            totals["deleted"] += int(child_stats.get("deleted") or 0)
            totals["failed"] += int(child_stats.get("failed") or 0)
            totals["nested"] += 1 + int(child_stats.get("nested") or 0)
            totals["nested_failed"] += int(child_stats.get("nested_failed") or 0)
        except Exception as ex:
            totals["failed"] += 1
            totals["nested_failed"] += 1
            log("    nested family purify failed '{}': {}".format(key, ex))
        finally:
            try:
                if child_doc is not None:
                    child_doc.Close(False)
            except Exception:
                pass
    return totals


def purify_family_document(
    fam_doc,
    aggressive=False,
    delete_type_names=None,
    depth=0,
    max_depth=0,
    ancestry=None,
):
    """Apply conservative or detached-only aggressive Family Purify cleanup."""
    totals = {"deleted": 0, "failed": 0, "nested": 0, "nested_failed": 0}

    def _run(name, fn):
        result = _family_txn(fam_doc, name, lambda d: fn(d))
        totals["deleted"] += int(result[0] or 0)
        totals["failed"] += int(result[1] or 0)

    group = TransactionGroup(fam_doc, "RBP Family Purify analysis")
    group.Start()
    try:
        if aggressive:
            _run("Purify non-core views", _purify_non_core_views)
            if delete_type_names:
                _run(
                    "Purify unused family types",
                    lambda d: _purify_family_types(d, delete_type_names),
                )
        _run("Purify imports and images", _purify_imports_images)
        _run("Purify unused subcategories", _purify_unused_subcategories)
        _run("Purify patterns and assets", _purify_unused_patterns_assets)
        _run("Purify unused materials", _purify_unused_materials)
        if aggressive:
            _run(
                "Purge unused family resources",
                lambda d: (_purge_via_api(d, 10) or 0, 0),
            )
        status = group.Assimilate()
        if status != TransactionStatus.Committed:
            raise RuntimeError("family transaction group returned {}".format(status))
    except Exception:
        try:
            group.RollBack()
        except Exception:
            pass
        raise
    if aggressive and max_depth > depth:
        owner_name = ""
        try:
            owner_name = _safe_str(fam_doc.OwnerFamily.Name).lower()
        except Exception:
            pass
        lineage = set(ancestry or [])
        if owner_name:
            lineage.add(owner_name)
        nested = _purify_nested_families(fam_doc, depth, max_depth, lineage)
        totals["deleted"] += int(nested.get("deleted") or 0)
        totals["failed"] += int(nested.get("failed") or 0)
        totals["nested"] += int(nested.get("nested") or 0)
        totals["nested_failed"] += int(nested.get("nested_failed") or 0)
    return totals


def _unpin(el):
    try:
        if getattr(el, "Pinned", False):
            el.Pinned = False
    except Exception:
        pass


def _delete_ids(the_doc, ids):
    """Delete element ids in chunks. Returns (deleted_count, failed_count)."""
    unique = []
    seen = set()
    for eid in ids:
        if eid is None:
            continue
        try:
            key = _id_value(eid)
        except Exception:
            continue
        if key in seen or key < 0:
            continue
        seen.add(key)
        unique.append(eid)
    if not unique:
        return 0, 0

    chunk_size = 400
    deleted_n = 0
    failed_n = 0
    i = 0
    while i < len(unique):
        chunk = unique[i : i + chunk_size]
        i += chunk_size
        try:
            net = NetList[ElementId](chunk)
            deleted = the_doc.Delete(net)
            count = deleted.Count if deleted is not None else len(chunk)
            deleted_n += int(count)
            continue
        except Exception:
            pass
        for eid in chunk:
            try:
                el = the_doc.GetElement(eid)
                if el is not None:
                    _unpin(el)
                the_doc.Delete(eid)
                deleted_n += 1
            except Exception:
                failed_n += 1
    return deleted_n, failed_n


def _collect_class(the_doc, cls, not_types=True):
    if cls is None:
        return []
    try:
        col = FilteredElementCollector(the_doc).OfClass(cls)
        if not_types:
            col = col.WhereElementIsNotElementType()
        return list(col.ToElements())
    except Exception:
        return []


def _is_level_element(el):
    try:
        if isinstance(el, Level):
            return True
    except Exception:
        pass
    try:
        cat = el.Category
        if cat is not None and _id_value(cat.Id) == int(BuiltInCategory.OST_Levels):
            return True
    except Exception:
        pass
    return False


def _is_grid_element(el):
    try:
        if Grid is not None and isinstance(el, Grid):
            return True
    except Exception:
        pass
    try:
        cat = el.Category
        if cat is not None and _id_value(cat.Id) == int(BuiltInCategory.OST_Grids):
            return True
    except Exception:
        pass
    return False


def _phase_key(el, attr):
    """IntegerElementId for CreatedPhaseId / DemolishedPhaseId, or None if unset."""
    try:
        cid = getattr(el, attr, None)
        if cid is None:
            return None
        n = _id_value(cid)
        if n < 0:
            return None
        return n
    except Exception:
        return None


def list_phases(the_doc):
    if Phase is None:
        return []
    phases = _collect_class(the_doc, Phase, not_types=False)

    def _seq(p):
        try:
            return int(p.SequenceNumber)
        except Exception:
            return 0

    return sorted(phases, key=_seq)


def _iter_model_instances(the_doc):
    """Document-wide CategoryType.Model instances, skipping levels/grids."""
    out = []
    for el in FilteredElementCollector(the_doc).WhereElementIsNotElementType().ToElements():
        try:
            cat = el.Category
            if cat is None or cat.CategoryType != CategoryType.Model:
                continue
            if _is_level_element(el) or _is_grid_element(el):
                continue
            out.append(el)
        except Exception:
            continue
    return out


def _is_phase_step(step_id):
    return _safe_str(step_id).startswith("phase_")


def _step_touches_keep_view(step_id):
    if step_id in ("views", "templates_filters", "geometry", "design_options"):
        return True
    return _is_phase_step(step_id)


def _filter_purge_ids(the_doc, ids):
    """Drop levels and grids from a purge id list (PDtool)."""
    out = []
    for eid in ids:
        try:
            el = the_doc.GetElement(eid)
            if el is None:
                continue
            if _is_level_element(el) or _is_grid_element(el):
                continue
            out.append(eid)
        except Exception:
            out.append(eid)
    return out


# ---------------------------------------------------------------------------
# Inventory (BIT Model Check substitute)
# ---------------------------------------------------------------------------


def _family_instance_counts(the_doc):
    counts = {}
    for fi in _collect_class(the_doc, FamilyInstance):
        try:
            fam = fi.Symbol.Family
            key = _eid(fam)
            if key is None:
                continue
            counts[key] = counts.get(key, 0) + 1
        except Exception:
            continue
    return counts


def _is_do_not_use(name):
    text = _safe_str(name).lower()
    if not text:
        return False
    if "anv" in text and "ej" in text:
        return True
    if "do not use" in text or "do_not_use" in text or "donotuse" in text:
        return True
    return False


def _index_exported_rfas(export_dir):
    """Map lowercase family name -> {path, bytes} from a folder of .rfa files."""
    index = {}
    if not export_dir or not os.path.isdir(export_dir):
        return index
    try:
        names = os.listdir(export_dir)
    except Exception:
        return index
    for name in names:
        if not name.lower().endswith(".rfa"):
            continue
        path = os.path.join(export_dir, name)
        if not os.path.isfile(path):
            continue
        base = os.path.splitext(name)[0]
        key = _safe_str(base).lower()
        index[key] = {"path": path, "bytes": _file_bytes(path), "name": _safe_str(base)}
    return index


def annotate_families_with_exports(rows, export_dir):
    """Flag do-not-use names and attach .rfa sizes from an exported family folder."""
    index = _index_exported_rfas(export_dir)
    for row in rows or []:
        name = _safe_str(row.get("name"))
        row["do_not_use"] = _is_do_not_use(name)
        info = index.get(name.lower()) if name else None
        row["in_export_folder"] = bool(info)
        if info:
            row["exported_rfa"] = info.get("path")
            if row.get("rfa_bytes") is None:
                row["rfa_bytes"] = info.get("bytes")
                row["rfa_mb"] = _mb(row["rfa_bytes"])
        else:
            row["exported_rfa"] = ""
    rows.sort(
        key=lambda r: (
            -int(bool(r.get("do_not_use"))),
            -(r.get("rfa_bytes") or 0),
            -int(r.get("instance_count") or 0),
            _safe_str(r.get("name")).lower(),
        )
    )
    return rows


def annotate_families_with_ranking_csv(rows, ranking_csv):
    """Attach prior measured RFA sizes without requiring the exported RFA folder."""
    if not ranking_csv or not os.path.isfile(ranking_csv):
        return rows
    ranked = {}
    try:
        with open(ranking_csv, "rb") as stream:
            for item in csv.DictReader(stream):
                name = _safe_str(item.get("name")).lower()
                if not name:
                    continue
                try:
                    size = int(item.get("rfa_bytes") or 0)
                except Exception:
                    size = 0
                if size > 0:
                    ranked[name] = size
    except Exception as ex:
        log("  family ranking CSV failed: {}".format(ex))
        return rows
    matched = 0
    for row in rows or []:
        size = ranked.get(_safe_str(row.get("name")).lower())
        if not size:
            continue
        row["rfa_bytes"] = size
        row["rfa_mb"] = _mb(size)
        row["ranking_csv"] = ranking_csv
        matched += 1
    log("  Family ranking CSV matched {}/{} loaded families".format(matched, len(rows or [])))
    return rows


def unmatched_exported_rfas(rows, export_dir):
    """Exported .rfa files whose names are not loaded in the model."""
    index = _index_exported_rfas(export_dir)
    loaded = set()
    for row in rows or []:
        name = _safe_str(row.get("name")).lower()
        if name:
            loaded.add(name)
    extra = []
    for key, info in index.items():
        if key in loaded:
            continue
        extra.append(
            {
                "name": info.get("name"),
                "rfa_bytes": info.get("bytes"),
                "rfa_mb": _mb(info.get("bytes")),
                "do_not_use": _is_do_not_use(info.get("name")),
                "path": info.get("path"),
            }
        )
    extra.sort(key=lambda r: (-(r.get("rfa_bytes") or 0), _safe_str(r.get("name")).lower()))
    return extra


def _family_category_rollup(families):
    buckets = {}
    for fam in families or []:
        cat = _safe_str(fam.get("category")) or "(none)"
        bucket = buckets.get(cat)
        if bucket is None:
            bucket = {
                "category": cat,
                "families": 0,
                "instances": 0,
                "rfa_bytes": 0,
                "do_not_use": 0,
            }
            buckets[cat] = bucket
        bucket["families"] += 1
        bucket["instances"] += int(fam.get("instance_count") or 0)
        bucket["rfa_bytes"] += int(fam.get("rfa_bytes") or 0)
        if fam.get("do_not_use"):
            bucket["do_not_use"] += 1
    rows = list(buckets.values())
    rows.sort(key=lambda r: (-r["rfa_bytes"], -r["instances"], r["category"].lower()))
    return rows


def write_family_csv(path, families):
    lines = [
        "name,category,is_in_place,instance_count,type_count,do_not_use,in_export_folder,"
        "rfa_mb,rfa_bytes,purify_status,purify_before_mb,purify_after_mb,purify_saved_mb,"
        "purify_reloaded,purify_removed,purify_failed,purify_nested,purify_nested_failed,"
        "purify_error"
    ]

    def _cell(v):
        text = "" if v is None else _safe_str(v)
        if "," in text or '"' in text:
            text = '"' + text.replace('"', '""') + '"'
        return text

    for fam in families or []:
        lines.append(
            ",".join(
                [
                    _cell(fam.get("name")),
                    _cell(fam.get("category")),
                    _cell(fam.get("is_in_place")),
                    _cell(fam.get("instance_count")),
                    _cell(fam.get("type_count")),
                    _cell(fam.get("do_not_use")),
                    _cell(fam.get("in_export_folder")),
                    _cell(fam.get("rfa_mb")),
                    _cell(fam.get("rfa_bytes")),
                    _cell(fam.get("purify_status")),
                    _cell(_mb(fam.get("purify_before_bytes"))),
                    _cell(_mb(fam.get("purify_after_bytes"))),
                    _cell(fam.get("purify_saved_mb")),
                    _cell(fam.get("purify_reloaded")),
                    _cell(fam.get("purify_removed")),
                    _cell(fam.get("purify_failed")),
                    _cell(fam.get("purify_nested")),
                    _cell(fam.get("purify_nested_failed")),
                    _cell(fam.get("purify_error")),
                ]
            )
        )
    _write_text(path, "\n".join(lines) + "\n")


def write_family_markdown(path, model_path, inventory, families, extra_exports, export_dir):
    placed = [f for f in (families or []) if int(f.get("instance_count") or 0) > 0]
    unused = [
        f
        for f in (families or [])
        if int(f.get("instance_count") or 0) == 0 and not f.get("is_in_place")
    ]
    dnu = [f for f in (families or []) if f.get("do_not_use")]
    dnu_placed = [f for f in dnu if int(f.get("instance_count") or 0) > 0]
    dnu_unused = [f for f in dnu if int(f.get("instance_count") or 0) == 0]
    matched = [f for f in (families or []) if f.get("in_export_folder")]
    dnu_bytes = sum((f.get("rfa_bytes") or 0) for f in dnu)
    placed_dnu_bytes = sum((f.get("rfa_bytes") or 0) for f in dnu_placed)
    matched_bytes = sum((f.get("rfa_bytes") or 0) for f in matched)
    extra_bytes = sum((f.get("rfa_bytes") or 0) for f in (extra_exports or []))
    extra_dnu = [f for f in (extra_exports or []) if f.get("do_not_use")]

    lines = []
    lines.append("# Family analytics")
    lines.append("")
    lines.append("- Model: `{}`".format(_safe_str(model_path)))
    lines.append("- Model size: **{} MB**".format(_mb(inventory.get("original_bytes") if inventory else None)))
    lines.append("- Exported .rfa folder: `{}`".format(_safe_str(export_dir) or "(not set)"))
    lines.append(
        "- Loaded families: **{}** ({} with instances, {} unused loadable, {} in-place)".format(
            len(families or []),
            len(placed),
            len(unused),
            (inventory or {}).get("in_place_families"),
        )
    )
    lines.append(
        "- Matched to export folder: {} families ({} MB of .rfa; nested content can overlap)".format(
            len(matched), _mb(matched_bytes)
        )
    )
    lines.append(
        "- Do-not-use name (`ANVÄND EJ` / `X_ANVÄND EJ` / `Z_ANVÄND EJ`): **{}** loaded, {} still placed, {} zero instances".format(
            len(dnu), len(dnu_placed), len(dnu_unused)
        )
    )
    lines.append(
        "- Placed do-not-use .rfa: {} MB across {} instances — Purge Unused cannot remove these".format(
            _mb(placed_dnu_bytes),
            sum(int(f.get("instance_count") or 0) for f in dnu_placed),
        )
    )
    if extra_exports:
        lines.append(
            "- Exported .rfa not loaded in this model: {} files ({} MB), {} do-not-use".format(
                len(extra_exports), _mb(extra_bytes), len(extra_dnu)
            )
        )
    lines.append("")
    lines.append("Purge Unused only deletes **unused definitions**. Families that still have instances stay, even if the name says do-not-use.")
    lines.append("")

    if dnu_placed:
        lines.append("## Do-not-use families still placed")
        lines.append("")
        lines.append("| Family | Category | Instances | Types | RFA MB | In export folder |")
        lines.append("| --- | --- | ---: | ---: | ---: | --- |")
        for fam in dnu_placed:
            lines.append(
                "| {} | {} | {} | {} | {} | {} |".format(
                    _safe_str(fam.get("name")).replace("|", "/"),
                    _safe_str(fam.get("category")).replace("|", "/"),
                    fam.get("instance_count"),
                    fam.get("type_count"),
                    fam.get("rfa_mb") if fam.get("rfa_mb") is not None else "",
                    fam.get("in_export_folder"),
                )
            )
        lines.append("")

    if dnu_unused:
        lines.append("## Do-not-use families with zero instances")
        lines.append("")
        lines.append("| Family | Category | Types | RFA MB | In export folder |")
        lines.append("| --- | --- | ---: | ---: | --- |")
        for fam in dnu_unused:
            lines.append(
                "| {} | {} | {} | {} | {} |".format(
                    _safe_str(fam.get("name")).replace("|", "/"),
                    _safe_str(fam.get("category")).replace("|", "/"),
                    fam.get("type_count"),
                    fam.get("rfa_mb") if fam.get("rfa_mb") is not None else "",
                    fam.get("in_export_folder"),
                )
            )
        lines.append("")

    by_size = sorted(
        [f for f in (families or []) if f.get("rfa_bytes")],
        key=lambda r: -(r.get("rfa_bytes") or 0),
    )
    if by_size:
        lines.append("## Largest loaded families (by exported .rfa)")
        lines.append("")
        lines.append("| Family | Category | Instances | RFA MB | Do-not-use |")
        lines.append("| --- | --- | ---: | ---: | --- |")
        for fam in by_size[:40]:
            lines.append(
                "| {} | {} | {} | {} | {} |".format(
                    _safe_str(fam.get("name")).replace("|", "/"),
                    _safe_str(fam.get("category")).replace("|", "/"),
                    fam.get("instance_count"),
                    fam.get("rfa_mb"),
                    fam.get("do_not_use"),
                )
            )
        lines.append("")

    by_count = sorted(
        families or [],
        key=lambda r: (-int(r.get("instance_count") or 0), _safe_str(r.get("name")).lower()),
    )
    lines.append("## Most placed families")
    lines.append("")
    lines.append("| Family | Category | Instances | Types | RFA MB | Do-not-use |")
    lines.append("| --- | --- | ---: | ---: | ---: | --- |")
    for fam in by_count[:40]:
        if int(fam.get("instance_count") or 0) <= 0:
            continue
        lines.append(
            "| {} | {} | {} | {} | {} | {} |".format(
                _safe_str(fam.get("name")).replace("|", "/"),
                _safe_str(fam.get("category")).replace("|", "/"),
                fam.get("instance_count"),
                fam.get("type_count"),
                fam.get("rfa_mb") if fam.get("rfa_mb") is not None else "",
                fam.get("do_not_use"),
            )
        )
    lines.append("")

    cats = _family_category_rollup(families)
    if cats:
        lines.append("## By category")
        lines.append("")
        lines.append("| Category | Families | Instances | RFA MB | Do-not-use |")
        lines.append("| --- | ---: | ---: | ---: | ---: |")
        for cat in cats[:30]:
            lines.append(
                "| {} | {} | {} | {} | {} |".format(
                    _safe_str(cat.get("category")).replace("|", "/"),
                    cat.get("families"),
                    cat.get("instances"),
                    _mb(cat.get("rfa_bytes")),
                    cat.get("do_not_use"),
                )
            )
        lines.append("")

    if extra_exports:
        lines.append("## Exported .rfa not loaded in the model")
        lines.append("")
        lines.append("| Family | RFA MB | Do-not-use |")
        lines.append("| --- | ---: | --- |")
        for fam in extra_exports[:40]:
            lines.append(
                "| {} | {} | {} |".format(
                    _safe_str(fam.get("name")).replace("|", "/"),
                    fam.get("rfa_mb"),
                    fam.get("do_not_use"),
                )
            )
        lines.append("")

    _write_text(path, "\n".join(lines))


def collect_family_inventory(the_doc, measure_sizes, tmp_dir):
    """List loaded families with instance counts (and optional .rfa sizes)."""
    inst_counts = _family_instance_counts(the_doc)
    rows = []
    for fam in _collect_class(the_doc, Family, not_types=False):
        try:
            fid = _eid(fam)
            is_inplace = bool(getattr(fam, "IsInPlace", False))
            cat_name = ""
            try:
                if fam.FamilyCategory is not None:
                    cat_name = _safe_str(fam.FamilyCategory.Name)
            except Exception:
                pass
            row = {
                "id": fid,
                "name": _safe_str(fam.Name),
                "category": cat_name,
                "is_in_place": is_inplace,
                "instance_count": int(inst_counts.get(fid, 0)),
                "type_count": 0,
                "rfa_bytes": None,
                "rfa_mb": None,
            }
            try:
                row["type_count"] = int(fam.GetFamilySymbolIds().Count)
            except Exception:
                pass
            if measure_sizes and not is_inplace:
                editable = True
                try:
                    if hasattr(fam, "IsEditable") and not fam.IsEditable:
                        editable = False
                except Exception:
                    pass
                if editable:
                    row["rfa_bytes"] = _measure_family_rfa_bytes(the_doc, fam, tmp_dir)
                    row["rfa_mb"] = _mb(row["rfa_bytes"])
            rows.append(row)
        except Exception:
            continue
    rows.sort(key=lambda r: (-(r.get("rfa_bytes") or 0), -r["instance_count"], r["name"].lower()))
    return rows


def _measure_family_rfa_bytes(the_doc, fam, tmp_dir):
    fam_doc = None
    path = None
    try:
        fam_doc = the_doc.EditFamily(fam)
        if not os.path.isdir(tmp_dir):
            os.makedirs(tmp_dir)
        path = os.path.join(tmp_dir, _safe_filename(fam.Name) + ".rfa")
        if os.path.isfile(path):
            try:
                os.remove(path)
            except Exception:
                pass
        fam_doc.SaveAs(path)
        size = _file_bytes(path)
        return size
    except Exception as ex:
        log("  family size skip '{}': {}".format(_safe_str(fam.Name), ex))
        return None
    finally:
        try:
            if fam_doc is not None:
                fam_doc.Close(False)
        except Exception:
            pass
        if path and os.path.isfile(path):
            try:
                os.remove(path)
            except Exception:
                pass


def collect_inventory(the_doc, model_path):
    inv = {
        "original_path": model_path,
        "original_bytes": _file_bytes(model_path),
        "is_workshared": bool(getattr(the_doc, "IsWorkshared", False)),
        "title": _safe_str(the_doc.Title),
        "warnings_total": 0,
        "views_total": 0,
        "view_templates": 0,
        "views_not_on_sheets": 0,
        "unused_view_templates": 0,
        "unused_view_filters": 0,
        "view_filters": 0,
        "views_by_type": {},
        "sheets": 0,
        "levels": 0,
        "revit_links": [],
        "cad_imports": 0,
        "cad_links": 0,
        "cad_model_links": 0,
        "image_types": 0,
        "image_imports": 0,
        "image_links": 0,
        "point_clouds": 0,
        "in_place_families": 0,
        "in_place_instances": 0,
        "loaded_families": 0,
        "family_instances": 0,
        "groups": 0,
        "unplaced_model_groups": 0,
        "unplaced_detail_groups": 0,
        "materials": 0,
        "design_options": 0,
        "design_option_sets": 0,
        "phases": [],
        "unphased_model": 0,
        "filled_regions": 0,
        "unused_text_types": 0,
        "unused_dimension_types": 0,
        "worksets": 0,
        "rooms_placed": 0,
        "rooms_unplaced": 0,
        "rooms_not_enclosed": 0,
        "parameter_elements": 0,
        "shared_parameter_elements": 0,
    }
    try:
        inv["warnings_total"] = len(list(the_doc.GetWarnings()))
    except Exception:
        pass

    sheeted_view_ids = set()
    if Viewport is not None:
        for vp in _collect_class(the_doc, Viewport):
            try:
                sheeted_view_ids.add(_eid(vp.ViewId) if hasattr(vp, "ViewId") else _eid(vp))
            except Exception:
                try:
                    sheeted_view_ids.add(_id_value(vp.ViewId))
                except Exception:
                    continue

    used_template_ids = set()
    used_filter_ids = set()
    for v in _collect_class(the_doc, View):
        try:
            if getattr(v, "IsTemplate", False):
                inv["view_templates"] += 1
                continue
            inv["views_total"] += 1
            vt = _safe_str(v.ViewType)
            inv["views_by_type"][vt] = inv["views_by_type"].get(vt, 0) + 1
            try:
                if v.ViewType != ViewType.DrawingSheet and _eid(v) not in sheeted_view_ids:
                    inv["views_not_on_sheets"] += 1
            except Exception:
                pass
            try:
                tid = _eid(v.ViewTemplateId)
                if tid:
                    used_template_ids.add(tid)
            except Exception:
                pass
            try:
                for fid in v.GetFilters():
                    used_filter_ids.add(_eid(fid))
            except Exception:
                pass
        except Exception:
            continue

    unused_tmpl = 0
    for v in _collect_class(the_doc, View):
        try:
            if not getattr(v, "IsTemplate", False):
                continue
            if _eid(v) not in used_template_ids:
                unused_tmpl += 1
        except Exception:
            continue
    inv["unused_view_templates"] = unused_tmpl

    if ParameterFilterElement is not None:
        filters = _collect_class(the_doc, ParameterFilterElement, not_types=False)
        inv["view_filters"] = len(filters)
        unused_f = 0
        for flt in filters:
            if _eid(flt) not in used_filter_ids:
                unused_f += 1
        inv["unused_view_filters"] = unused_f

    inv["sheets"] = len(_collect_class(the_doc, ViewSheet))
    inv["levels"] = len(_collect_class(the_doc, Level))

    for lt in _collect_class(the_doc, RevitLinkType, not_types=False):
        try:
            inv["revit_links"].append(
                {
                    "name": _safe_str(lt.Name),
                    "loaded": bool(getattr(lt, "IsLoaded", False)),
                    "nested": bool(getattr(lt, "IsNestedLink", False)),
                }
            )
        except Exception:
            pass

    for imp in _collect_class(the_doc, ImportInstance):
        try:
            if getattr(imp, "IsLinked", False):
                inv["cad_links"] += 1
            else:
                inv["cad_imports"] += 1
        except Exception:
            inv["cad_imports"] += 1
    if CADLinkType is not None:
        inv["cad_model_links"] = len(_collect_class(the_doc, CADLinkType, not_types=False))

    inv["image_types"] = 0
    if ImageType is not None:
        images = _collect_class(the_doc, ImageType, not_types=False)
        inv["image_types"] = len(images)
        for img in images:
            src = None
            try:
                src = getattr(img, "Source", None)
            except Exception:
                src = None
            src_name = _safe_str(src).lower()
            if "link" in src_name:
                inv["image_links"] += 1
            elif src_name:
                inv["image_imports"] += 1
    if PointCloudType is not None:
        inv["point_clouds"] = len(_collect_class(the_doc, PointCloudType, not_types=False))

    fams = _collect_class(the_doc, Family, not_types=False)
    inv["loaded_families"] = len(fams)
    inv["in_place_families"] = sum(1 for f in fams if getattr(f, "IsInPlace", False))

    insts = _collect_class(the_doc, FamilyInstance)
    inv["family_instances"] = len(insts)
    inv["in_place_instances"] = sum(1 for fi in insts if getattr(fi, "IsInPlace", False))

    try:
        from Autodesk.Revit.DB import Material

        inv["materials"] = len(_collect_class(the_doc, Material, not_types=False))
    except Exception:
        pass

    if Group is not None:
        inv["groups"] = len(_collect_class(the_doc, Group))
    if GroupType is not None:
        for gt in _collect_class(the_doc, GroupType, not_types=False):
            try:
                n = int(gt.Groups.Size)
            except Exception:
                continue
            if n > 0:
                continue
            cat_id = 0
            try:
                cat_id = _id_value(gt.Category.Id)
            except Exception:
                pass
            if cat_id == int(BuiltInCategory.OST_IOSDetailGroups):
                inv["unplaced_detail_groups"] += 1
            else:
                inv["unplaced_model_groups"] += 1

    if DesignOption is not None:
        opts = _collect_class(the_doc, DesignOption, not_types=False)
        inv["design_options"] = len(opts)
        sets = set()
        for opt in opts:
            try:
                if BuiltInParameter is not None:
                    p = opt.get_Parameter(BuiltInParameter.OPTION_SET_ID)
                    if p is not None:
                        sets.add(_eid(p.AsElementId()))
            except Exception:
                pass
        inv["design_option_sets"] = len(sets) if sets else (1 if opts else 0)

    phase_rows = []
    by_phase_id = {}
    for p in list_phases(the_doc):
        seq = 0
        try:
            seq = int(p.SequenceNumber)
        except Exception:
            seq = 0
        row = {
            "name": _safe_str(p.Name),
            "sequence": seq,
            "id": _eid(p),
            "created": 0,
            "demolished": 0,
        }
        phase_rows.append(row)
        if row["id"] is not None:
            by_phase_id[row["id"]] = row
    inv["phases"] = phase_rows
    unphased = 0
    for el in _iter_model_instances(the_doc):
        ck = _phase_key(el, "CreatedPhaseId")
        dk = _phase_key(el, "DemolishedPhaseId")
        if ck is None:
            unphased += 1
        elif ck in by_phase_id:
            by_phase_id[ck]["created"] += 1
        if dk is not None and dk in by_phase_id:
            by_phase_id[dk]["demolished"] += 1
    inv["unphased_model"] = unphased

    if FilledRegion is not None:
        inv["filled_regions"] = len(_collect_class(the_doc, FilledRegion))

    if TextNote is not None and TextNoteType is not None:
        used = set()
        for tn in _collect_class(the_doc, TextNote):
            try:
                used.add(_eid(tn.GetTypeId()))
            except Exception:
                continue
        unused_t = 0
        for t in _collect_class(the_doc, TextNoteType, not_types=False):
            if _eid(t) not in used:
                unused_t += 1
        inv["unused_text_types"] = unused_t

    if Dimension is not None and DimensionType is not None:
        used = set()
        for dim in _collect_class(the_doc, Dimension):
            try:
                used.add(_eid(dim.GetTypeId()))
            except Exception:
                continue
        unused_d = 0
        for t in _collect_class(the_doc, DimensionType, not_types=False):
            if _eid(t) not in used:
                unused_d += 1
        inv["unused_dimension_types"] = unused_d

    if FilteredWorksetCollector is not None and WorksetKind is not None:
        try:
            if the_doc.IsWorkshared:
                inv["worksets"] = len(
                    list(FilteredWorksetCollector(the_doc).OfKind(WorksetKind.UserWorkset).ToWorksets())
                )
        except Exception:
            pass

    if Room is not None:
        for room in _collect_class(the_doc, Room):
            try:
                loc = getattr(room, "Location", None)
                area = 0.0
                try:
                    area = float(room.Area)
                except Exception:
                    pass
                if loc is None:
                    inv["rooms_unplaced"] += 1
                elif area <= 0:
                    inv["rooms_not_enclosed"] += 1
                else:
                    inv["rooms_placed"] += 1
            except Exception:
                continue

    if ParameterElement is not None:
        inv["parameter_elements"] = len(_collect_class(the_doc, ParameterElement, not_types=False))
    if SharedParameterElement is not None:
        inv["shared_parameter_elements"] = len(
            _collect_class(the_doc, SharedParameterElement, not_types=False)
        )

    return inv


# ---------------------------------------------------------------------------
# Keep-view + Save As
# ---------------------------------------------------------------------------


def resolve_keep_view(the_doc):
    """Return the dedicated analysis 3D by name, or None. Never pick a random 3D."""
    for v in _collect_class(the_doc, View3D):
        if getattr(v, "IsTemplate", False):
            continue
        try:
            if _safe_str(v.Name) == KEEP_VIEW_NAME:
                return v
        except Exception:
            continue
    return None


def create_keep_view(the_doc):
    """Always create/reuse RBP_SizeAnalysis_3D. Do not reuse {3D} or any other view."""
    existing = resolve_keep_view(the_doc)
    if existing is not None:
        return existing

    vft = None
    for t in _collect_class(the_doc, ViewFamilyType, not_types=False):
        try:
            if t.ViewFamily == ViewFamily.ThreeDimensional:
                vft = t
                break
        except Exception:
            continue
    if vft is None:
        raise RuntimeError("No 3D ViewFamilyType; cannot create analysis view")

    created = []

    def _create(d):
        view = View3D.CreateIsometric(d, vft.Id)
        try:
            view.Name = KEEP_VIEW_NAME
        except Exception:
            pass
        created.append(view)
        return view

    ok, result = _run_transaction(the_doc, "Create size-analysis 3D view", _create)
    if not ok:
        raise RuntimeError("Could not create 3D view: {}".format(result))
    named = resolve_keep_view(the_doc)
    if named is not None:
        return named
    if created:
        return created[0]
    raise RuntimeError("Keep 3D view was created but could not be resolved by name")


def find_or_create_keep_view(the_doc):
    return create_keep_view(the_doc)


def _activate_view(the_doc, view):
    try:
        uidoc = uiapp.ActiveUIDocument
        if uidoc is not None and view is not None:
            uidoc.ActiveView = view
            return True
    except Exception:
        pass
    return False


def save_compact_snapshot(the_doc, path, keep_view):
    """Compact Save As to path. SaveAs always compacts. Returns bytes or None."""
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    if os.path.isfile(path):
        try:
            os.remove(path)
        except Exception:
            pass

    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    opts.Compact = True
    try:
        opts.MaximumBackups = 1
    except Exception:
        pass
    if keep_view is not None:
        try:
            opts.PreviewViewId = keep_view.Id
        except Exception:
            keep_view = resolve_keep_view(the_doc)
            if keep_view is not None:
                try:
                    opts.PreviewViewId = keep_view.Id
                except Exception:
                    pass
    try:
        if the_doc.IsWorkshared:
            ws = WorksharingSaveAsOptions()
            ws.SaveAsCentral = True
            opts.SetWorksharingOptions(ws)
    except Exception:
        pass

    the_doc.SaveAs(path, opts)
    size = _file_bytes(path)
    log("  saved {} ({} MB)".format(path, _mb(size)))
    return size


# ---------------------------------------------------------------------------
# Reduction steps
# ---------------------------------------------------------------------------


def _purge_via_api(the_doc, max_passes):
    deleted = 0
    if not hasattr(the_doc, "GetUnusedElements"):
        return None
    for _ in range(max_passes):
        try:
            unused = the_doc.GetUnusedElements(HashSet[ElementId]())
        except Exception:
            try:
                unused = the_doc.GetUnusedElements(NetList[ElementId]())
            except Exception as ex:
                log("  GetUnusedElements failed: {}".format(ex))
                return deleted if deleted else None
        if unused is None:
            break
        try:
            count = unused.Count
        except Exception:
            count = len(list(unused))
        if count <= 0:
            break
        ids = _filter_purge_ids(the_doc, list(unused))
        if not ids:
            break
        n, _failed = _delete_ids(the_doc, ids)
        deleted += n
        if n <= 0:
            break
    return deleted


def _purge_via_reflection(the_doc):
    method_names = (
        "GetUnusedFamilies",
        "GetUnusedSymbols",
        "GetUnusedImportCategories",
        "GetUnusedMaterials",
        "GetUnusedAppearances",
        "GetUnusedStructures",
        "GetUnusedThermals",
    )
    ids = []
    doc_type = the_doc.GetType()
    flags = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public
    for name in method_names:
        try:
            mi = doc_type.GetMethod(name, flags)
            if mi is None:
                continue
            result = mi.Invoke(the_doc, None)
            if result is None:
                continue
            for eid in result:
                ids.append(eid)
        except Exception:
            continue
    if not ids:
        return 0
    n, _f = _delete_ids(the_doc, ids)
    return n


def _purge_unused_families_fallback(the_doc):
    used_symbols = set()
    for fi in _collect_class(the_doc, FamilyInstance):
        try:
            used_symbols.add(_eid(fi.GetTypeId()))
        except Exception:
            try:
                used_symbols.add(_eid(fi.Symbol))
            except Exception:
                continue
    ids = []
    for fam in _collect_class(the_doc, Family, not_types=False):
        if getattr(fam, "IsInPlace", False):
            continue
        try:
            symbol_ids = list(fam.GetFamilySymbolIds())
        except Exception:
            continue
        any_used = False
        unused_syms = []
        for sid in symbol_ids:
            if _eid(sid) in used_symbols:
                any_used = True
            else:
                unused_syms.append(sid)
        if not any_used:
            ids.append(fam.Id)
        else:
            ids.extend(unused_syms)
    n, _f = _delete_ids(the_doc, ids)
    return n


def purge_unused(the_doc, max_passes):
    api = _purge_via_api(the_doc, max_passes)
    if api is not None:
        extra = 0
        # Older leftover types after API purge: one more family-only pass
        extra = _purge_unused_families_fallback(the_doc)
        return {
            "deleted": int(api) + int(extra),
            "failed": 0,
            "notes": "GetUnusedElements until empty (levels/grids skipped) + unused-family fallback",
        }

    reflected = 0
    for _ in range(max_passes):
        n = _purge_via_reflection(the_doc)
        if n <= 0:
            break
        reflected += n
    fallback = _purge_unused_families_fallback(the_doc)
    return {
        "deleted": int(reflected) + int(fallback),
        "failed": 0,
        "notes": "pre-2024 reflection/family fallback (GetUnusedElements not available)",
    }


def remove_revit_links(the_doc):
    ids = [el.Id for el in _collect_class(the_doc, RevitLinkInstance)]
    ids.extend(el.Id for el in _collect_class(the_doc, RevitLinkType, not_types=False))
    d, f = _delete_ids(the_doc, ids)
    return {"deleted": d, "failed": f, "notes": "RevitLinkInstance + RevitLinkType"}


def remove_cad(the_doc):
    ids = [el.Id for el in _collect_class(the_doc, ImportInstance)]
    if CADLinkType is not None:
        ids.extend(el.Id for el in _collect_class(the_doc, CADLinkType, not_types=False))
    d, f = _delete_ids(the_doc, ids)
    return {"deleted": d, "failed": f, "notes": "ImportInstance + CADLinkType"}


def remove_unplaced_groups(the_doc):
    if GroupType is None:
        return {"deleted": 0, "failed": 0, "notes": "GroupType API not available"}
    ids = []
    model_n = 0
    detail_n = 0
    for gt in _collect_class(the_doc, GroupType, not_types=False):
        try:
            if int(gt.Groups.Size) > 0:
                continue
            cat_id = 0
            try:
                cat_id = _id_value(gt.Category.Id)
            except Exception:
                pass
            if cat_id == int(BuiltInCategory.OST_IOSDetailGroups):
                detail_n += 1
            else:
                model_n += 1
            ids.append(gt.Id)
        except Exception:
            continue
    d, f = _delete_ids(the_doc, ids)
    return {
        "deleted": d,
        "failed": f,
        "notes": "unplaced group types (model {}, detail {})".format(model_n, detail_n),
    }


def remove_unused_templates_and_filters(the_doc):
    used_template_ids = set()
    used_filter_ids = set()
    for v in _collect_class(the_doc, View):
        try:
            if getattr(v, "IsTemplate", False):
                continue
            try:
                tid = _eid(v.ViewTemplateId)
                if tid:
                    used_template_ids.add(tid)
            except Exception:
                pass
            try:
                for fid in v.GetFilters():
                    used_filter_ids.add(_eid(fid))
            except Exception:
                pass
        except Exception:
            continue

    ids = []
    tmpl_n = 0
    for v in _collect_class(the_doc, View):
        try:
            if not getattr(v, "IsTemplate", False):
                continue
            if _eid(v) in used_template_ids:
                continue
            ids.append(v.Id)
            tmpl_n += 1
        except Exception:
            continue

    filt_n = 0
    if ParameterFilterElement is not None:
        for flt in _collect_class(the_doc, ParameterFilterElement, not_types=False):
            try:
                if _eid(flt) in used_filter_ids:
                    continue
                ids.append(flt.Id)
                filt_n += 1
            except Exception:
                continue

    d, f = _delete_ids(the_doc, ids)
    return {
        "deleted": d,
        "failed": f,
        "notes": "unused view templates {}, unused filters {}".format(tmpl_n, filt_n),
    }


def _bound_parameter_ids(the_doc):
    bound = set()
    try:
        it = the_doc.ParameterBindings.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            defn = it.Key
            try:
                bound.add(_eid(defn.Id) if hasattr(defn, "Id") else _id_value(defn.Id))
            except Exception:
                try:
                    bound.add(_id_value(defn.Id))
                except Exception:
                    continue
    except Exception:
        pass
    return bound


def _definition_is_builtin(defn):
    try:
        if BuiltInParameter is None:
            return False
        bip = getattr(defn, "BuiltInParameter", None)
        if bip is None:
            return False
        if bip == BuiltInParameter.INVALID:
            return False
        return True
    except Exception:
        return False


def _param_has_any_value(the_doc, definition, binding):
    cats = []
    try:
        cats = list(binding.Categories)
    except Exception:
        return True
    if not cats:
        return True
    name = _safe_str(getattr(definition, "Name", ""))
    guid = None
    try:
        guid = definition.GUID
    except Exception:
        guid = None
    for cat in cats:
        try:
            col = FilteredElementCollector(the_doc).OfCategoryId(cat.Id).WhereElementIsNotElementType()
            for el in col:
                p = None
                try:
                    if guid is not None:
                        p = el.get_Parameter(guid)
                except Exception:
                    p = None
                if p is None and name:
                    try:
                        p = el.LookupParameter(name)
                    except Exception:
                        p = None
                try:
                    if p is not None and p.HasValue:
                        return True
                except Exception:
                    continue
        except Exception:
            continue
    return False


def remove_unused_project_parameters(the_doc):
    """Delete unbound ParameterElements and bound project params with no populated values."""
    if ParameterElement is None:
        return {"deleted": 0, "failed": 0, "notes": "ParameterElement API not available"}

    bound_ids = _bound_parameter_ids(the_doc)
    ids = []
    unbound_n = 0
    empty_n = 0

    param_els = _collect_class(the_doc, ParameterElement, not_types=False)
    pe_by_id = {}
    for pe in param_els:
        pe_by_id[_eid(pe)] = pe
        if _eid(pe) not in bound_ids:
            ids.append(pe.Id)
            unbound_n += 1

    # Bound HasValue walks every instance per parameter and pretouches family
    # documents; that native scan has crashed Revit 2025 mid-peel. Off by default.
    if _env_flag("RBP_SCAN_EMPTY_PARAMS", False):
        try:
            it = the_doc.ParameterBindings.ForwardIterator()
            it.Reset()
            to_remove = []
            while it.MoveNext():
                defn = it.Key
                binding = it.Current
                try:
                    if _definition_is_builtin(defn):
                        continue
                except Exception:
                    pass
                try:
                    if _param_has_any_value(the_doc, defn, binding):
                        continue
                except Exception:
                    continue
                to_remove.append(defn)
                empty_n += 1
            for defn in to_remove:
                try:
                    the_doc.ParameterBindings.Remove(defn)
                except Exception:
                    continue
                try:
                    pid = _eid(defn.Id) if hasattr(defn, "Id") else _id_value(defn.Id)
                    pe = pe_by_id.get(pid)
                    if pe is not None:
                        ids.append(pe.Id)
                except Exception:
                    continue
        except Exception as ex:
            log("  parameter binding scan skipped: {}".format(ex))
    else:
        log("  skipping empty-bound-parameter HasValue scan (set RBP_SCAN_EMPTY_PARAMS=1 to enable)")

    d, f = _delete_ids(the_doc, ids)
    return {
        "deleted": d,
        "failed": f,
        "notes": "unbound ParameterElements {}, empty bound params {}{}".format(
            unbound_n,
            empty_n,
            "" if _env_flag("RBP_SCAN_EMPTY_PARAMS", False) else " (HasValue scan skipped)",
        ),
    }


def remove_images_and_pointclouds(the_doc):
    ids = []
    notes = []
    if ImageType is not None:
        ids.extend(el.Id for el in _collect_class(the_doc, ImageType, not_types=False))
        notes.append("ImageType")
    if PointCloudInstance is not None:
        ids.extend(el.Id for el in _collect_class(the_doc, PointCloudInstance))
        notes.append("PointCloudInstance")
    if PointCloudType is not None:
        ids.extend(el.Id for el in _collect_class(the_doc, PointCloudType, not_types=False))
        notes.append("PointCloudType")
    d, f = _delete_ids(the_doc, ids)
    return {"deleted": d, "failed": f, "notes": ", ".join(notes)}


def remove_sheets(the_doc):
    ids = [el.Id for el in _collect_class(the_doc, ViewSheet)]
    d, f = _delete_ids(the_doc, ids)
    return {"deleted": d, "failed": f, "notes": "all sheets"}


def remove_views_except(the_doc, keep_view):
    _activate_view(the_doc, keep_view)
    keep_id = _eid(keep_view)
    ids = []
    templates = 0
    for v in _collect_class(the_doc, View):
        try:
            if _eid(v) == keep_id:
                continue
            if getattr(v, "IsTemplate", False):
                templates += 1
                continue
            if v.ViewType == ViewType.DrawingSheet:
                continue
            ids.append(v.Id)
        except Exception:
            continue
    d, f = _delete_ids(the_doc, ids)
    return {
        "deleted": d,
        "failed": f,
        "notes": "kept 3D view id {}; {} view templates left for later purge".format(
            keep_id, templates
        ),
    }


def remove_inplace_families(the_doc):
    ids = []
    for fi in _collect_class(the_doc, FamilyInstance):
        try:
            if fi.IsInPlace:
                ids.append(fi.Id)
        except Exception:
            continue
    for fam in _collect_class(the_doc, Family, not_types=False):
        try:
            if fam.IsInPlace:
                ids.append(fam.Id)
        except Exception:
            continue
    d, f = _delete_ids(the_doc, ids)
    return {"deleted": d, "failed": f, "notes": "in-place instances + families"}


def remove_unused_loaded_families(the_doc, max_passes):
    inst_counts = _family_instance_counts(the_doc)
    ids = []
    kept = 0
    for fam in _collect_class(the_doc, Family, not_types=False):
        try:
            if fam.IsInPlace:
                continue
            fid = _eid(fam)
            if inst_counts.get(fid, 0) > 0:
                kept += 1
                continue
            ids.append(fam.Id)
        except Exception:
            continue
    d, f = _delete_ids(the_doc, ids)
    purge = purge_unused(the_doc, max_passes)
    return {
        "deleted": d + int(purge.get("deleted") or 0),
        "failed": f,
        "notes": "deleted {} unused loaded families, kept {} with instances; {}".format(
            d, kept, purge.get("notes") or "purge"
        ),
    }


def _write_family_purify_checkpoint(path, rows):
    lines = [
        "family,status,reloaded,before_bytes,after_bytes,saved_bytes,saved_mb,"
        "removed,failed,nested,nested_failed,error"
    ]

    def _cell(value):
        text = "" if value is None else _safe_str(value)
        if "," in text or '"' in text or "\n" in text or "\r" in text:
            text = '"' + text.replace('"', '""').replace("\r", " ").replace("\n", " ") + '"'
        return text

    for row in rows or []:
        lines.append(
            ",".join(
                [
                    _cell(row.get("name")),
                    _cell(row.get("purify_status")),
                    _cell(row.get("purify_reloaded")),
                    _cell(row.get("purify_before_bytes")),
                    _cell(row.get("purify_after_bytes")),
                    _cell(row.get("purify_saved_bytes")),
                    _cell(row.get("purify_saved_mb")),
                    _cell(row.get("purify_removed")),
                    _cell(row.get("purify_failed")),
                    _cell(row.get("purify_nested")),
                    _cell(row.get("purify_nested_failed")),
                    _cell(row.get("purify_error")),
                ]
            )
        )
    _write_text(path, "\n".join(lines) + "\n")


def purify_loaded_families(
    the_doc,
    temp_dir,
    family_rows,
    max_families,
    aggressive=False,
    reload_min_bytes=0,
    family_offset=0,
):
    """Edit, safely purify, measure, and reload loadable families into a detached RVT."""
    if not os.path.isdir(temp_dir):
        os.makedirs(temp_dir)
    checkpoint = os.path.join(os.path.dirname(temp_dir), "family_purification.csv")
    by_id = {}
    by_name = {}
    for row in family_rows or []:
        by_id[row.get("id")] = row
        by_name[_safe_str(row.get("name")).lower()] = row

    unused_type_names = _project_unused_type_names_by_family(the_doc) if aggressive else {}
    candidates = []
    for fam in _collect_class(the_doc, Family, not_types=False):
        try:
            if fam.IsInPlace:
                continue
            if hasattr(fam, "IsEditable") and not fam.IsEditable:
                continue
            row = by_id.get(_eid(fam)) or by_name.get(_safe_str(fam.Name).lower()) or {}
            if aggressive and int(row.get("instance_count") or 0) <= 0:
                continue
            candidates.append((-(row.get("rfa_bytes") or 0), _safe_str(fam.Name).lower(), fam.Id))
        except Exception:
            continue
    candidates.sort()
    family_offset = max(0, int(family_offset or 0))
    if max_families > 0:
        candidates = candidates[family_offset : family_offset + max_families]
    elif family_offset:
        candidates = candidates[family_offset:]

    processed = []
    successful = 0
    skipped = 0
    failed_families = 0
    total_removed = 0
    total_delete_failed = 0
    total_before = 0
    total_after = 0
    reloaded = 0
    reloaded_saved = 0
    total_nested = 0
    total_nested_failed = 0
    processed_family_ids = []
    reloaded_family_ids = []

    log(
        "  Family Purify {} analysis: {} editable families{} at offset {}; checkpoint {}".format(
            "aggressive" if aggressive else "conservative",
            len(candidates),
            " (capped)" if max_families > 0 else "",
            family_offset,
            checkpoint,
        )
    )
    for pos, candidate in enumerate(candidates):
        fam_doc = None
        before_path = None
        after_path = None
        fam = the_doc.GetElement(candidate[2])
        if fam is None:
            skipped += 1
            continue
        name = _safe_str(fam.Name)
        row = by_id.get(_eid(fam)) or by_name.get(name.lower())
        if row is None:
            row = {"id": _eid(fam), "name": name}
            family_rows.append(row)
            by_id[row.get("id")] = row
            by_name[name.lower()] = row
        safe = "{}_{}".format(_safe_filename(name), _eid(fam))
        before_path = os.path.join(temp_dir, safe + "_before.rfa")
        after_path = os.path.join(temp_dir, safe + "_after.rfa")
        row["purify_status"] = "started"
        row["purify_aggressive"] = bool(aggressive)
        row["purify_reloaded"] = False
        row["purify_error"] = ""
        processed.append(row)
        try:
            _write_family_purify_checkpoint(checkpoint, processed)
        except Exception as ex:
            log("  family-purify checkpoint failed: {}".format(ex))
        try:
            fam_doc = the_doc.EditFamily(fam)
            before_bytes = _save_family_measure(fam_doc, before_path, False)
            stats = purify_family_document(
                fam_doc,
                aggressive=aggressive,
                delete_type_names=unused_type_names.get(_eid(fam)),
                max_depth=2 if aggressive else 0,
            )
            after_bytes = _save_family_measure(fam_doc, after_path, True)
            saved = max(0, int(before_bytes or 0) - int(after_bytes or 0))
            should_reload = not aggressive or saved >= int(reload_min_bytes or 0)
            if should_reload:
                loaded = fam_doc.LoadFamily(the_doc, OverwriteFamilyLoadOptions())
                if loaded is None:
                    raise RuntimeError("LoadFamily returned no family")
                reloaded += 1
                reloaded_saved += saved
                reloaded_family_ids.append(_eid(fam))
                row["purify_reloaded"] = True
            row["purify_status"] = "reloaded" if should_reload else "measured"
            row["purify_before_bytes"] = before_bytes
            row["purify_after_bytes"] = after_bytes
            row["purify_saved_bytes"] = saved
            row["purify_saved_mb"] = _mb(saved)
            row["purify_removed"] = int(stats.get("deleted") or 0)
            row["purify_failed"] = int(stats.get("failed") or 0)
            row["purify_nested"] = int(stats.get("nested") or 0)
            row["purify_nested_failed"] = int(stats.get("nested_failed") or 0)
            successful += 1
            processed_family_ids.append(_eid(fam))
            total_before += int(before_bytes or 0)
            total_after += int(after_bytes or 0)
            total_removed += int(stats.get("deleted") or 0)
            total_delete_failed += int(stats.get("failed") or 0)
            total_nested += int(stats.get("nested") or 0)
            total_nested_failed += int(stats.get("nested_failed") or 0)
        except Exception as ex:
            row["purify_status"] = "failed"
            row["purify_error"] = _safe_str(ex)
            failed_families += 1
            log("  family purify failed '{}': {}".format(name, ex))
        finally:
            try:
                if fam_doc is not None:
                    fam_doc.Close(False)
            except Exception:
                pass
            for path in (before_path, after_path):
                if path and os.path.isfile(path):
                    try:
                        os.remove(path)
                    except Exception:
                        pass
            try:
                _write_family_purify_checkpoint(checkpoint, processed)
            except Exception as ex:
                log("  family-purify checkpoint failed: {}".format(ex))
            if (pos + 1) % 10 == 0:
                log(
                    "  Family Purify {}/{}: {} ok, {} failed".format(
                        pos + 1, len(candidates), successful, failed_families
                    )
                )
                try:
                    System.GC.Collect()
                    System.GC.WaitForPendingFinalizers()
                except Exception:
                    pass

    saved_total = max(0, total_before - total_after)
    return {
        "deleted": total_removed,
        "failed": failed_families + total_delete_failed,
        "notes": (
            "{} purified, {} reloaded ({} MB measured saving), {} family failures, "
            "{} skipped; compact RFA total {} -> {} MB (potential saving {} MB); "
            "{} profile from Family Purify"
        ).format(
            successful,
            reloaded,
            _mb(reloaded_saved),
            failed_families,
            skipped,
            _mb(total_before),
            _mb(total_after),
            _mb(saved_total),
            "aggressive detached-only" if aggressive else "conservative analysis",
        ),
        "families_processed": successful,
        "families_failed": failed_families,
        "before_bytes": total_before,
        "after_bytes": total_after,
        "saved_bytes": saved_total,
        "families_reloaded": reloaded,
        "reloaded_saved_bytes": reloaded_saved,
        "nested_processed": total_nested,
        "nested_failed": total_nested_failed,
        "aggressive": bool(aggressive),
        "reload_min_bytes": int(reload_min_bytes or 0),
        "family_offset": family_offset,
        "processed_family_ids": processed_family_ids,
        "reloaded_family_ids": reloaded_family_ids,
        "checkpoint": checkpoint,
    }


def delete_model_elements(the_doc):
    """Document-wide model instances. Skip levels/grids. Do not depend on a live view."""
    ids = []
    for el in _iter_model_instances(the_doc):
        try:
            _unpin(el)
            ids.append(el.Id)
        except Exception:
            continue
    d, f = _delete_ids(the_doc, ids)
    return {
        "deleted": d,
        "failed": f,
        "notes": "leftover CategoryType.Model after per-phase peel (levels/grids skipped)",
    }


def delete_visible_3d_elements(the_doc, keep_view):
    return delete_model_elements(the_doc)


def remove_design_options(the_doc):
    """Delete non-primary options, then option sets so survivors land in the main model."""
    if DesignOption is None:
        return {"deleted": 0, "failed": 0, "notes": "DesignOption API not available"}
    opts = _collect_class(the_doc, DesignOption, not_types=False)
    if not opts:
        return {"deleted": 0, "failed": 0, "notes": "no design options"}

    secondary_ids = []
    set_ids = []
    seen_sets = set()
    for opt in opts:
        try:
            if BuiltInParameter is not None:
                p = opt.get_Parameter(BuiltInParameter.OPTION_SET_ID)
                if p is not None:
                    sid = _eid(p.AsElementId())
                    if sid is not None and sid not in seen_sets:
                        seen_sets.add(sid)
                        set_ids.append(p.AsElementId())
        except Exception:
            pass
        try:
            if getattr(opt, "IsPrimary", False):
                continue
        except Exception:
            pass
        secondary_ids.append(opt.Id)

    d1, f1 = _delete_ids(the_doc, secondary_ids)

    cat_set_ids = []
    try:
        col = FilteredElementCollector(the_doc).OfCategory(BuiltInCategory.OST_DesignOptionSets)
        cat_set_ids = [el.Id for el in col.ToElements()]
    except Exception:
        cat_set_ids = []
    if cat_set_ids:
        set_ids = cat_set_ids

    d2, f2 = _delete_ids(the_doc, set_ids)

    leftover = _collect_class(the_doc, DesignOption, not_types=False)
    d3, f3 = _delete_ids(the_doc, [o.Id for o in leftover])

    return {
        "deleted": d1 + d2 + d3,
        "failed": f1 + f2 + f3,
        "notes": "secondary {}, option sets {}, leftover options {}".format(d1, d2, d3),
    }


def delete_elements_created_in_phase(the_doc, phase_id, phase_name, seq):
    """Delete CategoryType.Model instances whose CreatedPhaseId is this phase."""
    target = None
    try:
        target = _id_value(phase_id) if hasattr(phase_id, "Value") or hasattr(phase_id, "IntegerValue") else int(phase_id)
    except Exception:
        target = None
    ids = []
    if target is not None:
        for el in _iter_model_instances(the_doc):
            if _phase_key(el, "CreatedPhaseId") != target:
                continue
            try:
                _unpin(el)
                ids.append(el.Id)
            except Exception:
                continue
    d, f = _delete_ids(the_doc, ids)
    return {
        "deleted": d,
        "failed": f,
        "notes": "created in phase '{}' (seq {}, {} candidates)".format(
            _safe_str(phase_name), seq, len(ids)
        ),
    }


def _phase_action(phase_id, phase_name, seq):
    def _fn(d, _id=phase_id, _name=phase_name, _seq=seq):
        return delete_elements_created_in_phase(d, _id, _name, _seq)

    return _fn


def delete_all_but_one_level(the_doc):
    levels = _collect_class(the_doc, Level)
    if len(levels) <= 1:
        return {"deleted": 0, "failed": 0, "notes": "already {} level(s)".format(len(levels))}

    def _elev(lv):
        try:
            return float(lv.Elevation)
        except Exception:
            return 0.0

    levels_sorted = sorted(levels, key=_elev)
    keep = levels_sorted[0]
    ids = [lv.Id for lv in levels_sorted[1:]]
    d, f = _delete_ids(the_doc, ids)
    return {
        "deleted": d,
        "failed": f,
        "notes": "kept level '{}'".format(_safe_str(keep.Name)),
    }


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------


def _write_text(path, content):
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    data = content
    if not isinstance(data, bytes):
        try:
            data = content.encode("utf-8")
        except Exception:
            data = str(content)
    with open(path, "wb") as f:
        f.write(data)


def write_csv(path, steps, original_bytes):
    lines = [
        "step,id,title,file,bytes,mb,delta_bytes,delta_mb,pct_of_original,deleted,failed,notes"
    ]
    for s in steps:
        def _cell(v):
            text = "" if v is None else _safe_str(v)
            if "," in text or '"' in text:
                text = '"' + text.replace('"', '""') + '"'
            return text

        lines.append(
            ",".join(
                [
                    _cell(s.get("index")),
                    _cell(s.get("id")),
                    _cell(s.get("title")),
                    _cell(s.get("file")),
                    _cell(s.get("bytes")),
                    _cell(s.get("mb")),
                    _cell(s.get("delta_bytes")),
                    _cell(s.get("delta_mb")),
                    _cell(s.get("pct_of_original")),
                    _cell(s.get("deleted")),
                    _cell(s.get("failed")),
                    _cell(s.get("notes")),
                ]
            )
        )
    _write_text(path, "\n".join(lines) + "\n")


def write_markdown(path, model_path, original_bytes, steps, families):
    lines = []
    lines.append("# Revit file size analysis")
    lines.append("")
    lines.append("Source: [Autodesk article]({})".format(SOURCE_ARTICLE))
    lines.append("")
    lines.append("- Model: `{}`".format(_safe_str(model_path)))
    lines.append("- Original size: **{} MB** ({} bytes)".format(_mb(original_bytes), original_bytes))
    lines.append("")
    lines.append("## Size by step")
    lines.append("")
    lines.append("| Step | MB | Delta MB | % of original | Deleted | Notes |")
    lines.append("| --- | ---: | ---: | ---: | ---: | --- |")
    for s in steps:
        lines.append(
            "| {} {} | {} | {} | {} | {} | {} |".format(
                s.get("index"),
                s.get("title"),
                s.get("mb"),
                s.get("delta_mb"),
                s.get("pct_of_original"),
                s.get("deleted"),
                _safe_str(s.get("notes")).replace("|", "/"),
            )
        )
    phase_steps = [s for s in steps if _is_phase_step(s.get("id"))]
    if phase_steps:
        lines.append("")
        lines.append("## Size by created phase")
        lines.append("")
        lines.append("Compact Save As after deleting model instances with that CreatedPhaseId. Design options cleared first.")
        lines.append("")
        lines.append("| Phase | MB | Delta MB | % of original | Deleted | Notes |")
        lines.append("| --- | ---: | ---: | ---: | ---: | --- |")
        for s in phase_steps:
            lines.append(
                "| {} | {} | {} | {} | {} | {} |".format(
                    _safe_str(s.get("title")).replace("|", "/"),
                    s.get("mb"),
                    s.get("delta_mb"),
                    s.get("pct_of_original"),
                    s.get("deleted"),
                    _safe_str(s.get("notes")).replace("|", "/"),
                )
            )
    purified = [f for f in (families or []) if f.get("purify_status")]
    if purified:
        purified.sort(
            key=lambda r: (
                -int(r.get("purify_saved_bytes") or 0),
                _safe_str(r.get("name")).lower(),
            )
        )
        lines.append("")
        lines.append("## Family Purify results")
        lines.append("")
        if any(f.get("purify_aggressive") for f in purified):
            lines.append(
                "Aggressive detached-only profile: conservative resources plus non-core views, "
                "project-unused family types, repeated Purge Unused, and nested-family recursion. "
                "Only families above the measured reload threshold are reloaded."
            )
        else:
            lines.append(
                "Analysis profile: CAD/images plus heuristically unused subcategories, patterns, "
                "appearance assets, and materials. No parameters, types, views, formulas, or nested recursion."
            )
        lines.append("")
        lines.append(
            "| Family | Status | Reloaded | Before MB | After MB | Saved MB | "
            "Removed | Nested | Failed |"
        )
        lines.append("| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |")
        for fam in purified:
            lines.append(
                "| {} | {} | {} | {} | {} | {} | {} | {} | {} |".format(
                    _safe_str(fam.get("name")).replace("|", "/"),
                    fam.get("purify_status"),
                    fam.get("purify_reloaded"),
                    _mb(fam.get("purify_before_bytes")),
                    _mb(fam.get("purify_after_bytes")),
                    fam.get("purify_saved_mb"),
                    fam.get("purify_removed"),
                    fam.get("purify_nested"),
                    fam.get("purify_failed"),
                )
            )
    ranked = [s for s in steps if s.get("delta_bytes")]
    ranked.sort(key=lambda r: r.get("delta_bytes") or 0)
    if ranked:
        lines.append("")
        lines.append("## Largest size drops")
        lines.append("")
        for s in ranked[:8]:
            drop = -(s.get("delta_bytes") or 0)
            if drop <= 0:
                continue
            lines.append(
                "- **{}**: {} MB ({}% of original)".format(
                    s.get("title"),
                    _mb(drop),
                    _pct(drop, original_bytes),
                )
            )
    if families:
        lines.append("")
        lines.append("## Loaded families (top 25)")
        lines.append("")
        lines.append("| Family | Category | In-place | Instances | Types | RFA MB |")
        lines.append("| --- | --- | --- | ---: | ---: | ---: |")
        for fam in families[:25]:
            lines.append(
                "| {} | {} | {} | {} | {} | {} |".format(
                    _safe_str(fam.get("name")).replace("|", "/"),
                    _safe_str(fam.get("category")).replace("|", "/"),
                    fam.get("is_in_place"),
                    fam.get("instance_count"),
                    fam.get("type_count"),
                    fam.get("rfa_mb") if fam.get("rfa_mb") is not None else "",
                )
            )
    dnu = [f for f in (families or []) if f.get("do_not_use")]
    if dnu:
        placed = [f for f in dnu if int(f.get("instance_count") or 0) > 0]
        unused = [f for f in dnu if int(f.get("instance_count") or 0) == 0]
        dnu_bytes = sum((f.get("rfa_bytes") or 0) for f in dnu)
        placed_bytes = sum((f.get("rfa_bytes") or 0) for f in placed)
        lines.append("")
        lines.append("## Do-not-use families (`ANVÄND EJ` / `X_ANVÄND EJ`)")
        lines.append("")
        lines.append(
            "- {} flagged families ({} MB of exported .rfa, nested content can overlap)".format(
                len(dnu), _mb(dnu_bytes)
            )
        )
        lines.append(
            "- {} still placed ({} instances, {} MB .rfa) — Purge Unused cannot remove these".format(
                len(placed),
                sum(int(f.get("instance_count") or 0) for f in placed),
                _mb(placed_bytes),
            )
        )
        lines.append(
            "- {} with zero instances (purgeable / already unused)".format(len(unused))
        )
        lines.append("")
        lines.append("| Family | Category | Instances | Types | RFA MB |")
        lines.append("| --- | --- | ---: | ---: | ---: |")
        for fam in placed[:40]:
            lines.append(
                "| {} | {} | {} | {} | {} |".format(
                    _safe_str(fam.get("name")).replace("|", "/"),
                    _safe_str(fam.get("category")).replace("|", "/"),
                    fam.get("instance_count"),
                    fam.get("type_count"),
                    fam.get("rfa_mb") if fam.get("rfa_mb") is not None else "",
                )
            )
        unused_dnu = [f for f in unused if f.get("do_not_use")] if unused else []
        if unused_dnu:
            lines.append("")
            lines.append("Unused do-not-use families (zero instances, purgeable):")
            lines.append("")
            lines.append("| Family | Category | Types | RFA MB |")
            lines.append("| --- | --- | ---: | ---: |")
            for fam in unused_dnu[:40]:
                lines.append(
                    "| {} | {} | {} | {} |".format(
                        _safe_str(fam.get("name")).replace("|", "/"),
                        _safe_str(fam.get("category")).replace("|", "/"),
                        fam.get("type_count"),
                        fam.get("rfa_mb") if fam.get("rfa_mb") is not None else "",
                    )
                )
    lines.append("")
    _write_text(path, "\n".join(lines))


def _ascii_json_fallback(obj):
    """Last-resort JSON sanitizer for IronPython strings with unknown code pages."""
    if isinstance(obj, dict):
        return {
            _ascii_json_fallback(k): _ascii_json_fallback(v)
            for k, v in obj.items()
        }
    if isinstance(obj, (list, tuple)):
        return [_ascii_json_fallback(x) for x in obj]
    if isinstance(obj, (int, float, bool)) or obj is None:
        return obj
    try:
        text = unicode(obj)  # noqa: F821
    except Exception:
        try:
            text = repr(obj)
        except Exception:
            text = "<unserializable>"
    try:
        return text.encode("ascii", "replace")
    except Exception:
        return "".join(ch if ord(ch) < 128 else "?" for ch in text)


def write_json_report(path, payload):
    try:
        text = json.dumps(_json_safe(payload), indent=2, ensure_ascii=True)
    except Exception:
        text = json.dumps(_ascii_json_fallback(payload), indent=2, ensure_ascii=True)
    _write_text(path, text)


def _recommendations(original_bytes, steps, inventory, families):
    recs = []
    dnu_placed = [
        f
        for f in (families or [])
        if f.get("do_not_use") and int(f.get("instance_count") or 0) > 0
    ]
    if dnu_placed:
        recs.append(
            "Do-not-use families (ANVAND EJ) still have {} instances across {} families. "
            "Purge Unused cannot remove them until those instances are replaced or deleted.".format(
                sum(int(f.get("instance_count") or 0) for f in dnu_placed),
                len(dnu_placed),
            )
        )
    by_id = {}
    for s in steps or []:
        by_id[s.get("id")] = s
        drop = -(s.get("delta_bytes") or 0)
        if drop <= 0:
            continue
        sid = s.get("id")
        title = s.get("title")
        mb = _mb(drop)
        pct = _pct(drop, original_bytes)
        if sid == "images_pointclouds" and drop > 1024 * 1024:
            recs.append(
                "Images/PDFs dropped {} MB ({}% of original). Compress, unlink, or remove raster content.".format(
                    mb, pct
                )
            )
        elif sid == "views" and drop > 1024 * 1024:
            recs.append(
                "Deleting unused views dropped {} MB ({}%). Cull views not on sheets in production.".format(
                    mb, pct
                )
            )
        elif sid == "purge_unused" and drop > 1024 * 1024:
            recs.append(
                "Purge unused dropped {} MB ({}%). Repeat Purge Unused until empty on a detached copy, then compact Save As.".format(
                    mb, pct
                )
            )
        elif sid == "loaded_families" and drop > 1024 * 1024:
            recs.append(
                "Unused loaded families dropped {} MB ({}%). Delete unused families from the project browser, then purge.".format(
                    mb, pct
                )
            )
        elif sid == "family_purify" and drop > 1024 * 256:
            recs.append(
                "Purifying placed loadable families shed {} MB ({}%). This detached-copy estimate "
                "removes heuristically unused family resources without deleting parameters, types, or views.".format(mb, pct)
            )
        elif sid == "design_options" and drop > 1024 * 1024:
            recs.append(
                "Clearing design options dropped {} MB ({}%). Accept primary / delete unused options before a phase split.".format(
                    mb, pct
                )
            )
        elif _is_phase_step(sid) and drop > 1024 * 1024:
            recs.append(
                "{} accounted for {} MB ({}% of original). Bytes follow CreatedPhaseId, not phase visibility.".format(
                    title, mb, pct
                )
            )
        elif sid == "geometry" and drop > 1024 * 1024:
            recs.append(
                "Leftover unphased model accounted for {} MB ({}%) after the per-phase peel.".format(
                    mb, pct
                )
            )
        elif sid == "levels" and drop > 5 * 1024 * 1024:
            recs.append(
                "Deleting extra levels dropped {} MB ({}%). If geometry was still present, leftover 3D was tied to those levels.".format(
                    mb, pct
                )
            )
        elif sid == "unplaced_groups" and drop > 1024 * 512:
            recs.append(
                "Unplaced group types dropped {} MB. Delete unused group types from the project browser.".format(mb)
            )
        elif sid == "project_parameters" and drop > 1024 * 256:
            recs.append(
                "Unused project/shared parameters dropped {} MB. Purge Unused does not remove these; delete them from Project Parameters.".format(
                    mb
                )
            )
        elif sid == "templates_filters" and drop > 1024 * 256:
            recs.append(
                "Unused view templates/filters dropped {} MB.".format(mb)
            )
        elif (s.get("failed") or 0) > 0 or (_safe_str(s.get("notes")).startswith("transaction failed")):
            recs.append("{} failed ({}). Size for later steps may be mixed.".format(title, s.get("notes")))
    inv = inventory or {}
    phase_steps = [s for s in (steps or []) if _is_phase_step(s.get("id"))]
    if phase_steps:
        recs.append(
            "Per-phase MB is compact Save As after deleting model instances created in that phase. "
            "Demolished is a flag on the same element, not extra file bytes. Shared family types stay until leftover geometry."
        )
    if int(inv.get("views_not_on_sheets") or 0) > 50:
        recs.append(
            "{} views are not on sheets. That is usually cheap to cull and showed up in the views peel.".format(
                inv.get("views_not_on_sheets")
            )
        )
    if int(inv.get("image_types") or 0) > 0:
        recs.append(
            "Model has {} image types ({} imports, {} links). Raster content often survives purge.".format(
                inv.get("image_types"),
                inv.get("image_imports"),
                inv.get("image_links"),
            )
        )
    unused_loadable = 0
    for f in families or []:
        if int(f.get("instance_count") or 0) == 0 and not f.get("is_in_place"):
            unused_loadable += 1
    if unused_loadable:
        recs.append("{} unused loadable families were present before the peel.".format(unused_loadable))
    recs.append(
        "Compact Save As to a new path is what actually shrinks Revit 2025 files after a purge."
    )
    return recs


def write_html_report(path, model_path, original_bytes, steps, families, inventory, extra_exports=None):
    inv = inventory or {}
    last = steps[-1] if steps else {}
    last_mb = last.get("mb")
    ranked = [s for s in (steps or []) if s.get("delta_bytes")]
    ranked.sort(key=lambda r: r.get("delta_bytes") or 0)
    largest = ranked[0] if ranked else None
    largest_drop = -(largest.get("delta_bytes") or 0) if largest else 0
    dnu = [f for f in (families or []) if f.get("do_not_use")]
    dnu_placed = [f for f in dnu if int(f.get("instance_count") or 0) > 0]
    unused_fam = [
        f
        for f in (families or [])
        if int(f.get("instance_count") or 0) == 0 and not f.get("is_in_place")
    ]
    recs = _recommendations(original_bytes, steps, inv, families)
    by_size = sorted(
        [f for f in (families or []) if f.get("rfa_bytes")],
        key=lambda r: -(r.get("rfa_bytes") or 0),
    )
    purified = sorted(
        [f for f in (families or []) if f.get("purify_status")],
        key=lambda r: (
            -int(r.get("purify_saved_bytes") or 0),
            _safe_str(r.get("name")).lower(),
        ),
    )
    purified_ok = [
        f
        for f in purified
        if f.get("purify_status") in ("ok", "reloaded", "measured")
    ]
    purified_saved = sum(int(f.get("purify_saved_bytes") or 0) for f in purified_ok)
    aggressive_purify = any(f.get("purify_aggressive") for f in purified)
    cats = _family_category_rollup(families)

    max_mb = 1.0
    for s in steps or []:
        mb = s.get("mb")
        if mb and float(mb) > max_mb:
            max_mb = float(mb)
    if original_bytes:
        omb = float(_mb(original_bytes) or 0)
        if omb > max_mb:
            max_mb = omb

    bars = []
    svg_h = 28 * (len(steps or []) + 1)
    y = 8
    for s in steps or []:
        mb = float(s.get("mb") or 0)
        w = int(720.0 * mb / max_mb) if max_mb else 0
        failed = bool(s.get("failed")) or _safe_str(s.get("notes")).startswith("transaction failed")
        fill = "#b42318" if failed else "#175cd3"
        label = "{:02d} {}".format(s.get("index"), _html_escape(s.get("title")))
        delta = s.get("delta_mb")
        delta_s = "" if delta is None else "{:+.3f} MB".format(float(delta))
        bars.append(
            '<g transform="translate(0,{})">'
            '<text x="0" y="12" class="lbl">{}</text>'
            '<rect x="240" y="0" width="{}" height="16" fill="{}"></rect>'
            '<text x="{}" y="12" class="val">{} MB {}</text>'
            "</g>".format(y, label, w, fill, 248 + w, s.get("mb"), delta_s)
        )
        y += 26

    health_rows = [
        ("Warnings", inv.get("warnings_total")),
        ("Views", inv.get("views_total")),
        ("Views not on sheets", inv.get("views_not_on_sheets")),
        ("View templates", inv.get("view_templates")),
        ("Unused view templates", inv.get("unused_view_templates")),
        ("View filters", inv.get("view_filters")),
        ("Unused view filters", inv.get("unused_view_filters")),
        ("Sheets", inv.get("sheets")),
        ("Levels", inv.get("levels")),
        ("Revit links", len(inv.get("revit_links") or [])),
        ("CAD imports", inv.get("cad_imports")),
        ("CAD view-links", inv.get("cad_links")),
        ("CAD model-links", inv.get("cad_model_links")),
        ("Image types", inv.get("image_types")),
        ("Image imports", inv.get("image_imports")),
        ("Image links", inv.get("image_links")),
        ("Point clouds", inv.get("point_clouds")),
        ("Loaded families", inv.get("loaded_families")),
        ("In-place families", inv.get("in_place_families")),
        ("Unplaced model groups", inv.get("unplaced_model_groups")),
        ("Unplaced detail groups", inv.get("unplaced_detail_groups")),
        ("Design options", inv.get("design_options")),
        ("Phases", len(inv.get("phases") or [])),
        ("Unphased model", inv.get("unphased_model")),
        ("Filled regions", inv.get("filled_regions")),
        ("Unused text types", inv.get("unused_text_types")),
        ("Unused dimension types", inv.get("unused_dimension_types")),
        ("Worksets", inv.get("worksets")),
        ("Rooms placed", inv.get("rooms_placed")),
        ("Rooms unplaced", inv.get("rooms_unplaced")),
        ("Rooms not enclosed", inv.get("rooms_not_enclosed")),
        ("Parameter elements", inv.get("parameter_elements")),
        ("Shared parameter elements", inv.get("shared_parameter_elements")),
    ]

    fam_rows = []
    for fam in by_size[:40]:
        dnu_cls = ' class="dnu"' if fam.get("do_not_use") else ""
        fam_rows.append(
            "<tr{}><td>{}</td><td>{}</td><td class='n'>{}</td><td class='n'>{}</td><td>{}</td></tr>".format(
                dnu_cls,
                _html_escape(fam.get("name")),
                _html_escape(fam.get("category")),
                fam.get("instance_count"),
                fam.get("rfa_mb"),
                "yes" if fam.get("do_not_use") else "",
            )
        )

    purify_rows = []
    for fam in purified:
        cls = ' class="fail"' if fam.get("purify_status") == "failed" else ""
        purify_rows.append(
            "<tr{}><td>{}</td><td>{}</td><td>{}</td><td class='n'>{}</td><td class='n'>{}</td>"
            "<td class='n'>{}</td><td class='n'>{}</td><td class='n'>{}</td>"
            "<td class='n'>{}</td><td>{}</td></tr>".format(
                cls,
                _html_escape(fam.get("name")),
                _html_escape(fam.get("purify_status")),
                "yes" if fam.get("purify_reloaded") else "no",
                _mb(fam.get("purify_before_bytes")),
                _mb(fam.get("purify_after_bytes")),
                fam.get("purify_saved_mb"),
                fam.get("purify_removed"),
                fam.get("purify_nested"),
                fam.get("purify_failed"),
                _html_escape(fam.get("purify_error")),
            )
        )

    cat_rows = []
    for cat in cats[:20]:
        cat_rows.append(
            "<tr><td>{}</td><td class='n'>{}</td><td class='n'>{}</td><td class='n'>{}</td><td class='n'>{}</td></tr>".format(
                _html_escape(cat.get("category")),
                cat.get("families"),
                cat.get("instances"),
                _mb(cat.get("rfa_bytes")),
                cat.get("do_not_use"),
            )
        )

    rec_items = "".join("<li>{}</li>".format(_html_escape(r)) for r in recs)
    health_html = "".join(
        "<div class='cell'><div class='k'>{}</div><div class='v'>{}</div></div>".format(
            _html_escape(k), _html_escape(v)
        )
        for k, v in health_rows
    )
    step_rows = []
    for s in steps or []:
        cls = ' class="fail"' if (
            bool(s.get("failed")) or _safe_str(s.get("notes")).startswith("transaction failed")
        ) else ""
        step_rows.append(
            "<tr{}><td class='n'>{}</td><td>{}</td><td class='n'>{}</td><td class='n'>{}</td><td class='n'>{}</td><td class='n'>{}</td><td>{}</td></tr>".format(
                cls,
                s.get("index"),
                _html_escape(s.get("title")),
                s.get("mb"),
                s.get("delta_mb"),
                s.get("pct_of_original"),
                s.get("deleted"),
                _html_escape(s.get("notes")),
            )
        )

    phase_inv_rows = []
    for p in inv.get("phases") or []:
        phase_inv_rows.append(
            "<tr><td class='n'>{}</td><td>{}</td><td class='n'>{}</td><td class='n'>{}</td></tr>".format(
                p.get("sequence"),
                _html_escape(p.get("name")),
                p.get("created"),
                p.get("demolished"),
            )
        )
    phase_steps = [s for s in (steps or []) if _is_phase_step(s.get("id"))]
    phase_drop_total = 0
    for s in phase_steps:
        drop = -(s.get("delta_bytes") or 0)
        if drop > 0:
            phase_drop_total += drop
    phase_size_rows = []
    for s in phase_steps:
        drop = -(s.get("delta_bytes") or 0)
        share = _pct(drop, phase_drop_total) if (drop > 0 and phase_drop_total) else 0
        pcls = ' class="fail"' if (
            bool(s.get("failed")) or _safe_str(s.get("notes")).startswith("transaction failed")
        ) else ""
        phase_size_rows.append(
            "<tr{}><td>{}</td><td class='n'>{}</td><td class='n'>{}</td><td class='n'>{}</td><td class='n'>{}</td><td class='n'>{}</td><td>{}</td></tr>".format(
                pcls,
                _html_escape(s.get("title")),
                s.get("mb"),
                s.get("delta_mb"),
                s.get("pct_of_original"),
                share,
                s.get("deleted"),
                _html_escape(s.get("notes")),
            )
        )

    html = []
    html.append("<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'>")
    html.append("<title>Revit file size analysis</title>")
    html.append(
        "<style>"
        "body{font:14px/1.45 Segoe UI,system-ui,sans-serif;margin:24px;color:#101828;background:#fff;}"
        "h1,h2{font-weight:650;margin:1.4em 0 .5em;}"
        "h1{font-size:22px;margin-top:0;}"
        "a{color:#175cd3;}"
        ".muted{color:#475467;}"
        ".cards{display:flex;flex-wrap:wrap;gap:12px;margin:16px 0 24px;}"
        ".card{border:1px solid #eaecf0;padding:12px 14px;min-width:140px;}"
        ".card .k{color:#475467;font-size:12px;}"
        ".card .v{font-size:20px;font-weight:650;}"
        ".grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px;}"
        ".cell{border:1px solid #eaecf0;padding:8px 10px;}"
        ".cell .k{color:#475467;font-size:12px;}"
        ".cell .v{font-weight:600;}"
        "table{border-collapse:collapse;width:100%;margin:8px 0 24px;font-size:13px;}"
        "th,td{border-bottom:1px solid #eaecf0;padding:6px 8px;text-align:left;}"
        "th{cursor:pointer;background:#f9fafb;}"
        "td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;}"
        "tr.fail td{color:#b42318;}"
        "tr.dnu td{background:#fff6ed;}"
        "svg .lbl,svg .val{font:12px Segoe UI,sans-serif;fill:#344054;}"
        "ul{padding-left:18px;}"
        "</style></head><body>"
    )
    html.append("<h1>Revit file size analysis</h1>")
    html.append(
        "<p class='muted'>Model: {}<br>Session: {}<br>Source: <a href='{}'>Autodesk article</a></p>".format(
            _html_escape(model_path),
            _html_escape(sessionId),
            SOURCE_ARTICLE,
        )
    )
    html.append("<div class='cards'>")
    html.append(
        "<div class='card'><div class='k'>Original</div><div class='v'>{} MB</div></div>".format(
            _mb(original_bytes)
        )
    )
    html.append(
        "<div class='card'><div class='k'>After last step</div><div class='v'>{} MB</div></div>".format(
            last_mb
        )
    )
    html.append(
        "<div class='card'><div class='k'>Largest drop</div><div class='v'>{} MB</div><div class='k'>{}</div></div>".format(
            _mb(largest_drop),
            _html_escape(largest.get("title") if largest else ""),
        )
    )
    html.append(
        "<div class='card'><div class='k'>Warnings</div><div class='v'>{}</div></div>".format(
            inv.get("warnings_total")
        )
    )
    html.append(
        "<div class='card'><div class='k'>Unused families</div><div class='v'>{}</div></div>".format(
            len(unused_fam)
        )
    )
    html.append(
        "<div class='card'><div class='k'>Images</div><div class='v'>{}</div></div>".format(
            inv.get("image_types")
        )
    )
    html.append(
        "<div class='card'><div class='k'>Views not on sheets</div><div class='v'>{}</div></div>".format(
            inv.get("views_not_on_sheets")
        )
    )
    html.append(
        "<div class='card'><div class='k'>Do-not-use placed</div><div class='v'>{}</div></div>".format(
            len(dnu_placed)
        )
    )
    html.append(
        "<div class='card'><div class='k'>Phases</div><div class='v'>{}</div></div>".format(
            len(inv.get("phases") or [])
        )
    )
    if purified:
        html.append(
            "<div class='card'><div class='k'>Families purified</div><div class='v'>{}</div>"
            "<div class='k'>{} MB compact RFA reduction</div></div>".format(
                len(purified_ok), _mb(purified_saved)
            )
        )
    html.append("</div>")

    html.append("<h2>Size by step</h2>")
    html.append(
        '<svg viewBox="0 0 980 {}" width="100%" role="img" aria-label="Size by step">{}</svg>'.format(
            max(svg_h, 40), "".join(bars)
        )
    )
    html.append(
        "<table data-sort><thead><tr>"
        "<th class='n'>#</th><th>Step</th><th class='n'>MB</th><th class='n'>Delta MB</th>"
        "<th class='n'>% orig</th><th class='n'>Deleted</th><th>Notes</th>"
        "</tr></thead><tbody>{}</tbody></table>".format("".join(step_rows))
    )

    if phase_size_rows:
        html.append("<h2>Size by created phase</h2>")
        html.append(
            "<p class='muted'>Each row is compact Save As after deleting model instances with that "
            "CreatedPhaseId. Design options are cleared first. Demolished is a flag, not extra bytes. "
            "Share is of the sum of phase drops ({} MB).</p>".format(_mb(phase_drop_total))
        )
        html.append(
            "<table data-sort><thead><tr>"
            "<th>Phase</th><th class='n'>After MB</th><th class='n'>Delta MB</th>"
            "<th class='n'>% orig</th><th class='n'>% of phases</th><th class='n'>Deleted</th><th>Notes</th>"
            "</tr></thead><tbody>{}</tbody></table>".format("".join(phase_size_rows))
        )

    html.append("<h2>Health inventory (before peel)</h2>")
    html.append("<div class='grid'>{}</div>".format(health_html))

    if phase_inv_rows:
        html.append("<h2>Phases (before peel)</h2>")
        html.append(
            "<p class='muted'>Created = model instances with this CreatedPhaseId. "
            "Demolished in phase = same instances whose DemolishedPhaseId is this phase "
            "(not extra file size). Unphased model: {}.</p>".format(
                _html_escape(inv.get("unphased_model"))
            )
        )
        html.append(
            "<table data-sort><thead><tr><th class='n'>Seq</th><th>Phase</th>"
            "<th class='n'>Created</th><th class='n'>Demolished in phase</th>"
            "</tr></thead><tbody>{}</tbody></table>".format("".join(phase_inv_rows))
        )

    if purify_rows:
        html.append("<h2>Family Purify results</h2>")
        if aggressive_purify:
            html.append(
                "<p class='muted'>Aggressive detached-only benchmark: conservative resource cleanup "
                "plus non-core views, project-unused family types, repeated Purge Unused, and nested-family "
                "recursion to depth 2. Only families above the measured reload threshold are reloaded. "
                "Do not use these snapshots as production models without review.</p>"
            )
        else:
            html.append(
                "<p class='muted'>Justin Biju's resource-cleaning profile adapted for detached-copy analysis: "
                "remove CAD/images and unused subcategories, line/fill patterns, appearance assets, and "
                "materials. Parameters, family types, views, formulas, and nested recursion are untouched.</p>"
            )
        html.append(
            "<table data-sort><thead><tr><th>Family</th><th>Status</th><th>Reloaded</th>"
            "<th class='n'>Before MB</th><th class='n'>After MB</th>"
            "<th class='n'>Saved MB</th><th class='n'>Deleted incl. dependents</th>"
            "<th class='n'>Nested</th><th class='n'>Failed deletes</th><th>Error</th>"
            "</tr></thead><tbody>{}</tbody></table>".format("".join(purify_rows))
        )

    html.append("<h2>Loaded families (top 40 by .rfa MB)</h2>")
    html.append(
        "<p class='muted'>Orange rows are do-not-use names (ANVAND EJ). Nested family content can overlap in MB totals.</p>"
    )
    html.append(
        "<table data-sort><thead><tr><th>Family</th><th>Category</th><th class='n'>Instances</th>"
        "<th class='n'>RFA MB</th><th>Do-not-use</th></tr></thead><tbody>{}</tbody></table>".format(
            "".join(fam_rows)
        )
    )
    if cat_rows:
        html.append("<h2>Families by category</h2>")
        html.append(
            "<table data-sort><thead><tr><th>Category</th><th class='n'>Families</th><th class='n'>Instances</th>"
            "<th class='n'>RFA MB</th><th class='n'>Do-not-use</th></tr></thead><tbody>{}</tbody></table>".format(
                "".join(cat_rows)
            )
        )
    html.append("<h2>Recommendations</h2><ul>{}</ul>".format(rec_items))
    html.append(
        "<script>"
        "(function(){function cmp(a,b,n,dir){var x=a.cells[n].innerText,y=b.cells[n].innerText;"
        "var nx=parseFloat(x.replace(/[^0-9.+-]/g,'')),ny=parseFloat(y.replace(/[^0-9.+-]/g,''));"
        "if(!isNaN(nx)&&!isNaN(ny))return (nx-ny)*dir;return x.localeCompare(y)*dir;}"
        "document.querySelectorAll('table[data-sort] th').forEach(function(th,i){"
        "th.addEventListener('click',function(){var tb=th.closest('table').tBodies[0];"
        "var dir=th.dataset.dir==='1'?-1:1;th.dataset.dir=dir===1?'1':'0';"
        "var rows=[].slice.call(tb.rows);rows.sort(function(a,b){return cmp(a,b,i,dir);});"
        "rows.forEach(function(r){tb.appendChild(r);});});});})();"
        "</script></body></html>"
    )
    _write_text(path, "\n".join(html))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def _default_out_dir(model_path):
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(model_path or "model"))[0]
    parent = os.path.dirname(model_path) if model_path else os.environ.get("TEMP", r"C:\Temp")
    return os.path.join(parent, "_filesize_analysis", "{}_{}".format(_safe_filename(base), stamp))


def run_family_analytics(the_doc, model_path, out_dir=None):
    """Read-only family inventory + export-folder join. Does not Save As or purge."""
    out_dir = out_dir or os.environ.get("RBP_SIZE_OUT_DIR", "").strip() or _default_out_dir(model_path)
    family_export_dir = os.environ.get("RBP_FAMILY_EXPORT_DIR", "").strip()
    measure_families = _env_flag("RBP_MEASURE_FAMILY_SIZES", True)

    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)

    original_bytes = _file_bytes(model_path)
    log("=== Family analytics (read-only) ===")
    log("Model: {}".format(model_path))
    log("Original size: {} MB".format(_mb(original_bytes)))
    log("Output: {}".format(out_dir))
    if family_export_dir:
        log("Export folder: {}".format(family_export_dir))
    else:
        log("Export folder: (not set — set RBP_FAMILY_EXPORT_DIR to join .rfa sizes)")

    log("Collecting inventory...")
    inventory = collect_inventory(the_doc, model_path)
    fam_tmp = os.path.join(out_dir, "_family_tmp")
    if measure_families:
        log("Measuring loaded family .rfa sizes (slow)...")
    families = collect_family_inventory(the_doc, measure_families, fam_tmp)
    families = annotate_families_with_exports(families, family_export_dir)
    extra = unmatched_exported_rfas(families, family_export_dir)

    placed = [f for f in families if int(f.get("instance_count") or 0) > 0]
    unused = [f for f in families if int(f.get("instance_count") or 0) == 0 and not f.get("is_in_place")]
    dnu = [f for f in families if f.get("do_not_use")]
    dnu_placed = [f for f in dnu if int(f.get("instance_count") or 0) > 0]
    log(
        "Families: {} loaded, {} with instances, {} unused, {} in-place".format(
            len(families),
            len(placed),
            len(unused),
            inventory.get("in_place_families"),
        )
    )
    log(
        "Do-not-use names: {} ({} still placed, {} zero instances)".format(
            len(dnu), len(dnu_placed), len(dnu) - len(dnu_placed)
        )
    )
    log("Export matches: {}, unmatched .rfa: {}".format(
        len([f for f in families if f.get("in_export_folder")]),
        len(extra),
    ))

    stem = _safe_filename(os.path.splitext(os.path.basename(model_path or "model"))[0])
    fam_csv = os.path.join(out_dir, stem + "_families.csv")
    fam_md = os.path.join(out_dir, stem + "_families.md")
    json_path = os.path.join(out_dir, stem + "_families.json")
    html_path = os.path.join(out_dir, "report.html")
    html_stem = os.path.join(out_dir, stem + "_report.html")
    extra_exports = extra

    def _w_csv():
        write_family_csv(fam_csv, families)
        return fam_csv

    def _w_md():
        write_family_markdown(fam_md, model_path, inventory, families, extra_exports, family_export_dir)
        return fam_md

    def _w_json():
        write_json_report(json_path, payload)
        return json_path

    def _w_html():
        write_html_report(
            html_path, model_path, original_bytes, [], families, inventory, extra_exports
        )
        try:
            write_html_report(
                html_stem, model_path, original_bytes, [], families, inventory, extra_exports
            )
        except Exception:
            pass
        return html_path

    payload = {
        "model_path": model_path,
        "output_dir": out_dir,
        "export_dir": family_export_dir,
        "original_bytes": original_bytes,
        "original_mb": _mb(original_bytes),
        "inventory": inventory,
        "families": families,
        "unmatched_exports": extra,
        "do_not_use_count": len(dnu),
        "do_not_use_placed": len(dnu_placed),
        "do_not_use_unused": len(dnu) - len(dnu_placed),
        "steps": [],
    }
    _try_write("Families CSV", _w_csv)
    _try_write("Families report", _w_md)
    _try_write("JSON", _w_json)
    _try_write("HTML", _w_html)
    sidecar = os.environ.get("RBP_SIDECAR_PATH", "").strip()
    if sidecar:
        def _w_side():
            write_json_report(sidecar, payload)
            return sidecar

        _try_write("Sidecar", _w_side)
    return payload


def run_family_purify_only(the_doc, model_path, out_dir):
    """Compact baseline, purify families, compact again, and write focused reports."""
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    max_families = max(0, _env_int("RBP_PURIFY_FAMILY_MAX", 0))
    family_offset = max(0, _env_int("RBP_PURIFY_FAMILY_OFFSET", 0))
    aggressive = _env_flag("RBP_FAMILY_PURIFY_AGGRESSIVE", False)
    ranking_csv = os.environ.get("RBP_FAMILY_RANKING_CSV", "").strip()
    reload_min_kb = max(
        0,
        _env_int("RBP_PURIFY_RELOAD_MIN_KB", 5120 if aggressive else 0),
    )
    family_export_dir = os.environ.get("RBP_FAMILY_EXPORT_DIR", "").strip()
    original_bytes = _file_bytes(model_path)
    stem = _safe_filename(os.path.splitext(os.path.basename(model_path or "model"))[0])

    keep_view = create_keep_view(the_doc)
    _activate_view(the_doc, keep_view)
    inventory = collect_inventory(the_doc, model_path)
    families = collect_family_inventory(the_doc, False, os.path.join(out_dir, "_family_tmp"))
    families = annotate_families_with_exports(families, family_export_dir)
    families = annotate_families_with_ranking_csv(families, ranking_csv)

    before_path = os.path.join(out_dir, stem + "_00_before_family_purify.rvt")
    before_bytes = save_compact_snapshot(the_doc, before_path, keep_view)
    steps = [
        {
            "index": 0,
            "id": "compact",
            "title": "Compact baseline before Family Purify",
            "file": before_path,
            "bytes": before_bytes,
            "mb": _mb(before_bytes),
            "delta_bytes": (
                int(before_bytes) - int(original_bytes)
                if before_bytes is not None and original_bytes is not None
                else None
            ),
            "delta_mb": _mb(
                int(before_bytes) - int(original_bytes)
                if before_bytes is not None and original_bytes is not None
                else None
            ),
            "pct_of_original": _pct(before_bytes, original_bytes),
            "deleted": 0,
            "failed": 0,
            "notes": "normalized compact baseline",
        }
    ]

    result = purify_loaded_families(
        the_doc,
        os.path.join(out_dir, "_family_purify_tmp"),
        families,
        max_families,
        aggressive=aggressive,
        reload_min_bytes=reload_min_kb * 1024,
        family_offset=family_offset,
    )
    if aggressive:
        ok, purged = _run_transaction(
            the_doc,
            "Purge unused types from purified families",
            lambda d: _purge_project_types_for_families(
                d, result.get("reloaded_family_ids")
            ),
        )
        if ok:
            result["project_types_purged"] = int(purged[0] or 0)
            result["project_type_purge_failed"] = int(purged[1] or 0)
            result["deleted"] = int(result.get("deleted") or 0) + int(purged[0] or 0)
            result["failed"] = int(result.get("failed") or 0) + int(purged[1] or 0)
            result["notes"] += "; {} unused project family types purged".format(
                int(purged[0] or 0)
            )
        else:
            result["notes"] += "; project family-type purge failed: {}".format(purged)
    keep_view = resolve_keep_view(the_doc) or create_keep_view(the_doc)
    after_path = os.path.join(out_dir, stem + "_01_family_purify.rvt")
    after_bytes = save_compact_snapshot(the_doc, after_path, keep_view)
    delta = (
        int(after_bytes) - int(before_bytes)
        if after_bytes is not None and before_bytes is not None
        else None
    )
    steps.append(
        {
            "index": 1,
            "id": "family_purify",
            "title": "Purify loadable families",
            "file": after_path,
            "bytes": after_bytes,
            "mb": _mb(after_bytes),
            "delta_bytes": delta,
            "delta_mb": _mb(delta),
            "pct_of_original": _pct(after_bytes, original_bytes),
            "deleted": int(result.get("deleted") or 0),
            "failed": int(result.get("failed") or 0),
            "notes": _safe_str(result.get("notes")),
        }
    )

    payload = {
        "source_article": SOURCE_ARTICLE,
        "session_id": sessionId,
        "mode": "family_purify_only",
        "aggressive": aggressive,
        "ranking_csv": ranking_csv,
        "reload_min_kb": reload_min_kb,
        "model_path": model_path,
        "output_dir": out_dir,
        "original_bytes": original_bytes,
        "original_mb": _mb(original_bytes),
        "inventory": inventory,
        "families": families,
        "family_purification": result,
        "steps": steps,
        "purify_family_max": max_families,
        "purify_family_offset": family_offset,
    }

    csv_path = os.path.join(out_dir, stem + "_filesize.csv")
    md_path = os.path.join(out_dir, stem + "_filesize.md")
    json_path = os.path.join(out_dir, stem + "_filesize.json")
    fam_csv = os.path.join(out_dir, stem + "_families.csv")
    html_path = os.path.join(out_dir, "report.html")
    html_stem = os.path.join(out_dir, stem + "_report.html")

    _try_write("CSV", lambda: (write_csv(csv_path, steps, original_bytes), csv_path)[1])
    _try_write(
        "Report",
        lambda: (write_markdown(md_path, model_path, original_bytes, steps, families), md_path)[1],
    )
    _try_write("Families CSV", lambda: (write_family_csv(fam_csv, families), fam_csv)[1])
    _try_write("JSON", lambda: (write_json_report(json_path, payload), json_path)[1])
    _try_write(
        "HTML",
        lambda: (
            write_html_report(html_path, model_path, original_bytes, steps, families, inventory),
            write_html_report(html_stem, model_path, original_bytes, steps, families, inventory),
            html_path,
        )[2],
    )
    return payload


def run(the_doc, model_path):
    out_dir = os.environ.get("RBP_SIZE_OUT_DIR", "").strip() or _default_out_dir(model_path)
    keep_files = _env_flag("RBP_KEEP_STEP_FILES", True)
    purge_passes = _env_int("RBP_PURGE_PASSES", 15)
    purify_families = _env_flag("RBP_PURIFY_FAMILIES", False)
    purify_family_max = max(0, _env_int("RBP_PURIFY_FAMILY_MAX", 0))
    # Purification already measures each family before/after. Avoid a duplicate
    # EditFamily pass by default; callers can explicitly opt back in.
    measure_families = _env_flag("RBP_MEASURE_FAMILY_SIZES", not purify_families)
    family_export_dir = os.environ.get("RBP_FAMILY_EXPORT_DIR", "").strip()

    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)

    if _env_flag("RBP_FAMILY_PURIFY_ONLY", False):
        return run_family_purify_only(the_doc, model_path, out_dir)

    if _env_flag("RBP_FAMILY_ANALYTICS_ONLY", False):
        return run_family_analytics(the_doc, model_path, out_dir)

    original_bytes = _file_bytes(model_path)
    log("Model: {}".format(model_path))
    log("Original size: {} MB".format(_mb(original_bytes)))
    log("Output: {}".format(out_dir))
    log("Keep step files: {}".format(keep_files))

    keep_view = create_keep_view(the_doc)
    _activate_view(the_doc, keep_view)
    log("Keep view: {}".format(_safe_str(keep_view.Name)))

    def _keep():
        v = resolve_keep_view(the_doc)
        if v is None:
            v = create_keep_view(the_doc)
        return v

    log("Collecting inventory...")
    inventory = collect_inventory(the_doc, model_path)
    fam_tmp = os.path.join(out_dir, "_family_tmp")
    if measure_families:
        log("Measuring loaded family .rfa sizes (slow)...")
    families = collect_family_inventory(the_doc, measure_families, fam_tmp)
    if family_export_dir:
        log("Matching exported families in {}".format(family_export_dir))
    families = annotate_families_with_exports(families, family_export_dir)

    placed = [f for f in families if f.get("instance_count", 0) > 0]
    unused = [f for f in families if f.get("instance_count", 0) == 0 and not f.get("is_in_place")]
    dnu = [f for f in families if f.get("do_not_use")]
    dnu_placed = [f for f in dnu if int(f.get("instance_count") or 0) > 0]
    log(
        "Families: {} loaded, {} with instances, {} unused, {} in-place".format(
            len(families),
            len(placed),
            len(unused),
            inventory.get("in_place_families"),
        )
    )
    log(
        "Do-not-use names: {} ({} still placed, {} zero instances)".format(
            len(dnu), len(dnu_placed), len(dnu) - len(dnu_placed)
        )
    )
    log(
        "Links: {} Revit, {} CAD links, {} CAD imports, {} images, {} point clouds".format(
            len(inventory.get("revit_links") or []),
            inventory.get("cad_links"),
            inventory.get("cad_imports"),
            inventory.get("image_types"),
            inventory.get("point_clouds"),
        )
    )
    log(
        "Views: {}, sheets: {}, levels: {}, warnings: {}".format(
            inventory.get("views_total"),
            inventory.get("sheets"),
            inventory.get("levels"),
            inventory.get("warnings_total"),
        )
    )
    log(
        "Design options: {}, option sets: {}".format(
            inventory.get("design_options"),
            inventory.get("design_option_sets"),
        )
    )
    for p in inventory.get("phases") or []:
        log(
            "  Phase seq {} '{}': created {}, demolished-in {}".format(
                p.get("sequence"),
                p.get("name"),
                p.get("created"),
                p.get("demolished"),
            )
        )
    log("Unphased model instances: {}".format(inventory.get("unphased_model")))
    log(
        "Family Purify analysis: {}{}".format(
            purify_families,
            " (max {})".format(purify_family_max) if purify_family_max else "",
        )
    )
    n_phases = len(inventory.get("phases") or [])
    snapshot_n = _expected_snapshot_count(n_phases, purify_families)
    if original_bytes and keep_files:
        est = _mb(original_bytes * snapshot_n)
        log(
            "Disk warning: {} snapshots may use on the order of {} MB".format(
                snapshot_n, est
            )
        )

    steps = []
    prev_bytes = original_bytes
    prev_snapshot = None
    stem = _safe_filename(os.path.splitext(os.path.basename(model_path or "model"))[0])

    def record_step(index, step_id, title, snapshot_path, size_bytes, deleted, failed, notes):
        delta = None
        if size_bytes is not None and prev_bytes is not None:
            delta = int(size_bytes) - int(prev_bytes)
        row = {
            "index": index,
            "id": step_id,
            "title": title,
            "file": snapshot_path,
            "bytes": size_bytes,
            "mb": _mb(size_bytes),
            "delta_bytes": delta,
            "delta_mb": _mb(delta),
            "pct_of_original": _pct(size_bytes, original_bytes),
            "deleted": deleted,
            "failed": failed,
            "notes": notes,
        }
        steps.append(row)
        drop = ""
        if delta is not None:
            drop = "  delta {:+.3f} MB".format(_mb(delta))
        log("[{:02d}] {} -> {} MB{}".format(index, title, _mb(size_bytes), drop))
        return size_bytes

    def snapshot(index, step_id):
        name = "{}_{:02d}_{}.rvt".format(stem, index, step_id)
        return os.path.join(out_dir, name)

    def maybe_delete_previous(current_path):
        if keep_files or not prev_snapshot:
            return
        if prev_snapshot == current_path:
            return
        try:
            if os.path.isfile(prev_snapshot):
                os.remove(prev_snapshot)
                log("  deleted previous snapshot (RBP_KEEP_STEP_FILES=0)")
        except Exception as ex:
            log("  could not delete previous snapshot: {}".format(ex))

    # 00 compact
    path0 = snapshot(0, "compact")
    log("[00] Compact Save As (Autodesk: compact the model)...")
    try:
        size0 = save_compact_snapshot(the_doc, path0, _keep())
        prev_bytes = record_step(0, "compact", "Compact Save As", path0, size0, 0, 0, "SaveAs always compacts")
        prev_snapshot = path0
    except Exception as ex:
        log("FATAL: compact Save As failed: {}".format(ex))
        raise

    log("Purge unused: looping GetUnusedElements up to {} times until nothing remains".format(purge_passes))
    plan = [
        (1, "purge_unused", "Purge unused", lambda d: purge_unused(d, purge_passes)),
        (2, "revit_links", "Remove Revit links", remove_revit_links),
        (3, "cad", "Remove CAD links/imports", remove_cad),
        (4, "unplaced_groups", "Delete unplaced group types", remove_unplaced_groups),
        (5, "images_pointclouds", "Remove images/PDFs/point clouds", remove_images_and_pointclouds),
        (6, "sheets", "Remove all sheets", remove_sheets),
        (7, "views", "Remove views except keep 3D", lambda d: remove_views_except(d, _keep())),
        (8, "templates_filters", "Unused view templates + filters", remove_unused_templates_and_filters),
        (9, "inplace", "Remove in-place families", remove_inplace_families),
        (10, "project_parameters", "Unused project/shared parameters", remove_unused_project_parameters),
        (11, "loaded_families", "Remove unused loaded families + purge", lambda d: remove_unused_loaded_families(d, purge_passes)),
    ]
    next_i = 12
    if purify_families:
        purify_tmp = os.path.join(out_dir, "_family_purify_tmp")
        plan.append(
            (
                next_i,
                "family_purify",
                "Purify placed loadable families",
                lambda d: purify_loaded_families(
                    d, purify_tmp, families, purify_family_max
                ),
            )
        )
        next_i += 1
    plan.append((next_i, "design_options", "Clear design options", remove_design_options))
    next_i += 1
    phase_n = 0
    for p in list_phases(the_doc):
        phase_n += 1
        seq = 0
        try:
            seq = int(p.SequenceNumber)
        except Exception:
            seq = 0
        pname = _safe_str(p.Name)
        sid = "phase_{:02d}".format(phase_n)
        title = "Created in phase '{}'".format(pname)
        plan.append((next_i, sid, title, _phase_action(p.Id, pname, seq)))
        next_i += 1
    plan.append((next_i, "geometry", "Delete leftover unphased model", delete_model_elements))
    next_i += 1
    plan.append((next_i, "levels", "Delete all but one level", delete_all_but_one_level))
    next_i += 1
    # Families (and their types/materials) whose instances went with the phase
    # and geometry peel are only unused now; purge so their size is attributed.
    plan.append(
        (
            next_i,
            "final_purge",
            "Final purge (families freed by the geometry peel)",
            lambda d: remove_unused_loaded_families(d, purge_passes),
        )
    )

    for index, step_id, title, action in plan:
        log("[{:02d}] {}...".format(index, title))
        keep_view = _keep()
        if _step_touches_keep_view(step_id):
            _activate_view(the_doc, keep_view)
        if step_id == "family_purify":
            # EditFamily and LoadFamily require the project document to be unmodifiable.
            try:
                result = action(the_doc)
                ok = True
            except Exception as ex:
                result = ex
                ok = False
        else:
            ok, result = _run_transaction(the_doc, "Size analysis " + step_id, action)
        deleted = 0
        failed = 0
        notes = ""
        if not ok:
            notes = "transaction failed: {}".format(result)
            log("  " + notes)
        else:
            if isinstance(result, dict):
                deleted = int(result.get("deleted") or 0)
                failed = int(result.get("failed") or 0)
                notes = _safe_str(result.get("notes"))
            log("  deleted {} (failed {}) {}".format(deleted, failed, notes))

        if _step_touches_keep_view(step_id):
            keep_view = _keep()

        snap_path = snapshot(index, step_id)
        try:
            size_n = save_compact_snapshot(the_doc, snap_path, keep_view)
        except Exception as ex:
            log("  Save As failed: {}".format(ex))
            size_n = None
            notes = (notes + " | SaveAs failed: " + _safe_str(ex)).strip(" |")
        prev_bytes = record_step(index, step_id, title, snap_path, size_n, deleted, failed, notes) or prev_bytes
        maybe_delete_previous(snap_path)
        prev_snapshot = snap_path

    ranked = sorted(
        [s for s in steps if s.get("delta_bytes") is not None],
        key=lambda r: r["delta_bytes"],
    )
    payload = {
        "source_article": SOURCE_ARTICLE,
        "session_id": sessionId,
        "model_path": model_path,
        "output_dir": out_dir,
        "original_bytes": original_bytes,
        "original_mb": _mb(original_bytes),
        "inventory": inventory,
        "families": families,
        "steps": steps,
        "largest_drops": ranked[:8],
        "keep_step_files": keep_files,
        "measured_family_sizes": measure_families,
        "purified_families": purify_families,
        "purify_family_max": purify_family_max,
    }

    csv_path = os.path.join(out_dir, stem + "_filesize.csv")
    md_path = os.path.join(out_dir, stem + "_filesize.md")
    json_path = os.path.join(out_dir, stem + "_filesize.json")
    fam_csv = os.path.join(out_dir, stem + "_families.csv")
    html_path = os.path.join(out_dir, "report.html")
    html_stem = os.path.join(out_dir, stem + "_report.html")

    def _w_csv():
        write_csv(csv_path, steps, original_bytes)
        return csv_path

    def _w_md():
        write_markdown(md_path, model_path, original_bytes, steps, families)
        return md_path

    def _w_fam():
        write_family_csv(fam_csv, families)
        return fam_csv

    def _w_json():
        write_json_report(json_path, payload)
        return json_path

    def _w_html():
        write_html_report(html_path, model_path, original_bytes, steps, families, inventory)
        try:
            write_html_report(html_stem, model_path, original_bytes, steps, families, inventory)
        except Exception:
            pass
        return html_path

    _try_write("CSV", _w_csv)
    _try_write("Report", _w_md)
    _try_write("Families CSV", _w_fam)
    _try_write("JSON", _w_json)
    _try_write("HTML", _w_html)

    sidecar = os.environ.get("RBP_SIDECAR_PATH", "").strip()
    if sidecar:
        def _w_side():
            write_json_report(sidecar, payload)
            return sidecar

        _try_write("Sidecar", _w_side)

    log("=== Largest size drops ===")
    for s in ranked:
        drop = -(s.get("delta_bytes") or 0)
        if drop <= 0:
            continue
        log(
            "  {}: {:+.3f} MB ({}% of original)".format(
                s.get("title"),
                _mb(s.get("delta_bytes")),
                _pct(drop, original_bytes),
            )
        )
    return payload


# RBP runs this file as its task script; importing it (pyRevit) must not.
if RUNNING_IN_RBP and __name__ != "analyze_filesize":
    Output()
    log("=== RBP file size analysis starting ===")
    log("Article: {}".format(SOURCE_ARTICLE))
    try:
        if doc is None:
            raise RuntimeError("No script document. Run this as an RBP task script on a Revit file.")
        run(doc, revitFilePath)
    except Exception:
        log("FATAL:")
        log(traceback.format_exc())
        raise
    finally:
        log("=== RBP file size analysis complete ===")
