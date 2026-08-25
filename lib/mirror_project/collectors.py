# -*- coding: utf-8 -*-
"""Revit collectors for Mirror Project snapshots.

Each public collect_* function returns a data dict. Feature-missing APIs
set ``availability`` to ``not_available`` instead of failing silently.
"""

from __future__ import print_function

import math

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    FilteredWorksetCollector,
    StorageType,
    View,
    ViewSheet,
    Wall,
    WorksetKind,
)
from Autodesk.Revit.DB import ImportInstance, Family, FamilyInstance, Group, Level
from Autodesk.Revit.DB import RevitLinkInstance, RevitLinkType

from revit.compat import get_element_id_value

from .geometry import aabb_from_corners, bbox_corners, round_xyz, xyz_mm_from_feet
from .schema import COORDINATE_FRAME_INTERNAL, STATUS_NOT_AVAILABLE, UNITS_MM

_INVALID_GROUP = -1
_SAMPLE_LIMIT = 20


def _uid(element):
    try:
        return element.UniqueId
    except Exception:
        return None


def _cat_name(element):
    try:
        if element.Category:
            return element.Category.Name
    except Exception:
        pass
    return ''


def _id_value(element_id):
    try:
        return get_element_id_value(element_id)
    except Exception:
        return None


def _is_grouped(element):
    try:
        gid = element.GroupId
        if gid is None:
            return False
        if gid == ElementId.InvalidElementId:
            return False
        return _id_value(gid) not in (None, _INVALID_GROUP)
    except Exception:
        return False


def _param_int(element, bip):
    try:
        param = element.get_Parameter(bip)
        if param and param.HasValue and param.StorageType == StorageType.Integer:
            return param.AsInteger()
    except Exception:
        pass
    return None


def _param_str(element, bip):
    try:
        param = element.get_Parameter(bip)
        if param and param.HasValue:
            if param.StorageType == StorageType.String:
                return param.AsString() or ''
            return param.AsValueString() or ''
    except Exception:
        pass
    return ''


def _location_mm(element):
    try:
        loc = element.Location
        if loc is None:
            return None
        if hasattr(loc, 'Point') and loc.Point is not None:
            return round_xyz(xyz_mm_from_feet(loc.Point))
        if hasattr(loc, 'Curve') and loc.Curve is not None:
            curve = loc.Curve
            return round_xyz(xyz_mm_from_feet(curve.GetEndPoint(0)))
    except Exception:
        pass
    return None


def _curve_ends_mm(element):
    try:
        loc = element.Location
        if loc is None or not hasattr(loc, 'Curve') or loc.Curve is None:
            return None, None
        curve = loc.Curve
        start = round_xyz(xyz_mm_from_feet(curve.GetEndPoint(0)))
        end = round_xyz(xyz_mm_from_feet(curve.GetEndPoint(1)))
        return start, end
    except Exception:
        return None, None


def _xyz_list(xyz):
    mm = xyz_mm_from_feet(xyz)
    return round_xyz(mm) if mm is not None else None


def collect_identity(doc):
    app = doc.Application
    fingerprint = {
        'title': doc.Title,
        'path_name': doc.PathName or '',
        'is_workshared': bool(doc.IsWorkshared),
        'is_detached': _safe_bool(doc, 'IsDetached'),
        'is_cloud': _safe_bool(doc, 'IsModelInCloud'),
        'is_family': bool(doc.IsFamilyDocument),
        'revit_version': _safe_str(app, 'VersionNumber'),
        'revit_build': _safe_str(app, 'VersionBuild'),
        'username': _safe_str(app, 'Username'),
        'central_guid': '',
    }
    try:
        if doc.IsWorkshared and hasattr(doc, 'GetWorksharingCentralModelPath'):
            fingerprint['central_guid'] = str(doc.WorksharingCentralGUID)
    except Exception:
        fingerprint['central_guid'] = STATUS_NOT_AVAILABLE
    return fingerprint


def _safe_bool(obj, name):
    try:
        if hasattr(obj, name):
            return bool(getattr(obj, name))
    except Exception:
        pass
    return STATUS_NOT_AVAILABLE


def _safe_str(obj, name):
    try:
        value = getattr(obj, name)
        return str(value) if value is not None else ''
    except Exception:
        return ''


def collect_worksets(doc):
    if not doc.IsWorkshared:
        return {
            'workshared': False,
            'names': [],
            'items': [],
        }
    items = []
    names = []
    collector = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    for ws in collector:
        name = ws.Name
        names.append(name)
        items.append({
            'name': name,
            'id': _workset_id_value(ws),
            'visible_by_default': bool(ws.IsVisibleByDefault),
            'kind': str(ws.Kind),
        })
    names.sort()
    return {
        'workshared': True,
        'names': names,
        'items': items,
    }


def _workset_id_value(workset):
    wid = workset.Id
    if hasattr(wid, 'IntegerValue'):
        try:
            return int(wid.IntegerValue)
        except Exception:
            pass
    if hasattr(wid, 'Value'):
        try:
            return int(wid.Value)
        except Exception:
            pass
    return str(wid)


def collect_warnings(doc):
    if not hasattr(doc, 'GetWarnings'):
        return {'availability': STATUS_NOT_AVAILABLE, 'items': [], 'total': 0}
    grouped = {}
    total = 0
    for warning in doc.GetWarnings():
        total += 1
        guid = ''
        description = ''
        severity = ''
        failing = []
        additional = []
        try:
            definition = warning.GetFailureDefinitionId()
            if definition is not None and hasattr(definition, 'Guid'):
                guid = str(definition.Guid)
        except Exception:
            guid = STATUS_NOT_AVAILABLE
        try:
            description = warning.GetDescriptionText() or ''
        except Exception:
            description = ''
        try:
            if hasattr(warning, 'GetSeverity'):
                severity = str(warning.GetSeverity())
        except Exception:
            severity = STATUS_NOT_AVAILABLE
        failing = _element_id_list(warning, 'GetFailingElements')
        if not failing:
            failing = _element_id_list(warning, 'GetFailingElementIds')
        additional = _element_id_list(warning, 'GetAdditionalElements')
        if not additional:
            additional = _element_id_list(warning, 'GetAdditionalElementIds')
        key = guid or description
        bucket = grouped.get(key)
        if bucket is None:
            bucket = {
                'failure_guid': guid,
                'description': description,
                'severity': severity,
                'count': 0,
                'affected_count': 0,
                'failing_ids': [],
            }
            grouped[key] = bucket
        bucket['count'] += 1
        bucket['affected_count'] += len(failing)
        for eid in failing:
            if eid not in bucket['failing_ids'] and len(bucket['failing_ids']) < 50:
                bucket['failing_ids'].append(eid)
    items = [grouped[k] for k in sorted(grouped.keys())]
    return {
        'total': total,
        'group_count': len(items),
        'items': items,
    }


def _element_id_list(warning, method_name):
    ids = []
    try:
        method = getattr(warning, method_name, None)
        if method is None:
            return ids
        for eid in method():
            value = _id_value(eid)
            if value is not None:
                ids.append(value)
    except Exception:
        pass
    return ids


def collect_coordinates(doc):
    data = {
        'units': UNITS_MM,
        'coordinate_frame': COORDINATE_FRAME_INTERNAL,
        'internal_origin_mm': [0.0, 0.0, 0.0],
        'project_base_point_mm': None,
        'survey_point_mm': None,
        'survey_shared_mm': None,
        'survey_clipped': STATUS_NOT_AVAILABLE,
        'true_north_deg': None,
        'active_location': '',
        'locations': [],
    }
    pbp, survey = _base_points(doc)
    if pbp is not None:
        try:
            data['project_base_point_mm'] = _xyz_list(pbp.Position)
        except Exception:
            data['project_base_point_mm'] = STATUS_NOT_AVAILABLE
    else:
        data['project_base_point_mm'] = STATUS_NOT_AVAILABLE
    if survey is not None:
        try:
            data['survey_point_mm'] = _xyz_list(survey.Position)
        except Exception:
            data['survey_point_mm'] = STATUS_NOT_AVAILABLE
        try:
            data['survey_shared_mm'] = _xyz_list(survey.SharedPosition)
        except Exception:
            data['survey_shared_mm'] = STATUS_NOT_AVAILABLE
        try:
            if hasattr(survey, 'Clipped'):
                data['survey_clipped'] = bool(survey.Clipped)
        except Exception:
            data['survey_clipped'] = STATUS_NOT_AVAILABLE
    else:
        data['survey_point_mm'] = STATUS_NOT_AVAILABLE

    locations = []
    try:
        active = doc.ActiveProjectLocation
        if active is not None:
            data['active_location'] = active.Name
    except Exception:
        data['active_location'] = STATUS_NOT_AVAILABLE
    try:
        for loc in doc.ProjectLocations:
            item = {'name': loc.Name}
            try:
                from Autodesk.Revit.DB import XYZ
                pos = loc.GetProjectPosition(XYZ.Zero)
                item['east_west_mm'] = round(pos.EastWest * 304.8, 6)
                item['north_south_mm'] = round(pos.NorthSouth * 304.8, 6)
                item['elevation_mm'] = round(pos.Elevation * 304.8, 6)
                angle_deg = pos.Angle * 180.0 / math.pi
                item['angle_deg'] = round(angle_deg, 6)
                if loc.Name == data['active_location'] or not data['true_north_deg']:
                    data['true_north_deg'] = round(angle_deg, 6)
            except Exception as ex:
                item['error'] = str(ex)
            try:
                tf = loc.GetTotalTransform()
                item['transform'] = _serialize_transform(tf)
            except Exception:
                item['transform'] = STATUS_NOT_AVAILABLE
            locations.append(item)
    except Exception:
        data['locations'] = STATUS_NOT_AVAILABLE
        return data
    data['locations'] = locations
    return data


def _base_points(doc):
    pbp = None
    survey = None
    try:
        from Autodesk.Revit.DB import BasePoint
    except Exception:
        return None, None
    try:
        if hasattr(BasePoint, 'GetProjectBasePoint'):
            pbp = BasePoint.GetProjectBasePoint(doc)
        if hasattr(BasePoint, 'GetSurveyPoint'):
            survey = BasePoint.GetSurveyPoint(doc)
        if pbp is not None or survey is not None:
            return pbp, survey
    except Exception:
        pass
    try:
        for bp in FilteredElementCollector(doc).OfClass(BasePoint):
            try:
                if bp.IsShared:
                    survey = bp
                else:
                    pbp = bp
            except Exception:
                continue
    except Exception:
        pass
    return pbp, survey


def _serialize_transform(tf):
    if tf is None:
        return STATUS_NOT_AVAILABLE
    data = {
        'origin_mm': _xyz_list(tf.Origin),
        'basis_x': round_xyz((tf.BasisX.X, tf.BasisX.Y, tf.BasisX.Z)),
        'basis_y': round_xyz((tf.BasisY.X, tf.BasisY.Y, tf.BasisY.Z)),
        'basis_z': round_xyz((tf.BasisZ.X, tf.BasisZ.Y, tf.BasisZ.Z)),
    }
    try:
        data['determinant'] = round(float(tf.Determinant), 6)
    except Exception:
        data['determinant'] = STATUS_NOT_AVAILABLE
    return data


def collect_links(doc):
    items = []
    for link in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        item = {
            'unique_id': _uid(link),
            'element_id': _id_value(link.Id),
            'name': '',
            'loaded': False,
            'pinned': _safe_bool(link, 'Pinned'),
            'transform': STATUS_NOT_AVAILABLE,
            'bbox': STATUS_NOT_AVAILABLE,
            'shared_site': STATUS_NOT_AVAILABLE,
            'room_bounding': STATUS_NOT_AVAILABLE,
        }
        try:
            item['name'] = link.Name
        except Exception:
            pass
        try:
            link_type = doc.GetElement(link.GetTypeId())
            if isinstance(link_type, RevitLinkType):
                item['loaded'] = bool(RevitLinkType.IsLoaded(doc, link_type.Id))
        except Exception:
            try:
                item['loaded'] = link.GetLinkDocument() is not None
            except Exception:
                item['loaded'] = STATUS_NOT_AVAILABLE
        try:
            item['transform'] = _serialize_transform(link.GetTotalTransform())
        except Exception:
            item['transform'] = STATUS_NOT_AVAILABLE
        try:
            bb = link.get_BoundingBox(None)
            item['bbox'] = _bbox_payload(bb)
        except Exception:
            item['bbox'] = STATUS_NOT_AVAILABLE
        try:
            param = link.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING)
            if param:
                item['room_bounding'] = bool(param.AsInteger())
        except Exception:
            pass
        items.append(item)
    return {
        'count': len(items),
        'items': items,
    }


def _bbox_payload(bb):
    if bb is None:
        return None
    try:
        bx = (bb.Transform.BasisX.X, bb.Transform.BasisX.Y, bb.Transform.BasisX.Z)
        by = (bb.Transform.BasisY.X, bb.Transform.BasisY.Y, bb.Transform.BasisY.Z)
        bz = (bb.Transform.BasisZ.X, bb.Transform.BasisZ.Y, bb.Transform.BasisZ.Z)
        mn = xyz_mm_from_feet(bb.Min)
        mx = xyz_mm_from_feet(bb.Max)
        o_mm = xyz_mm_from_feet(bb.Transform.Origin)
        corners = bbox_corners(mn, mx, local_origin=o_mm, basis_x=bx, basis_y=by, basis_z=bz)
        return aabb_from_corners(corners)
    except Exception:
        return STATUS_NOT_AVAILABLE


def _room_state(element):
    location = None
    try:
        location = element.Location
    except Exception:
        location = None
    area = 0.0
    try:
        area = float(element.Area)
    except Exception:
        area = 0.0
    has_boundary = False
    try:
        from Autodesk.Revit.DB import SpatialElementBoundaryOptions
        opts = SpatialElementBoundaryOptions()
        loops = element.GetBoundarySegments(opts)
        if loops:
            for loop in loops:
                if loop and len(list(loop)) > 0:
                    has_boundary = True
                    break
    except Exception:
        has_boundary = STATUS_NOT_AVAILABLE
    if location is None:
        return 'unplaced'
    if area == 0.0 and has_boundary is True:
        return 'redundant'
    if area == 0.0 and has_boundary is False:
        return 'not_enclosed'
    if area == 0.0 and has_boundary == STATUS_NOT_AVAILABLE:
        return 'not_enclosed'
    return 'placed'


def _spatial_item(element):
    return {
        'unique_id': _uid(element),
        'element_id': _id_value(element.Id),
        'name': _param_str(element, BuiltInParameter.ROOM_NAME),
        'number': _param_str(element, BuiltInParameter.ROOM_NUMBER),
        'state': _room_state(element),
        'area': _safe_area(element),
        'level': _level_name(element),
    }


def _safe_area(element):
    try:
        return round(float(element.Area), 6)
    except Exception:
        return STATUS_NOT_AVAILABLE


def _level_name(element):
    try:
        level_id = element.LevelId
        if level_id and level_id != ElementId.InvalidElementId:
            level = element.Document.GetElement(level_id)
            if level:
                return level.Name
    except Exception:
        pass
    try:
        return element.Level.Name
    except Exception:
        return ''


def collect_rooms(doc):
    return _collect_spatial(doc, BuiltInCategory.OST_Rooms)


def collect_spaces(doc):
    return _collect_spatial(doc, BuiltInCategory.OST_MEPSpaces)


def collect_areas(doc):
    return _collect_spatial(doc, BuiltInCategory.OST_Areas)


def _collect_spatial(doc, bic):
    items = []
    by_state = {
        'placed': 0,
        'unplaced': 0,
        'not_enclosed': 0,
        'redundant': 0,
    }
    collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
    for element in collector:
        item = _spatial_item(element)
        items.append(item)
        state = item.get('state')
        if state in by_state:
            by_state[state] += 1
    return {
        'count': len(items),
        'by_state': by_state,
        'items': items,
    }


def collect_constraints(doc):
    items = []
    locked_count = 0
    by_class = {}
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_Constraints
    ).WhereElementIsNotElementType()
    for element in collector:
        record = classify_constraint(element)
        items.append(record)
        cls = record.get('classification') or 'unsupported_type'
        by_class[cls] = by_class.get(cls, 0) + 1
        if record.get('is_locked'):
            locked_count += 1
    return {
        'count': len(items),
        'locked_count': locked_count,
        'by_class': by_class,
        'items': items,
    }


def classify_constraint(element):
    from Autodesk.Revit.DB import Dimension

    record = {
        'unique_id': _uid(element),
        'element_id': _id_value(element.Id),
        'category': _cat_name(element),
        'is_locked': False,
        'number_of_segments': STATUS_NOT_AVAILABLE,
        'grouped': _is_grouped(element),
        'labeled': False,
        'classification': 'unsupported_type',
        'reason': 'Not a Dimension constraint',
    }
    if element is None:
        record['classification'] = 'missing'
        record['reason'] = 'Element is None'
        return record
    try:
        if not element.IsValidObject:
            record['classification'] = 'missing'
            record['reason'] = 'Invalid element'
            return record
    except Exception:
        pass
    if not isinstance(element, Dimension):
        return record
    try:
        record['is_locked'] = bool(element.IsLocked)
    except Exception:
        record['is_locked'] = STATUS_NOT_AVAILABLE
        record['classification'] = 'unmodifiable'
        record['reason'] = 'IsLocked unavailable'
        return record
    try:
        record['number_of_segments'] = int(element.NumberOfSegments)
    except Exception:
        record['number_of_segments'] = STATUS_NOT_AVAILABLE
    try:
        label = None
        if hasattr(element, 'FamilyLabel'):
            label = element.FamilyLabel
        record['labeled'] = label is not None
    except Exception:
        record['labeled'] = STATUS_NOT_AVAILABLE
    if record['grouped']:
        record['classification'] = 'grouped'
        record['reason'] = 'Constraint is grouped'
        return record
    segs = record['number_of_segments']
    if segs != STATUS_NOT_AVAILABLE and segs > 1:
        record['classification'] = 'multi_segment'
        record['reason'] = 'Linear dimension has more than one segment'
        return record
    if record['labeled'] is True:
        record['classification'] = 'labeled'
        record['reason'] = 'Dimension is labeled'
        return record
    if not record['is_locked']:
        record['classification'] = 'eligible'
        record['reason'] = 'Unlocked Dimension'
        return record
    record['classification'] = 'eligible'
    record['reason'] = 'API-eligible locked Dimension'
    return record


def collect_model_scan(doc):
    """Single-pass instance scan for counts, walls, families, groups, MEP, extents."""
    counts = {
        'instances': 0,
        'types': 0,
        'views': 0,
        'sheets': 0,
        'groups': 0,
        'imports': 0,
        'families': 0,
        'by_category': {},
    }
    walls = []
    doors = []
    windows = []
    hosted = []
    groups = []
    mep_items = []
    unused_total = 0
    samples = []
    min_pt = None
    max_pt = None

    try:
        type_count = FilteredElementCollector(doc).WhereElementIsElementType().GetElementCount()
        counts['types'] = int(type_count)
    except Exception:
        counts['types'] = len(list(FilteredElementCollector(doc).WhereElementIsElementType().ToElementIds()))

    try:
        family_count = FilteredElementCollector(doc).OfClass(Family).GetElementCount()
        counts['families'] = int(family_count)
    except Exception:
        counts['families'] = len(list(FilteredElementCollector(doc).OfClass(Family).ToElementIds()))

    instances = FilteredElementCollector(doc).WhereElementIsNotElementType()
    for element in instances:
        counts['instances'] += 1
        cat = _cat_name(element) or 'No category'
        counts['by_category'][cat] = counts['by_category'].get(cat, 0) + 1

        if isinstance(element, ViewSheet):
            counts['sheets'] += 1
        elif isinstance(element, View):
            counts['views'] += 1
        if isinstance(element, ImportInstance):
            counts['imports'] += 1
        if isinstance(element, Group):
            counts['groups'] += 1
            groups.append(_group_item(doc, element))

        if isinstance(element, Wall):
            wall_item = _wall_item(doc, element)
            walls.append(wall_item)
            min_pt, max_pt = _expand_extents(min_pt, max_pt, wall_item.get('start_mm'))
            min_pt, max_pt = _expand_extents(min_pt, max_pt, wall_item.get('end_mm'))
            if len(samples) < _SAMPLE_LIMIT and wall_item.get('start_mm'):
                samples.append({
                    'id': wall_item.get('unique_id'),
                    'point_mm': wall_item.get('start_mm'),
                })

        if isinstance(element, FamilyInstance):
            cat_id = None
            try:
                cat_id = element.Category.Id
            except Exception:
                cat_id = None
            is_door = _is_bic(cat_id, BuiltInCategory.OST_Doors)
            is_window = _is_bic(cat_id, BuiltInCategory.OST_Windows)
            fam_item = _family_item(doc, element)
            if is_door:
                doors.append(fam_item)
            elif is_window:
                windows.append(fam_item)
            elif _is_risk_family(fam_item):
                hosted.append(fam_item)
            loc = fam_item.get('location_mm')
            min_pt, max_pt = _expand_extents(min_pt, max_pt, loc)
            if (is_door or is_window) and loc and len(samples) < _SAMPLE_LIMIT * 2:
                samples.append({'id': fam_item.get('unique_id'), 'point_mm': loc})

        unused, mep_item = _mep_unused(element)
        if unused:
            unused_total += unused
            mep_items.append(mep_item)

    levels = list(FilteredElementCollector(doc).OfClass(Level))
    return {
        'counts': {
            'instances': counts['instances'],
            'types': counts['types'],
            'views': counts['views'],
            'sheets': counts['sheets'],
            'groups': counts['groups'],
            'imports': counts['imports'],
            'families': counts['families'],
            'doors': len(doors),
            'windows': len(windows),
            'walls': len(walls),
            'by_category': counts['by_category'],
        },
        'walls': {
            'count': len(walls),
            'attached_count': sum(1 for w in walls if w.get('attach_top') or w.get('attach_base')),
            'items': walls,
        },
        'doors': {'count': len(doors), 'items': doors},
        'windows': {'count': len(windows), 'items': windows},
        'hosted_families': {'count': len(hosted), 'items': hosted},
        'groups': {
            'instance_count': len(groups),
            'nested_count': sum(1 for g in groups if g.get('nested')),
            'items': groups,
        },
        'mep_connectors': {
            'unused_count': unused_total,
            'element_count': len(mep_items),
            'items': mep_items,
        },
        'extents': {
            'min_mm': list(min_pt) if min_pt else None,
            'max_mm': list(max_pt) if max_pt else None,
            'level_count': len(levels),
            'levels': [_level_item(lv) for lv in levels],
        },
        'samples': {'points': samples},
    }


def _is_bic(cat_id, bic):
    if cat_id is None:
        return False
    try:
        return _id_value(cat_id) == int(bic)
    except Exception:
        return False


def _expand_extents(min_pt, max_pt, point):
    xyz = point
    if not xyz or len(xyz) < 3:
        return min_pt, max_pt
    if min_pt is None:
        return (xyz[0], xyz[1], xyz[2]), (xyz[0], xyz[1], xyz[2])
    return (
        (min(min_pt[0], xyz[0]), min(min_pt[1], xyz[1]), min(min_pt[2], xyz[2])),
        (max(max_pt[0], xyz[0]), max(max_pt[1], xyz[1]), max(max_pt[2], xyz[2])),
    )


def _level_item(level):
    item = {
        'unique_id': _uid(level),
        'name': level.Name,
        'elevation_mm': None,
    }
    try:
        item['elevation_mm'] = round(float(level.Elevation) * 304.8, 6)
    except Exception:
        item['elevation_mm'] = STATUS_NOT_AVAILABLE
    return item


def _wall_item(doc, wall):
    start, end = _curve_ends_mm(wall)
    item = {
        'unique_id': _uid(wall),
        'element_id': _id_value(wall.Id),
        'start_mm': start,
        'end_mm': end,
        'attach_top': [],
        'attach_base': [],
        'grouped': _is_grouped(wall),
        'curtain': STATUS_NOT_AVAILABLE,
    }
    try:
        item['curtain'] = wall.CurtainGrid is not None
    except Exception:
        pass
    top_ids, base_ids, attach_status = _wall_attachments(doc, wall)
    if attach_status == STATUS_NOT_AVAILABLE:
        item['attach_top'] = STATUS_NOT_AVAILABLE
        item['attach_base'] = STATUS_NOT_AVAILABLE
    else:
        item['attach_top'] = top_ids
        item['attach_base'] = base_ids
    return item


def _wall_attachments(doc, wall):
    if hasattr(wall, 'GetAttachmentIds'):
        try:
            from Autodesk.Revit.DB import AttachmentLocation
            top = _host_uids(doc, wall.GetAttachmentIds(AttachmentLocation.Top))
            base = _host_uids(doc, wall.GetAttachmentIds(AttachmentLocation.Base))
            return top, base, STATUS_OK_ATTACH
        except Exception:
            pass
    # Parameter fallback: presence only, hosts unknown.
    top_flag = _param_int(wall, BuiltInParameter.WALL_TOP_IS_ATTACHED)
    base_flag = _param_int(wall, BuiltInParameter.WALL_BOTTOM_IS_ATTACHED)
    top = ['attached'] if top_flag == 1 else []
    base = ['attached'] if base_flag == 1 else []
    if top_flag is None and base_flag is None:
        return [], [], STATUS_NOT_AVAILABLE
    return top, base, 'parameter'


STATUS_OK_ATTACH = 'ok'


def _host_uids(doc, ids):
    uids = []
    if not ids:
        return uids
    for eid in ids:
        try:
            host = doc.GetElement(eid)
            uid = _uid(host) if host is not None else None
            if uid:
                uids.append(uid)
            else:
                uids.append(str(_id_value(eid)))
        except Exception:
            continue
    return uids


def _family_item(doc, fi):
    item = {
        'unique_id': _uid(fi),
        'element_id': _id_value(fi.Id),
        'category': _cat_name(fi),
        'location_mm': _location_mm(fi),
        'mirrored': _safe_bool(fi, 'Mirrored'),
        'hand_flipped': _safe_bool(fi, 'HandFlipped'),
        'facing_flipped': _safe_bool(fi, 'FacingFlipped'),
        'workplane_flipped': _safe_bool(fi, 'IsWorkPlaneFlipped'),
        'facing': STATUS_NOT_AVAILABLE,
        'hand': STATUS_NOT_AVAILABLE,
        'hosted': False,
        'host_uid': '',
        'from_room': '',
        'to_room': '',
        'grouped': _is_grouped(fi),
    }
    try:
        facing = fi.FacingOrientation
        item['facing'] = round_xyz((facing.X, facing.Y, facing.Z))
    except Exception:
        pass
    try:
        hand = fi.HandOrientation
        item['hand'] = round_xyz((hand.X, hand.Y, hand.Z))
    except Exception:
        pass
    try:
        host = fi.Host
        if host is not None:
            item['hosted'] = True
            item['host_uid'] = _uid(host) or ''
    except Exception:
        pass
    _fill_from_to_room(doc, fi, item)
    return item


def _is_risk_family(item):
    if item.get('hosted'):
        return True
    for flag in ('mirrored', 'hand_flipped', 'facing_flipped', 'workplane_flipped'):
        if item.get(flag) is True:
            return True
    return False


def _fill_from_to_room(doc, fi, item):
    phase = None
    try:
        created = fi.CreatedPhaseId
        if created and created != ElementId.InvalidElementId:
            phase = doc.GetElement(created)
    except Exception:
        phase = None
    if phase is None:
        try:
            phase = doc.ActiveView.CreatedPhaseId
            phase = doc.GetElement(phase)
        except Exception:
            phase = None
    try:
        if phase is not None and hasattr(fi, 'get_FromRoom'):
            room = fi.get_FromRoom(phase)
            if room:
                item['from_room'] = _uid(room) or ''
        if phase is not None and hasattr(fi, 'get_ToRoom'):
            room = fi.get_ToRoom(phase)
            if room:
                item['to_room'] = _uid(room) or ''
    except Exception:
        item['from_room'] = item.get('from_room') or STATUS_NOT_AVAILABLE
        item['to_room'] = item.get('to_room') or STATUS_NOT_AVAILABLE


def _group_item(doc, group):
    members = []
    nested = False
    try:
        members = list(group.GetMemberIds())
        for mid in members:
            member = doc.GetElement(mid)
            if isinstance(member, Group):
                nested = True
                break
    except Exception:
        members = []
        nested = STATUS_NOT_AVAILABLE
    type_name = ''
    try:
        type_name = group.GroupType.Name
    except Exception:
        pass
    return {
        'unique_id': _uid(group),
        'element_id': _id_value(group.Id),
        'type_name': type_name,
        'member_count': len(members) if members is not None else STATUS_NOT_AVAILABLE,
        'nested': nested,
        'location_mm': _location_mm(group),
    }


def _mep_unused(element):
    manager = _connector_manager(element)
    if manager is None:
        return 0, None
    unused = 0
    try:
        for connector in manager.Connectors:
            try:
                if not connector.IsConnected:
                    unused += 1
            except Exception:
                continue
    except Exception:
        return 0, None
    if unused <= 0:
        return 0, None
    return unused, {
        'unique_id': _uid(element),
        'element_id': _id_value(element.Id),
        'category': _cat_name(element),
        'unused': unused,
    }


def _connector_manager(element):
    try:
        if hasattr(element, 'ConnectorManager') and element.ConnectorManager:
            return element.ConnectorManager
    except Exception:
        pass
    try:
        mep = element.MEPModel
        if mep is not None and mep.ConnectorManager:
            return mep.ConnectorManager
    except Exception:
        pass
    return None
