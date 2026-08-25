# -*- coding: utf-8 -*-
"""Create persistent isolated 3D review views from a review-package JSON."""

from __future__ import print_function

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BoundingBoxXYZ,
    ElementId,
    FilteredElementCollector,
    Transaction,
    TransactionGroup,
    View,
    View3D,
    ViewFamily,
    ViewFamilyType,
    XYZ,
)

from .schema import REVIEW_SCHEMA_VERSION


MM_TO_FEET = 1.0 / 304.8
DEFAULT_MARGIN_MM = 1000.0
DEFAULT_MAX_ELEMENTS_PER_VIEW = 150

_BUCKET_ORDER = {
    'Blocking': 0,
    'Missing': 1,
    'Geometry': 2,
    'Relationships': 3,
    'Review': 4,
}

_CHECK_LABELS = {
    'constraints': 'Constraints',
    'doors': 'Doors',
    'groups': 'Model Groups',
    'hosted_families': 'Hosted Families',
    'links': 'Revit Links',
    'mep_connectors': 'MEP Connectors',
    'rooms': 'Rooms',
    'spaces': 'Spaces',
    'walls': 'Walls',
    'windows': 'Windows',
}

_BUCKET_GUIDANCE = {
    'Blocking': (
        'BLOCKING {label}',
        'Resolve Before Approval',
        'Resolve this blocking failure before approving the mirrored model.',
    ),
    'Missing': (
        'Missing {label}',
        'Restore or Approve Deletion',
        'Restore the missing element or document approval for its deletion.',
    ),
    'Geometry': (
        'Moved {label}',
        'Reposition or Approve Movement',
        'Reposition the element or document approval for its new location.',
    ),
    'Relationships': (
        '{label} Relationships',
        'Repair or Approve Change',
        'Repair the host, constraint, group, or connector relationship, or approve the change.',
    ),
    'Review': (
        '{label} Changes',
        'Verify Design Intent',
        'Verify the change against design intent and document the decision.',
    ),
}


def _view_type_id(doc):
    for item in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if item.ViewFamily == ViewFamily.ThreeDimensional:
            return item.Id
    return None


def _existing_view_names(doc):
    names = set()
    for view in FilteredElementCollector(doc).OfClass(View):
        try:
            names.add(view.Name)
        except Exception:
            pass
    return names


def _unique_name(base_name, existing):
    candidate = base_name
    suffix = 2
    while candidate in existing:
        candidate = '{} ({})'.format(base_name, suffix)
        suffix += 1
    existing.add(candidate)
    return candidate


def _group_guidance(bucket, check):
    label = _CHECK_LABELS.get(check, 'Elements')
    issue_format, name_action, suggested_action = _BUCKET_GUIDANCE.get(
        bucket, _BUCKET_GUIDANCE['Review'])
    issue = issue_format.format(label=label)
    return {
        'issue': issue,
        'name_action': name_action,
        'suggested_action': suggested_action,
    }


def _bbox_points(element):
    try:
        box = element.get_BoundingBox(None)
    except Exception:
        box = None
    if box is None:
        return []
    transform = box.Transform
    points = []
    for x in (box.Min.X, box.Max.X):
        for y in (box.Min.Y, box.Max.Y):
            for z in (box.Min.Z, box.Max.Z):
                points.append(transform.OfPoint(XYZ(x, y, z)))
    return points


def _section_box(elements, margin_mm):
    points = []
    for element in elements:
        points.extend(_bbox_points(element))
    if not points:
        return None
    margin = float(margin_mm) * MM_TO_FEET
    box = BoundingBoxXYZ()
    box.Min = XYZ(
        min(point.X for point in points) - margin,
        min(point.Y for point in points) - margin,
        min(point.Z for point in points) - margin,
    )
    box.Max = XYZ(
        max(point.X for point in points) + margin,
        max(point.Y for point in points) + margin,
        max(point.Z for point in points) + margin,
    )
    return box


def _create_view_shell(doc, view_type_id, name, elements, margin_mm,
                       section_failures):
    view = View3D.CreateIsometric(doc, view_type_id)
    view.Name = name
    box = _section_box(elements, margin_mm)
    if box is not None:
        try:
            view.SetSectionBox(box)
            view.IsSectionBoxActive = True
        except Exception as ex:
            section_failures.append({
                'view': name,
                'error': str(ex),
            })
    return view


def _chunks(items, size):
    for index in range(0, len(items), size):
        yield items[index:index + size]


def _resolve_targets(doc, package):
    grouped = {}
    unresolved = []
    non_graphical = []
    resolved_count = 0
    for target in package.get('targets') or []:
        uid = target.get('unique_id')
        if not uid:
            continue
        try:
            element = doc.GetElement(uid)
        except Exception:
            element = None
        if element is None:
            unresolved.append(uid)
            continue
        resolved_count += 1
        if not _bbox_points(element):
            non_graphical.append(uid)
            continue
        bucket = target.get('bucket') or 'Review'
        checks = target.get('checks') or []
        check = checks[0] if checks else 'elements'
        grouped.setdefault((bucket, check), []).append(element)
    return grouped, resolved_count, unresolved, non_graphical


def _isolation_ids(doc, target_elements, view):
    ids = List[ElementId]()
    seen = set()
    non_hideable_targets = []

    def append_element(element):
        uid = element.UniqueId
        if uid in seen:
            return False
        seen.add(uid)
        added = False
        try:
            if element.CanBeHidden(view):
                ids.Add(element.Id)
                added = True
        except Exception:
            pass
        try:
            member_ids = element.GetMemberIds()
        except Exception:
            member_ids = []
        for member_id in member_ids:
            member = doc.GetElement(member_id)
            if member is not None and append_element(member):
                added = True
        return added

    for target in target_elements:
        if not append_element(target):
            non_hideable_targets.append(target.UniqueId)
    return ids, non_hideable_targets


def create_review_views(doc, package, margin_mm=DEFAULT_MARGIN_MM,
                        max_elements_per_view=DEFAULT_MAX_ELEMENTS_PER_VIEW):
    """Create persistent isolated views; return audit counts and names."""
    if package.get('schema_version') != REVIEW_SCHEMA_VERSION:
        raise ValueError(
            'Review package schema {} does not match {}.'.format(
                package.get('schema_version'), REVIEW_SCHEMA_VERSION))
    view_type_id = _view_type_id(doc)
    if view_type_id is None:
        raise ValueError('No isometric 3D ViewFamilyType is available.')

    grouped, resolved_count, unresolved, non_graphical = _resolve_targets(doc, package)
    existing_names = _existing_view_names(doc)
    created = []
    isolation_failures = []
    section_failures = []
    work_items = []

    transaction_group = TransactionGroup(doc, 'Mirror Project: Create Review Views')
    transaction_group.Start()
    try:
        create_transaction = Transaction(doc, 'Mirror Project: Create Review View Shells')
        create_transaction.Start()
        try:
            group_keys = sorted(
                grouped.keys(),
                key=lambda key: (_BUCKET_ORDER.get(key[0], 99), key[1]),
            )
            all_elements = []
            for group_key in group_keys:
                all_elements.extend(grouped.get(group_key) or [])
            if all_elements:
                all_name = _unique_name(
                    'MP Review - ALL Missing and Review Elements - Select Individually',
                    existing_names,
                )
                all_view = _create_view_shell(
                    doc,
                    view_type_id,
                    all_name,
                    all_elements,
                    margin_mm,
                    section_failures,
                )
                all_item = {
                    'name': all_name,
                    'element_count': 0,
                    'target_count': len(all_elements),
                    'bucket': 'All',
                    'check': 'all',
                    'issue': 'ALL Missing and Review Elements',
                    'suggested_action': (
                        'Select each isolated element, inspect its recorded issue, '
                        'then repair it or document approval for the change.'
                    ),
                    'is_aggregate': True,
                }
                created.append(all_item)
                work_items.append((all_view, all_elements, all_item))

            for bucket, check in group_keys:
                elements = grouped.get((bucket, check)) or []
                guidance = _group_guidance(bucket, check)
                batches = list(_chunks(elements, max_elements_per_view))
                for index, batch in enumerate(batches):
                    suffix = ' {:02d}'.format(index + 1) if len(batches) > 1 else ''
                    name = _unique_name(
                        'MP Review - {} - {}{}'.format(
                            guidance.get('issue'),
                            guidance.get('name_action'),
                            suffix,
                        ),
                        existing_names,
                    )
                    view = _create_view_shell(
                        doc,
                        view_type_id,
                        name,
                        batch,
                        margin_mm,
                        section_failures,
                    )
                    item = {
                        'name': name,
                        'element_count': 0,
                        'target_count': len(batch),
                        'bucket': bucket,
                        'check': check,
                        'issue': guidance.get('issue'),
                        'suggested_action': guidance.get('suggested_action'),
                        'is_aggregate': False,
                    }
                    created.append(item)
                    work_items.append((view, batch, item))
            create_transaction.Commit()
        except Exception:
            create_transaction.RollBack()
            raise
        finally:
            create_transaction.Dispose()

        configure_transaction = Transaction(
            doc, 'Mirror Project: Configure Review View Isolation')
        configure_transaction.Start()
        try:
            for view, batch, item in work_items:
                ids, non_hideable = _isolation_ids(doc, batch, view)
                non_graphical.extend(non_hideable)
                item['element_count'] = ids.Count
                if not ids.Count:
                    continue
                try:
                    view.IsolateElementsTemporary(ids)
                    view.ConvertTemporaryHideIsolateToPermanent()
                    try:
                        if view.IsTemporaryHideIsolateActive():
                            isolation_failures.append({
                                'view': item.get('name'),
                                'error': 'Temporary isolation remained active.',
                            })
                    except Exception:
                        pass
                except Exception as ex:
                    isolation_failures.append({
                        'view': item.get('name'),
                        'error': str(ex),
                    })
            configure_transaction.Commit()
        except Exception:
            configure_transaction.RollBack()
            raise
        finally:
            configure_transaction.Dispose()
        transaction_group.Assimilate()
    except Exception:
        transaction_group.RollBack()
        raise
    finally:
        transaction_group.Dispose()

    return {
        'requested': len(package.get('targets') or []),
        'resolved': resolved_count,
        'unresolved': sorted(set(unresolved)),
        'non_graphical': sorted(set(non_graphical)),
        'created_views': created,
        'isolation_failures': isolation_failures,
        'section_failures': section_failures,
    }
