# -*- coding: utf-8 -*-
"""Classify and disable eligible locked constraints. Do not restore."""

from __future__ import print_function

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Dimension,
    FilteredElementCollector,
    Transaction,
    TransactionStatus,
)

from revit.revit_utils import is_element_editable

from .collectors import classify_constraint


def classify_document_constraints(doc):
    """Return classified constraint records for the document."""
    buckets = {
        'eligible': [],
        'grouped': [],
        'multi_segment': [],
        'labeled': [],
        'unsupported_type': [],
        'missing': [],
        'unmodifiable': [],
        'unlocked': [],
    }
    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_Constraints
    ).WhereElementIsNotElementType()
    for element in collector:
        record = classify_constraint(element)
        if record.get('classification') == 'eligible' and not record.get('is_locked'):
            buckets['unlocked'].append(record)
            continue
        if record.get('classification') == 'eligible' and record.get('is_locked'):
            editable, reason = _editable(doc, element)
            if not editable:
                record['classification'] = 'unmodifiable'
                record['reason'] = reason
                buckets['unmodifiable'].append(record)
                continue
        cls = record.get('classification') or 'unsupported_type'
        if cls not in buckets:
            buckets['unsupported_type'].append(record)
        else:
            buckets[cls].append(record)
    return buckets


def _editable(doc, element):
    try:
        return is_element_editable(doc, element)
    except Exception as ex:
        return False, 'Editability check failed: {}'.format(ex)


def disable_eligible_constraints(doc, eligible_records):
    """Unlock eligible locked dimensions. Returns committed/failed/skipped."""
    committed = []
    failed = []
    skipped = []
    if not eligible_records:
        return {
            'committed': committed,
            'failed': failed,
            'skipped': skipped,
            'transaction_status': 'skipped',
        }

    targets = []
    for record in eligible_records:
        uid = record.get('unique_id')
        element = None
        if uid:
            try:
                element = doc.GetElement(uid)
            except Exception:
                element = None
        if element is None:
            skipped.append(_with_status(record, 'skipped', 'Element not found'))
            continue
        if not isinstance(element, Dimension):
            skipped.append(_with_status(record, 'skipped', 'Not a Dimension'))
            continue
        try:
            if not element.IsLocked:
                skipped.append(_with_status(record, 'skipped', 'Already unlocked'))
                continue
        except Exception as ex:
            skipped.append(_with_status(record, 'skipped', 'IsLocked read failed: {}'.format(ex)))
            continue
        targets.append((record, element))

    if not targets:
        return {
            'committed': committed,
            'failed': failed,
            'skipped': skipped,
            'transaction_status': 'skipped',
        }

    transaction = Transaction(doc, 'Mirror Project: Disable Constraints')
    started = False
    try:
        transaction.Start()
        started = True
        for record, element in targets:
            try:
                element.IsLocked = False
                if element.IsLocked:
                    failed.append(_with_status(record, 'failed', 'IsLocked remained True'))
                else:
                    committed.append(_with_status(record, 'committed', 'Unlocked'))
            except Exception as ex:
                failed.append(_with_status(record, 'failed', str(ex)))
        status = transaction.Commit()
        started = False
        status_name = str(status)
        if status != TransactionStatus.Committed:
            for record in committed:
                record['status'] = 'failed'
                record['reason'] = 'Transaction not committed: {}'.format(status_name)
                failed.append(record)
            committed = []
        return {
            'committed': committed,
            'failed': failed,
            'skipped': skipped,
            'transaction_status': status_name,
        }
    except Exception as ex:
        if started:
            try:
                transaction.RollBack()
            except Exception:
                pass
        for record, _element in targets:
            failed.append(_with_status(record, 'failed', 'Transaction error: {}'.format(ex)))
        return {
            'committed': [],
            'failed': failed,
            'skipped': skipped,
            'transaction_status': 'RolledBack',
        }


def _with_status(record, status, reason):
    updated = dict(record)
    updated['status'] = status
    updated['reason'] = reason
    return updated
