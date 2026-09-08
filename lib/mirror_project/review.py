# -*- coding: utf-8 -*-
"""Build portable element targets for source-model review views."""

from __future__ import print_function

from .schema import (
    REVIEW_SCHEMA_VERSION,
    SEVERITY_BLOCKING,
    SEVERITY_REVIEW,
)


_ELEMENT_CHECKS = (
    'constraints',
    'doors',
    'groups',
    'hosted_families',
    'links',
    'mep_connectors',
    'rooms',
    'spaces',
    'walls',
    'windows',
)

_RELATIONSHIP_CATEGORIES = (
    'attachment_base',
    'attachment_top',
    'from_to_room',
    'residual_lock',
    'unrecorded',
)

_BUCKET_ORDER = {
    'Blocking': 0,
    'Missing': 1,
    'Geometry': 2,
    'Relationships': 3,
    'Review': 4,
}


def _bucket_for(item):
    if item.get('severity') == SEVERITY_BLOCKING:
        return 'Blocking'
    after = (item.get('after') or '').lower()
    action = (item.get('action') or '').lower()
    if after == 'missing' or ' missing ' in ' {} '.format(action) or action.startswith('lost '):
        return 'Missing'
    if item.get('category') in ('location', 'facing'):
        return 'Geometry'
    if item.get('category') in _RELATIONSHIP_CATEGORIES:
        return 'Relationships'
    if item.get('check') in ('constraints', 'groups', 'mep_connectors'):
        return 'Relationships'
    return 'Review'


def _higher_priority(left, right):
    return left if _BUCKET_ORDER[left] <= _BUCKET_ORDER[right] else right


def build_review_package(before, after, comparison, exported_utc=None):
    """Return compact JSON-safe review targets from un-compacted findings."""
    findings = comparison.get('detail_findings') or comparison.get('findings') or []
    targets = {}
    non_element_findings = []

    for item in findings:
        severity = item.get('severity')
        if severity not in (SEVERITY_BLOCKING, SEVERITY_REVIEW):
            continue
        uid = item.get('element') or ''
        check = item.get('check') or ''
        if not uid or check not in _ELEMENT_CHECKS:
            non_element_findings.append({
                'severity': severity,
                'check': check,
                'category': item.get('category') or '',
                'before': item.get('before') or '',
                'after': item.get('after') or '',
                'action': item.get('action') or '',
            })
            continue

        bucket = _bucket_for(item)
        target = targets.get(uid)
        if target is None:
            target = {
                'unique_id': uid,
                'bucket': bucket,
                'severity': severity,
                'checks': [],
                'categories': [],
                'reasons': [],
            }
            targets[uid] = target
        else:
            target['bucket'] = _higher_priority(target['bucket'], bucket)
            if severity == SEVERITY_BLOCKING:
                target['severity'] = SEVERITY_BLOCKING

        if check not in target['checks']:
            target['checks'].append(check)
        category = item.get('category') or ''
        if category and category not in target['categories']:
            target['categories'].append(category)
        reason = item.get('action') or ''
        if reason and reason not in target['reasons']:
            target['reasons'].append(reason)

    ordered = sorted(
        targets.values(),
        key=lambda item: (_BUCKET_ORDER[item['bucket']], item['unique_id']),
    )
    return {
        'schema_version': REVIEW_SCHEMA_VERSION,
        'exported_utc': exported_utc,
        'baseline_fingerprint': before.get('fingerprint') or {},
        'compared_fingerprint': after.get('fingerprint') or {},
        'transform': comparison.get('transform') or {},
        'summary': comparison.get('summary') or {},
        'targets': ordered,
        'non_element_findings': non_element_findings,
    }
