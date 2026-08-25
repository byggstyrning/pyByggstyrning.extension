# -*- coding: utf-8 -*-
"""Tolerance-aware comparison of Mirror Project snapshots.

Operates on plain dicts so tests can run outside Revit. UniqueId is the
primary match key (Mirror Project transforms in place). Geometric locations
are compared after applying an inferred named transform so expected
reflection/rotation is not reported as movement.
"""

from __future__ import print_function

from .geometry import (
    angle_deg,
    apply_named_point,
    apply_named_vector,
    as_xyz,
    infer_best_transform,
    within_tol,
)
from .schema import (
    DEFAULT_ANGLE_TOL_DEG,
    DEFAULT_POSITION_TOL_MM,
    SCHEMA_VERSION,
    SEVERITY_BLOCKING,
    SEVERITY_INFO,
    SEVERITY_REVIEW,
    STATUS_OK,
)


def _metric_data(snapshot, name):
    metrics = snapshot.get('metrics') or {}
    result = metrics.get(name) or {}
    if result.get('status') != STATUS_OK:
        return None
    return result.get('data')


def _finding(severity, check, category, element, before, after, action):
    return {
        'severity': severity,
        'check': check,
        'category': category or '',
        'element': element or '',
        'before': _as_text(before),
        'after': _as_text(after),
        'action': action or '',
        'occurrences': 1,
        'examples': element or '',
    }


def _as_text(value):
    if value is None:
        return ''
    if isinstance(value, (list, tuple, dict)):
        try:
            import json
            return json.dumps(value, sort_keys=True, ensure_ascii=True)
        except Exception:
            return str(value)
    return str(value)


def _index_by_uid(items):
    indexed = {}
    for item in items or []:
        uid = item.get('unique_id')
        if uid:
            indexed[uid] = item
    return indexed


def _count_map(data):
    if not data:
        return {}
    return dict(data.get('by_category') or {})


def _sample_pairs(before, after):
    pairs = []
    before_samples = (_metric_data(before, 'samples') or {}).get('points') or []
    after_map = {}
    after_samples = (_metric_data(after, 'samples') or {}).get('points') or []
    for item in after_samples:
        sid = item.get('id')
        if sid:
            after_map[sid] = as_xyz(item.get('point_mm'))
    for item in before_samples:
        sid = item.get('id')
        if sid in ('pbp', 'survey'):
            continue
        b = as_xyz(item.get('point_mm'))
        a = after_map.get(sid)
        if sid and b is not None and a is not None:
            pairs.append((b, a))
    if pairs:
        return pairs
    # Fallback: door/wall location UniqueIds.
    for metric in ('doors', 'walls'):
        b_data = _metric_data(before, metric) or {}
        a_data = _metric_data(after, metric) or {}
        b_items = _index_by_uid(b_data.get('items') or [])
        a_items = _index_by_uid(a_data.get('items') or [])
        for uid, b_item in b_items.items():
            a_item = a_items.get(uid)
            if not a_item:
                continue
            bp = as_xyz(b_item.get('location_mm') or b_item.get('start_mm'))
            ap = as_xyz(a_item.get('location_mm') or a_item.get('start_mm'))
            if bp is not None and ap is not None:
                pairs.append((bp, ap))
            if len(pairs) >= 24:
                return pairs
    return pairs


def _guide(snapshot):
    return snapshot.get('guide') or {}


def _tol(before, after):
    guide = _guide(after) or _guide(before)
    tols = guide.get('tolerances') or {}
    pos = tols.get('position_mm', DEFAULT_POSITION_TOL_MM)
    ang = tols.get('angle_deg', DEFAULT_ANGLE_TOL_DEG)
    try:
        pos = float(pos)
    except (TypeError, ValueError):
        pos = DEFAULT_POSITION_TOL_MM
    try:
        ang = float(ang)
    except (TypeError, ValueError):
        ang = DEFAULT_ANGLE_TOL_DEG
    return pos, ang


def compare_snapshots(before, after, manifest=None):
    """Return ``{summary, transform, findings}`` for two snapshots."""
    findings = []
    pos_tol, ang_tol = _tol(before, after)

    findings.extend(_schema_findings(before, after))
    findings.extend(_fingerprint_findings(before, after))

    pairs = _sample_pairs(before, after)
    transform = infer_best_transform(pairs, position_tol_mm=pos_tol)
    if transform.get('confidence') == 'ambiguous':
        findings.append(_finding(
            SEVERITY_REVIEW, 'transform', 'coordinates', '',
            transform.get('candidates'), transform.get('name'),
            'Best-fit native transform is ambiguous; review geometry manually.',
        ))
    elif transform.get('confidence') == 'low':
        findings.append(_finding(
            SEVERITY_REVIEW, 'transform', 'coordinates', '',
            None, transform.get('rms_mm'),
            'Could not match a named transform; geometry diffs were suppressed to avoid noise.',
        ))
    elif transform.get('confidence') == 'not_available':
        findings.append(_finding(
            SEVERITY_INFO, 'transform', 'coordinates', '',
            None, None,
            'Not enough stable points to infer a named transform.',
        ))
    elif transform.get('name') and transform.get('name') != 'identity':
        findings.append(_finding(
            SEVERITY_INFO, 'transform', 'coordinates', '',
            'identity', transform.get('name'),
            'Comparing locations after applying inferred transform {} (rms {} mm).'.format(
                transform.get('name'), transform.get('rms_mm')),
        ))

    xform_name = (
        transform.get('name')
        if transform.get('confidence') in ('high', 'ambiguous')
        else None
    )

    findings.extend(_count_findings(before, after))
    findings.extend(_warning_findings(before, after))
    findings.extend(_workset_findings(before, after))
    findings.extend(_room_findings(before, after, 'rooms'))
    findings.extend(_room_findings(before, after, 'spaces'))
    findings.extend(_coordinate_findings(before, after, xform_name, pos_tol, ang_tol))
    findings.extend(_link_findings(before, after, xform_name, pos_tol))
    findings.extend(_wall_findings(before, after, xform_name, pos_tol))
    findings.extend(_constraint_findings(before, after, manifest))
    findings.extend(_group_findings(before, after))
    findings.extend(_family_findings(before, after, xform_name, pos_tol, ang_tol, 'doors'))
    findings.extend(_family_findings(before, after, xform_name, pos_tol, ang_tol, 'windows'))
    findings.extend(_family_findings(before, after, xform_name, pos_tol, ang_tol, 'hosted_families'))
    findings.extend(_mep_findings(before, after))
    findings.extend(_extent_findings(before, after, xform_name, pos_tol))

    detail_findings = findings
    findings = _compact_findings(detail_findings)
    summary = _summarize(findings, transform)
    summary['detail_total'] = len(detail_findings)
    return {
        'summary': summary,
        'transform': transform,
        'findings': findings,
        'detail_findings': detail_findings,
    }


def _schema_findings(before, after):
    findings = []
    bver = before.get('schema_version')
    aver = after.get('schema_version')
    if bver != SCHEMA_VERSION or aver != SCHEMA_VERSION:
        findings.append(_finding(
            SEVERITY_BLOCKING, 'schema', '', '',
            bver, aver,
            'Snapshot schema version mismatch; results may be incomplete.',
        ))
    return findings


def _fingerprint_findings(before, after):
    findings = []
    bf = before.get('fingerprint') or {}
    af = after.get('fingerprint') or {}
    if bf.get('revit_build') and af.get('revit_build') and bf.get('revit_build') != af.get('revit_build'):
        findings.append(_finding(
            SEVERITY_REVIEW, 'identity', '', '',
            bf.get('revit_build'), af.get('revit_build'),
            'Revit build changed between snapshots.',
        ))
    if bf.get('title') and af.get('title') and bf.get('title') != af.get('title'):
        findings.append(_finding(
            SEVERITY_INFO, 'identity', '', '',
            bf.get('title'), af.get('title'),
            'Document title differs (expected for a detached/renamed copy).',
        ))
    return findings


def _count_findings(before, after):
    findings = []
    b_data = _metric_data(before, 'counts') or {}
    a_data = _metric_data(after, 'counts') or {}
    blocking_fields = ('doors', 'windows', 'rooms', 'spaces', 'sheets')
    review_fields = ('instances', 'types', 'views', 'groups', 'imports', 'families', 'walls')
    for field in blocking_fields + review_fields:
        bv = b_data.get(field)
        av = a_data.get(field)
        if bv is None or av is None or bv == av:
            continue
        severity = SEVERITY_BLOCKING if field in blocking_fields else SEVERITY_REVIEW
        findings.append(_finding(
            severity, 'counts', field, '',
            bv, av,
            'Category/count drifted by {}.'.format(av - bv),
        ))
    b_counts = _count_map(b_data)
    a_counts = _count_map(a_data)
    keys = set(b_counts.keys()) | set(a_counts.keys())
    duplicate_categories = {
        'Doors': 'doors',
        'Windows': 'windows',
        'Walls': 'walls',
        'Rooms': 'rooms',
        'Spaces': 'spaces',
        'Sheets': 'sheets',
        'Views': 'views',
    }
    for key in sorted(keys):
        bv = int(b_counts.get(key, 0) or 0)
        av = int(a_counts.get(key, 0) or 0)
        if bv == av:
            continue
        field = duplicate_categories.get(key)
        if field and b_data.get(field) == bv and a_data.get(field) == av:
            continue
        findings.append(_finding(
            SEVERITY_REVIEW, 'counts', key, '',
            bv, av,
            'Category/count drifted by {}.'.format(av - bv),
        ))
    return findings


def _warning_groups(data):
    groups = {}
    for item in (data or {}).get('items') or []:
        key = item.get('failure_guid') or item.get('description') or ''
        if not key:
            continue
        groups[key] = {
            'count': int(item.get('count') or 0),
            'description': item.get('description') or '',
            'affected': int(item.get('affected_count') or 0),
        }
    return groups


def _warning_findings(before, after):
    findings = []
    b_data = _metric_data(before, 'warnings')
    a_data = _metric_data(after, 'warnings')
    if b_data is None or a_data is None:
        if b_data is None and a_data is None:
            return findings
        findings.append(_finding(
            SEVERITY_REVIEW, 'warnings', '', '',
            None if b_data is None else 'ok',
            None if a_data is None else 'ok',
            'Warning collector unavailable on one snapshot.',
        ))
        return findings
    b_groups = _warning_groups(b_data)
    a_groups = _warning_groups(a_data)
    keys = set(b_groups.keys()) | set(a_groups.keys())
    for key in sorted(keys):
        b_item = b_groups.get(key)
        a_item = a_groups.get(key)
        if not b_item and a_item:
            findings.append(_finding(
                SEVERITY_REVIEW, 'warnings', '', key,
                0, a_item.get('count'),
                'New warning group: {}'.format(a_item.get('description')),
            ))
        elif b_item and not a_item:
            findings.append(_finding(
                SEVERITY_INFO, 'warnings', '', key,
                b_item.get('count'), 0,
                'Warning group cleared: {}'.format(b_item.get('description')),
            ))
        elif b_item and a_item and b_item.get('count') != a_item.get('count'):
            findings.append(_finding(
                SEVERITY_REVIEW, 'warnings', '', key,
                b_item.get('count'), a_item.get('count'),
                'Warning count changed: {}'.format(a_item.get('description')),
            ))
    return findings


def _workset_findings(before, after):
    findings = []
    b_data = _metric_data(before, 'worksets') or {}
    a_data = _metric_data(after, 'worksets') or {}
    b_names = set((b_data.get('names') or []))
    a_names = set((a_data.get('names') or []))
    missing = sorted(b_names - a_names)
    added = sorted(a_names - b_names)
    for name in missing:
        findings.append(_finding(
            SEVERITY_REVIEW, 'worksets', '', name,
            'present', 'missing',
            'Workset missing after operation.',
        ))
    for name in added:
        findings.append(_finding(
            SEVERITY_INFO, 'worksets', '', name,
            'missing', 'present',
            'New workset appeared after operation.',
        ))
    return findings


def _room_findings(before, after, metric):
    findings = []
    b_data = _metric_data(before, metric) or {}
    a_data = _metric_data(after, metric) or {}
    b_states = b_data.get('by_state') or {}
    a_states = a_data.get('by_state') or {}
    for state in set(list(b_states.keys()) + list(a_states.keys())):
        bv = int(b_states.get(state, 0) or 0)
        av = int(a_states.get(state, 0) or 0)
        if bv == av:
            continue
        severity = SEVERITY_BLOCKING if state in ('not_enclosed', 'redundant', 'unplaced') and av > bv else SEVERITY_REVIEW
        findings.append(_finding(
            severity, metric, state, '',
            bv, av,
            '{} {} count changed.'.format(metric, state),
        ))
    b_items = _index_by_uid(b_data.get('items') or [])
    a_items = _index_by_uid(a_data.get('items') or [])
    for uid, b_item in b_items.items():
        a_item = a_items.get(uid)
        if not a_item:
            findings.append(_finding(
                SEVERITY_REVIEW, metric, b_item.get('state'), uid,
                b_item.get('number'), '',
                'Lost {} after operation.'.format(metric[:-1] if metric.endswith('s') else metric),
            ))
            continue
        if b_item.get('state') != a_item.get('state'):
            findings.append(_finding(
                SEVERITY_REVIEW, metric, uid, uid,
                b_item.get('state'), a_item.get('state'),
                'Spatial state changed.',
            ))
    return findings


def _coordinate_findings(before, after, xform_name, pos_tol, ang_tol):
    findings = []
    b_data = _metric_data(before, 'coordinates') or {}
    a_data = _metric_data(after, 'coordinates') or {}
    if not b_data or not a_data:
        return findings
    b_angle = b_data.get('true_north_deg')
    a_angle = a_data.get('true_north_deg')
    if b_angle is not None and a_angle is not None:
        try:
            delta = abs(float(a_angle) - float(b_angle))
            while delta > 180.0:
                delta = abs(delta - 360.0)
            if delta > ang_tol:
                findings.append(_finding(
                    SEVERITY_REVIEW, 'coordinates', 'true_north', '',
                    b_angle, a_angle,
                    'True North angle changed by {} deg.'.format(round(delta, 4)),
                ))
        except (TypeError, ValueError):
            pass
    for key, label in (
        ('project_base_point_mm', 'project_base_point'),
        ('survey_point_mm', 'survey_point'),
    ):
        if not xform_name:
            break
        bp = as_xyz(b_data.get(key))
        ap = as_xyz(a_data.get(key))
        if bp is None or ap is None:
            continue
        predicted = apply_named_point(xform_name, bp)
        if not within_tol(predicted, ap, pos_tol):
            findings.append(_finding(
                SEVERITY_REVIEW, 'coordinates', label, '',
                _as_text(list(bp)), _as_text(list(ap)),
                'Control point moved beyond tolerance after inferred transform.',
            ))
    if b_data.get('survey_clipped') != a_data.get('survey_clipped') and (
            b_data.get('survey_clipped') is not None and a_data.get('survey_clipped') is not None):
        findings.append(_finding(
            SEVERITY_REVIEW, 'coordinates', 'survey_clip', '',
            b_data.get('survey_clipped'), a_data.get('survey_clipped'),
            'Survey Point clip state changed.',
        ))
    return findings


def _link_findings(before, after, xform_name, pos_tol):
    findings = []
    b_items = _index_by_uid((_metric_data(before, 'links') or {}).get('items') or [])
    a_items = _index_by_uid((_metric_data(after, 'links') or {}).get('items') or [])
    for uid, b_item in b_items.items():
        a_item = a_items.get(uid)
        if not a_item:
            findings.append(_finding(
                SEVERITY_REVIEW, 'links', b_item.get('name'), uid,
                'loaded' if b_item.get('loaded') else 'unloaded',
                'missing',
                'Link instance missing after operation.',
            ))
            continue
        if b_item.get('loaded') and not a_item.get('loaded'):
            findings.append(_finding(
                SEVERITY_REVIEW, 'links', b_item.get('name'), uid,
                'loaded', 'unloaded',
                'Link became unloaded.',
            ))
        b_origin = as_xyz((b_item.get('transform') or {}).get('origin_mm'))
        a_origin = as_xyz((a_item.get('transform') or {}).get('origin_mm'))
        if xform_name and b_origin is not None and a_origin is not None:
            predicted = apply_named_point(xform_name, b_origin)
            if not within_tol(predicted, a_origin, pos_tol):
                findings.append(_finding(
                    SEVERITY_REVIEW, 'links', b_item.get('name'), uid,
                    list(b_origin), list(a_origin),
                    'Link origin moved beyond tolerance.',
                ))
        if b_item.get('pinned') != a_item.get('pinned'):
            findings.append(_finding(
                SEVERITY_INFO, 'links', b_item.get('name'), uid,
                b_item.get('pinned'), a_item.get('pinned'),
                'Link pin state changed.',
            ))
    return findings


def _wall_findings(before, after, xform_name, pos_tol):
    findings = []
    b_items = _index_by_uid((_metric_data(before, 'walls') or {}).get('items') or [])
    a_items = _index_by_uid((_metric_data(after, 'walls') or {}).get('items') or [])
    for uid, b_item in b_items.items():
        a_item = a_items.get(uid)
        if not a_item:
            findings.append(_finding(
                SEVERITY_REVIEW, 'walls', '', uid,
                'present', 'missing',
                'Wall missing after operation.',
            ))
            continue
        for loc in ('top', 'base'):
            b_ids = set(b_item.get('attach_{}'.format(loc)) or [])
            a_ids = set(a_item.get('attach_{}'.format(loc)) or [])
            if b_ids != a_ids:
                findings.append(_finding(
                    SEVERITY_REVIEW, 'walls', 'attachment_{}'.format(loc), uid,
                    sorted(b_ids), sorted(a_ids),
                    'Wall {} attachment hosts changed.'.format(loc),
                ))
        b_start = as_xyz(b_item.get('start_mm'))
        b_end = as_xyz(b_item.get('end_mm'))
        a_start = as_xyz(a_item.get('start_mm'))
        a_end = as_xyz(a_item.get('end_mm'))
        if xform_name and b_start is not None and a_start is not None:
            predicted_start = apply_named_point(xform_name, b_start)
            direct = within_tol(predicted_start, a_start, pos_tol)
            swapped = False
            predicted_end = None
            if b_end is not None and a_end is not None:
                predicted_end = apply_named_point(xform_name, b_end)
                direct = direct and within_tol(predicted_end, a_end, pos_tol)
                swapped = (
                    within_tol(predicted_start, a_end, pos_tol)
                    and within_tol(predicted_end, a_start, pos_tol)
                )
            if not direct and not swapped:
                findings.append(_finding(
                    SEVERITY_REVIEW, 'walls', 'location', uid,
                    [list(b_start), list(b_end) if b_end is not None else None],
                    [list(a_start), list(a_end) if a_end is not None else None],
                    'Wall endpoints moved beyond tolerance after inferred transform.',
                ))
    return findings


def _expected_unlocked(manifest):
    expected = set()
    if not manifest:
        return expected
    results = manifest.get('unlock_results') or {}
    for item in results.get('committed') or []:
        uid = item.get('unique_id')
        if uid:
            expected.add(uid)
    classified = manifest.get('classified') or {}
    for item in classified.get('preexisting_unlocked_since_baseline') or []:
        uid = item.get('unique_id')
        if uid:
            expected.add(uid)
    return expected


def _constraint_findings(before, after, manifest):
    findings = []
    expected = _expected_unlocked(manifest)
    b_items = _index_by_uid((_metric_data(before, 'constraints') or {}).get('items') or [])
    a_items = _index_by_uid((_metric_data(after, 'constraints') or {}).get('items') or [])
    for uid, b_item in b_items.items():
        a_item = a_items.get(uid)
        was_locked = bool(b_item.get('is_locked'))
        now_locked = bool(a_item.get('is_locked')) if a_item else False
        if uid in expected:
            if a_item and now_locked:
                findings.append(_finding(
                    SEVERITY_BLOCKING, 'constraints', 'residual_lock', uid,
                    True, True,
                    'Constraint was unlocked in the manifest but is still locked.',
                ))
            continue
        if was_locked and a_item and not now_locked:
            findings.append(_finding(
                SEVERITY_REVIEW, 'constraints', 'unrecorded', uid,
                True, False,
                'Locked constraint changed without a mutation manifest record.',
            ))
        if was_locked and not a_item:
            findings.append(_finding(
                SEVERITY_REVIEW, 'constraints', b_item.get('classification'), uid,
                True, 'missing',
                'Constraint element missing after operation.',
            ))
        if was_locked and a_item and now_locked and b_item.get('classification') != 'eligible':
            findings.append(_finding(
                SEVERITY_INFO, 'constraints', b_item.get('classification'), uid,
                True, True,
                'Residual lock remains (excluded from auto-disable).',
            ))
    residual_after = 0
    a_data = _metric_data(after, 'constraints') or {}
    residual_after = int(a_data.get('locked_count') or 0)
    if residual_after:
        findings.append(_finding(
            SEVERITY_REVIEW, 'constraints', 'residual', '',
            (_metric_data(before, 'constraints') or {}).get('locked_count'),
            residual_after,
            'Locked constraints remain after preparation/mirror.',
        ))
    return findings


def _group_findings(before, after):
    findings = []
    b_data = _metric_data(before, 'groups') or {}
    a_data = _metric_data(after, 'groups') or {}
    if b_data.get('instance_count') != a_data.get('instance_count') and (
            b_data.get('instance_count') is not None and a_data.get('instance_count') is not None):
        findings.append(_finding(
            SEVERITY_REVIEW, 'groups', 'instances', '',
            b_data.get('instance_count'), a_data.get('instance_count'),
            'Group instance count changed.',
        ))
    b_items = _index_by_uid(b_data.get('items') or [])
    a_items = _index_by_uid(a_data.get('items') or [])
    for uid, b_item in b_items.items():
        a_item = a_items.get(uid)
        if not a_item:
            findings.append(_finding(
                SEVERITY_REVIEW, 'groups', b_item.get('type_name'), uid,
                'present', 'missing',
                'Group instance missing after operation.',
            ))
            continue
        if int(b_item.get('member_count') or 0) != int(a_item.get('member_count') or 0):
            findings.append(_finding(
                SEVERITY_REVIEW, 'groups', b_item.get('type_name'), uid,
                b_item.get('member_count'), a_item.get('member_count'),
                'Group member count changed (possible lost members).',
            ))
        if bool(b_item.get('nested')) != bool(a_item.get('nested')):
            findings.append(_finding(
                SEVERITY_REVIEW, 'groups', b_item.get('type_name'), uid,
                b_item.get('nested'), a_item.get('nested'),
                'Nested-group state changed.',
            ))
    return findings


def _family_findings(before, after, xform_name, pos_tol, ang_tol, metric):
    findings = []
    is_mirror = bool(xform_name and xform_name.startswith('mirror_'))
    b_items = _index_by_uid((_metric_data(before, metric) or {}).get('items') or [])
    a_items = _index_by_uid((_metric_data(after, metric) or {}).get('items') or [])
    for uid, b_item in b_items.items():
        a_item = a_items.get(uid)
        if not a_item:
            if metric == 'hosted_families':
                findings.append(_finding(
                    SEVERITY_INFO, metric,
                    'cohort_exit:{}'.format(b_item.get('category') or 'unknown'), uid,
                    'risk cohort', 'not in risk cohort',
                    'Element left the post-mirror risk-family cohort; this alone '
                    'does not prove the Revit element was deleted.',
                ))
                continue
            findings.append(_finding(
                SEVERITY_REVIEW, metric, b_item.get('category'), uid,
                'present', 'missing',
                'Family instance missing after operation.',
            ))
            continue
        for flag in ('mirrored', 'hand_flipped', 'facing_flipped', 'workplane_flipped'):
            if b_item.get(flag) != a_item.get(flag):
                findings.append(_finding(
                    SEVERITY_INFO if is_mirror else SEVERITY_REVIEW,
                    metric, flag, uid,
                    b_item.get(flag), a_item.get(flag),
                    ('Expected mirror orientation state change; sample-check design intent.'
                     if is_mirror else
                     'Handedness/orientation flag changed; verify design intent.'),
                ))
        bp = as_xyz(b_item.get('location_mm'))
        ap = as_xyz(a_item.get('location_mm'))
        if xform_name and bp is not None and ap is not None:
            predicted = apply_named_point(xform_name, bp)
            if not within_tol(predicted, ap, pos_tol):
                findings.append(_finding(
                    SEVERITY_REVIEW, metric, 'location', uid,
                    list(bp), list(ap),
                    'Instance location moved beyond tolerance.',
                ))
        bv = as_xyz(b_item.get('facing'))
        av = as_xyz(a_item.get('facing'))
        if xform_name and bv is not None and av is not None:
            predicted_v = apply_named_vector(xform_name, bv)
            reversed_v = (-predicted_v[0], -predicted_v[1], -predicted_v[2])
            flip_changed = (
                b_item.get('mirrored') != a_item.get('mirrored')
                or b_item.get('facing_flipped') != a_item.get('facing_flipped')
            )
            reversed_expected = (
                is_mirror and flip_changed
                and angle_deg(reversed_v, av) <= ang_tol
            )
            if angle_deg(predicted_v, av) > ang_tol and not reversed_expected:
                findings.append(_finding(
                    SEVERITY_REVIEW, metric, 'facing', uid,
                    list(bv), list(av),
                    'Facing vector does not match inferred transform.',
                ))
        if metric in ('doors', 'windows'):
            b_from = b_item.get('from_room') or ''
            b_to = b_item.get('to_room') or ''
            a_from = a_item.get('from_room') or ''
            a_to = a_item.get('to_room') or ''
            if b_from != a_from or b_to != a_to:
                exact_swap = b_from == a_to and b_to == a_from
                findings.append(_finding(
                    SEVERITY_INFO if is_mirror and exact_swap else SEVERITY_REVIEW,
                    metric,
                    'from_to_room_swapped' if is_mirror and exact_swap else 'from_to_room',
                    uid,
                    '{} / {}'.format(b_from, b_to),
                    '{} / {}'.format(a_from, a_to),
                    ('From/To room order swapped but adjacency pair was preserved.'
                     if is_mirror and exact_swap else
                     'From/To room changed; verify egress/handing.'),
                ))
    return findings


def _compact_findings(findings):
    """Aggregate repetitive audit signals while retaining raw details separately."""
    grouped = {}
    passthrough = []
    orientation_flags = (
        'mirrored', 'hand_flipped', 'facing_flipped', 'workplane_flipped',
    )
    for item in findings:
        check = item.get('check')
        category = item.get('category')
        should_group = (
            (check in ('doors', 'windows', 'hosted_families')
             and category in orientation_flags + (
                 'from_to_room', 'from_to_room_swapped'))
            or (check == 'hosted_families'
                and category.startswith('cohort_exit:'))
            or (check == 'constraints'
                and category in ('unrecorded', 'grouped', 'unmodifiable'))
        )
        if not should_group:
            passthrough.append(item)
            continue
        key = (
            item.get('severity'),
            check,
            category,
            item.get('action'),
        )
        grouped.setdefault(key, []).append(item)

    compact = list(passthrough)
    for key in sorted(grouped.keys()):
        rows = grouped[key]
        first = rows[0]
        examples = [row.get('element') for row in rows if row.get('element')][:10]
        compact_item = dict(first)
        compact_item['element'] = ''
        compact_item['before'] = 'varied' if len(rows) > 1 else first.get('before')
        compact_item['after'] = 'varied' if len(rows) > 1 else first.get('after')
        compact_item['occurrences'] = len(rows)
        compact_item['examples'] = '; '.join(examples)
        compact_item['action'] = '{} occurrences. {}'.format(
            len(rows), first.get('action') or '')
        compact.append(compact_item)
    return compact


def _mep_findings(before, after):
    findings = []
    b_data = _metric_data(before, 'mep_connectors') or {}
    a_data = _metric_data(after, 'mep_connectors') or {}
    b_unused = int(b_data.get('unused_count') or 0)
    a_unused = int(a_data.get('unused_count') or 0)
    if a_unused > b_unused:
        findings.append(_finding(
            SEVERITY_BLOCKING, 'mep_connectors', '', '',
            b_unused, a_unused,
            'Disconnected/unused MEP connectors increased.',
        ))
    elif a_unused != b_unused:
        findings.append(_finding(
            SEVERITY_REVIEW, 'mep_connectors', '', '',
            b_unused, a_unused,
            'Unused MEP connector count changed.',
        ))
    b_items = _index_by_uid(b_data.get('items') or [])
    a_items = _index_by_uid(a_data.get('items') or [])
    for uid, a_item in a_items.items():
        b_item = b_items.get(uid)
        b_n = int((b_item or {}).get('unused') or 0)
        a_n = int(a_item.get('unused') or 0)
        if a_n > b_n:
            findings.append(_finding(
                SEVERITY_BLOCKING, 'mep_connectors', a_item.get('category'), uid,
                b_n, a_n,
                'Element gained unused MEP connectors.',
            ))
    return findings


def _extent_findings(before, after, xform_name, pos_tol):
    findings = []
    if not xform_name:
        return findings
    b_data = _metric_data(before, 'extents') or {}
    a_data = _metric_data(after, 'extents') or {}
    b_min = as_xyz(b_data.get('min_mm'))
    b_max = as_xyz(b_data.get('max_mm'))
    a_min = as_xyz(a_data.get('min_mm'))
    a_max = as_xyz(a_data.get('max_mm'))
    if None in (b_min, b_max, a_min, a_max):
        return findings
    corners = [
        (b_min[0], b_min[1], b_min[2]),
        (b_min[0], b_min[1], b_max[2]),
        (b_min[0], b_max[1], b_min[2]),
        (b_min[0], b_max[1], b_max[2]),
        (b_max[0], b_min[1], b_min[2]),
        (b_max[0], b_min[1], b_max[2]),
        (b_max[0], b_max[1], b_min[2]),
        (b_max[0], b_max[1], b_max[2]),
    ]
    predicted = [apply_named_point(xform_name, c) for c in corners]
    pmin = (
        min(p[0] for p in predicted),
        min(p[1] for p in predicted),
        min(p[2] for p in predicted),
    )
    pmax = (
        max(p[0] for p in predicted),
        max(p[1] for p in predicted),
        max(p[2] for p in predicted),
    )
    if not within_tol(pmin, a_min, pos_tol * 5.0) or not within_tol(pmax, a_max, pos_tol * 5.0):
        findings.append(_finding(
            SEVERITY_REVIEW, 'extents', '', '',
            [list(b_min), list(b_max)], [list(a_min), list(a_max)],
            'Overall model extents diverged after inferred transform.',
        ))
    b_levels = b_data.get('level_count')
    a_levels = a_data.get('level_count')
    if b_levels is not None and a_levels is not None and b_levels != a_levels:
        findings.append(_finding(
            SEVERITY_REVIEW, 'extents', 'levels', '',
            b_levels, a_levels,
            'Level count changed.',
        ))
    return findings


def _summarize(findings, transform):
    counts = {
        SEVERITY_BLOCKING: 0,
        SEVERITY_REVIEW: 0,
        SEVERITY_INFO: 0,
    }
    occurrence_counts = dict(counts)
    for item in findings:
        sev = item.get('severity')
        if sev in counts:
            counts[sev] += 1
            occurrence_counts[sev] += int(item.get('occurrences') or 1)
    return {
        'blocking': counts[SEVERITY_BLOCKING],
        'review': counts[SEVERITY_REVIEW],
        'informational': counts[SEVERITY_INFO],
        'blocking_occurrences': occurrence_counts[SEVERITY_BLOCKING],
        'review_occurrences': occurrence_counts[SEVERITY_REVIEW],
        'informational_occurrences': occurrence_counts[SEVERITY_INFO],
        'total': len(findings),
        'transform_name': transform.get('name'),
        'transform_confidence': transform.get('confidence'),
        'passed': counts[SEVERITY_BLOCKING] == 0,
    }
