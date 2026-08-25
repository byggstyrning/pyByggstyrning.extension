# -*- coding: utf-8 -*-
"""Assemble a versioned Mirror Project snapshot from collectors."""

from __future__ import print_function

import datetime

from .collectors import (
    collect_areas,
    collect_constraints,
    collect_coordinates,
    collect_identity,
    collect_links,
    collect_model_scan,
    collect_rooms,
    collect_spaces,
    collect_warnings,
    collect_worksets,
)
from .geometry import round_xyz
from .result import run_collector
from .schema import (
    COORDINATE_FRAME_INTERNAL,
    DEFAULT_ANGLE_TOL_DEG,
    DEFAULT_POSITION_TOL_MM,
    SCHEMA_VERSION,
    UNITS_MM,
)


def utc_now():
    return datetime.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')


def capture_snapshot(doc):
    """Run collectors and return a complete snapshot dict."""
    metrics = {}
    for name, fn in (
        ('identity', collect_identity),
        ('worksets', collect_worksets),
        ('warnings', collect_warnings),
        ('coordinates', collect_coordinates),
        ('links', collect_links),
        ('rooms', collect_rooms),
        ('spaces', collect_spaces),
        ('areas', collect_areas),
        ('constraints', collect_constraints),
    ):
        key, result = run_collector(name, fn, doc)
        metrics[key] = result

    key, scan_result = run_collector('model_scan', collect_model_scan, doc)
    if scan_result.get('status') == 'ok' and isinstance(scan_result.get('data'), dict):
        scan = scan_result['data']
        for nested in (
            'counts', 'walls', 'doors', 'windows', 'hosted_families',
            'groups', 'mep_connectors', 'extents', 'samples',
        ):
            metrics[nested] = {
                'status': 'ok',
                'data': scan.get(nested),
                'timing_ms': scan_result.get('timing_ms'),
                'errors': [],
            }
    else:
        for nested in (
            'counts', 'walls', 'doors', 'windows', 'hosted_families',
            'groups', 'mep_connectors', 'extents', 'samples',
        ):
            metrics[nested] = {
                'status': scan_result.get('status'),
                'data': None,
                'timing_ms': scan_result.get('timing_ms'),
                'errors': scan_result.get('errors') or [],
            }

    _add_control_samples(metrics)

    identity = (metrics.get('identity') or {}).get('data') or {}
    snapshot = {
        'schema_version': SCHEMA_VERSION,
        'captured_utc': utc_now(),
        'fingerprint': identity,
        'guide': {
            'units': UNITS_MM,
            'coordinate_frame': COORDINATE_FRAME_INTERNAL,
            'tolerances': {
                'position_mm': DEFAULT_POSITION_TOL_MM,
                'angle_deg': DEFAULT_ANGLE_TOL_DEG,
            },
            'scope': 'whole_document',
        },
        'metrics': metrics,
    }
    return snapshot


def _add_control_samples(metrics):
    samples_result = metrics.get('samples') or {}
    data = samples_result.get('data') or {}
    points = list(data.get('points') or [])
    coords = (metrics.get('coordinates') or {}).get('data') or {}
    pbp = coords.get('project_base_point_mm')
    survey = coords.get('survey_point_mm')
    extra = []
    if isinstance(pbp, (list, tuple)):
        extra.append({'id': 'pbp', 'point_mm': round_xyz(pbp)})
    if isinstance(survey, (list, tuple)):
        extra.append({'id': 'survey', 'point_mm': round_xyz(survey)})
    metrics['samples'] = {
        'status': samples_result.get('status') or 'ok',
        'data': {'points': extra + points},
        'timing_ms': samples_result.get('timing_ms') or 0.0,
        'errors': samples_result.get('errors') or [],
    }


def preflight_summary(snapshot):
    """Compact dict for the preflight Markdown report."""
    metrics = snapshot.get('metrics') or {}

    def data(name):
        result = metrics.get(name) or {}
        return result.get('status'), result.get('data') or {}, result.get('errors') or []

    _status, counts, _err = data('counts')
    _status, warnings, _err = data('warnings')
    _status, rooms, _err = data('rooms')
    _status, spaces, _err = data('spaces')
    _status, walls, _err = data('walls')
    _status, constraints, _err = data('constraints')
    _status, groups, _err = data('groups')
    _status, mep, _err = data('mep_connectors')
    _status, links, _err = data('links')
    _status, coords, _err = data('coordinates')
    identity = snapshot.get('fingerprint') or {}
    failed = []
    for name, result in metrics.items():
        if result.get('status') == 'error':
            failed.append(name)
    return {
        'title': identity.get('title'),
        'revit_build': identity.get('revit_build'),
        'workshared': identity.get('is_workshared'),
        'detached': identity.get('is_detached'),
        'cloud': identity.get('is_cloud'),
        'instances': counts.get('instances'),
        'doors': counts.get('doors'),
        'windows': counts.get('windows'),
        'walls': counts.get('walls'),
        'views': counts.get('views'),
        'sheets': counts.get('sheets'),
        'warning_total': warnings.get('total'),
        'warning_groups': warnings.get('group_count'),
        'rooms': rooms.get('count'),
        'room_states': rooms.get('by_state'),
        'spaces': spaces.get('count'),
        'space_states': spaces.get('by_state'),
        'attached_walls': walls.get('attached_count'),
        'locked_constraints': constraints.get('locked_count'),
        'constraint_classes': constraints.get('by_class'),
        'groups': groups.get('instance_count'),
        'nested_groups': groups.get('nested_count'),
        'mep_unused': mep.get('unused_count'),
        'links': links.get('count'),
        'true_north_deg': coords.get('true_north_deg'),
        'failed_collectors': failed,
    }
