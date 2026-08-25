# -*- coding: utf-8 -*-
"""Pure-Python tests for Mirror Project snapshot/diff helpers."""

from __future__ import print_function

import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
if LIB not in sys.path:
    sys.path.insert(0, LIB)

from mirror_project import schema  # noqa: E402
from mirror_project.diff import compare_snapshots  # noqa: E402
from mirror_project.geometry import (  # noqa: E402
    angle_deg,
    apply_named_point,
    apply_named_vector,
    infer_best_transform,
    reflect_point,
    rotate_point_z,
    within_tol,
)
from mirror_project.reporting import export_findings_csv  # noqa: E402
from mirror_project.review import build_review_package  # noqa: E402
from mirror_project.storage import load_json, save_json  # noqa: E402


def _metric(data):
    return {
        'status': 'ok',
        'data': data,
        'timing_ms': 1.0,
        'errors': [],
    }


def _snapshot(metrics, title='A', extra=None):
    snap = {
        'schema_version': schema.SCHEMA_VERSION,
        'captured_utc': '2026-08-24T08:00:00Z',
        'fingerprint': {
            'title': title,
            'revit_build': '26.0.0.0',
        },
        'guide': {
            'units': 'mm',
            'coordinate_frame': 'internal',
            'tolerances': {
                'position_mm': 2.0,
                'angle_deg': 0.5,
            },
        },
        'metrics': metrics,
    }
    if extra:
        snap.update(extra)
    return snap


class SchemaStorageTests(unittest.TestCase):
    def test_round_trip(self):
        payload = {
            'schema_version': schema.SCHEMA_VERSION,
            'nested': {'count': 3, 'ok': True, 'missing': None},
            'items': ['door', 1, 2.5],
        }
        handle, path = tempfile.mkstemp(suffix='.json')
        os.close(handle)
        try:
            save_json(path, payload)
            loaded = load_json(path)
            self.assertEqual(loaded['schema_version'], schema.SCHEMA_VERSION)
            self.assertEqual(loaded['nested']['count'], 3)
            self.assertEqual(loaded['items'][0], 'door')
        finally:
            os.remove(path)

    def test_json_safe_cp1252_bytes_and_unicode(self):
        import json
        from mirror_project.storage import json_safe

        converted = json_safe({'title': u'Malm\u00d6'})
        self.assertEqual(converted['title'], u'Malm\u00d6')
        json.dumps(converted, ensure_ascii=True)

        # Py2/IronPython: str is bytes. 0xD6 is cp1252 Ö (Revit Windows strings).
        if isinstance(b'x', str):
            converted = json_safe({'title': 'Malm\xd6'})
            json.dumps(converted, ensure_ascii=True)
            self.assertIn(u'\u00d6', converted['title'])

    def test_swedish_round_trip_via_safe_dumper(self):
        payload = {
            'schema_version': schema.SCHEMA_VERSION,
            'fingerprint': {'title': u'Malm\u00f6', 'path_name': u'C:\\Projekt\\Örebro'},
            'items': [u'vägg', u'rum 101'],
        }
        handle, path = tempfile.mkstemp(suffix='.json')
        os.close(handle)
        try:
            save_json(path, payload)
            loaded = load_json(path)
            self.assertEqual(loaded['fingerprint']['title'], u'Malm\u00f6')
            self.assertEqual(loaded['items'][1], u'rum 101')
        finally:
            os.remove(path)


class GeometryTests(unittest.TestCase):
    def test_reflect_x(self):
        out = reflect_point((10.0, 5.0, 1.0), (0.0, 0.0, 0.0), (1.0, 0.0, 0.0))
        self.assertTrue(within_tol(out, (-10.0, 5.0, 1.0), 0.001))

    def test_rotate_180(self):
        import math
        out = rotate_point_z((10.0, 4.0, 2.0), (0.0, 0.0, 0.0), math.pi)
        self.assertTrue(within_tol(out, (-10.0, -4.0, 2.0), 0.001))

    def test_infer_mirror_x(self):
        pairs = [
            ((100.0, 20.0, 0.0), (-100.0, 20.0, 0.0)),
            ((40.0, -8.0, 3.0), (-40.0, -8.0, 3.0)),
            ((12.0, 50.0, 9.0), (-12.0, 50.0, 9.0)),
        ]
        result = infer_best_transform(pairs, position_tol_mm=2.0)
        self.assertEqual(result['name'], 'mirror_x')
        self.assertEqual(result['confidence'], 'high')
        self.assertLess(result['rms_mm'], 0.1)

    def test_infer_mirror_x_ignores_fixed_survey_outlier(self):
        pairs = [
            ((100.0, 20.0, 0.0), (-100.0, 20.0, 0.0)),
            ((40.0, -8.0, 3.0), (-40.0, -8.0, 3.0)),
            ((12.0, 50.0, 9.0), (-12.0, 50.0, 9.0)),
            # Native Mirror Project can leave the Survey Point unchanged.
            ((-483059.0, -721594.0, -20250.0),
             (-483059.0, -721594.0, -20250.0)),
        ]
        result = infer_best_transform(pairs, position_tol_mm=2.0)
        self.assertEqual(result['name'], 'mirror_x')
        self.assertEqual(result['confidence'], 'high')
        self.assertEqual(result['n_inliers'], 3)
        self.assertEqual(result['inlier_ratio'], 0.75)

    def test_infer_rotate_z_90(self):
        pairs = [
            ((10.0, 0.0, 0.0), apply_named_point('rotate_z_90', (10.0, 0.0, 0.0))),
            ((0.0, 25.0, 1.0), apply_named_point('rotate_z_90', (0.0, 25.0, 1.0))),
            ((8.0, 4.0, 2.0), apply_named_point('rotate_z_90', (8.0, 4.0, 2.0))),
        ]
        result = infer_best_transform(pairs, position_tol_mm=2.0)
        self.assertEqual(result['name'], 'rotate_z_90')
        self.assertEqual(result['confidence'], 'high')

    def test_ambiguous_origin_only(self):
        pairs = [((0.0, 0.0, 0.0), (0.0, 0.0, 0.0))]
        result = infer_best_transform(pairs, position_tol_mm=2.0)
        self.assertEqual(result['confidence'], 'ambiguous')

    def test_vector_mirror(self):
        vec = apply_named_vector('mirror_x', (1.0, 0.0, 0.0))
        self.assertTrue(within_tol(vec, (-1.0, 0.0, 0.0), 0.001))
        self.assertLess(angle_deg((0.0, 1.0, 0.0), (0.0, 1.0, 0.0)), 0.01)


class ReviewPackageTests(unittest.TestCase):
    def test_missing_element_is_exported_for_source_model(self):
        before = _snapshot({})
        after = _snapshot({}, title='Mirrored')
        comparison = {
            'summary': {'blocking': 1},
            'transform': {'name': 'mirror_x', 'confidence': 'high'},
            'findings': [],
            'detail_findings': [{
                'severity': schema.SEVERITY_REVIEW,
                'check': 'windows',
                'category': 'Windows',
                'element': 'window-uid-1',
                'before': 'present',
                'after': 'missing',
                'action': 'Family instance missing after operation.',
            }, {
                'severity': schema.SEVERITY_BLOCKING,
                'check': 'counts',
                'category': 'windows',
                'element': '',
                'before': '2',
                'after': '1',
                'action': 'Category/count drifted by -1.',
            }],
        }
        package = build_review_package(
            before,
            after,
            comparison,
            exported_utc='2026-08-24T12:00:00Z',
        )
        self.assertEqual(package['schema_version'], schema.REVIEW_SCHEMA_VERSION)
        self.assertEqual(len(package['targets']), 1)
        self.assertEqual(package['targets'][0]['unique_id'], 'window-uid-1')
        self.assertEqual(package['targets'][0]['bucket'], 'Missing')
        self.assertEqual(len(package['non_element_findings']), 1)

    def test_target_uses_highest_priority_bucket(self):
        comparison = {
            'detail_findings': [{
                'severity': schema.SEVERITY_REVIEW,
                'check': 'walls',
                'category': 'location',
                'element': 'wall-uid-1',
                'action': 'Location changed.',
            }, {
                'severity': schema.SEVERITY_BLOCKING,
                'check': 'mep_connectors',
                'category': '',
                'element': 'wall-uid-1',
                'action': 'Element gained unused connectors.',
            }],
        }
        package = build_review_package(_snapshot({}), _snapshot({}), comparison)
        self.assertEqual(len(package['targets']), 1)
        self.assertEqual(package['targets'][0]['bucket'], 'Blocking')
        self.assertEqual(package['targets'][0]['severity'], schema.SEVERITY_BLOCKING)


class DiffTests(unittest.TestCase):
    def _base_metrics(self):
        return {
            'counts': _metric({
                'instances': 10,
                'doors': 2,
                'windows': 2,
                'walls': 4,
                'views': 3,
                'sheets': 1,
                'groups': 1,
                'imports': 0,
                'families': 5,
                'by_category': {'Doors': 2, 'Walls': 4},
            }),
            'warnings': _metric({
                'total': 1,
                'group_count': 1,
                'items': [{
                    'failure_guid': 'guid-a',
                    'description': 'Highlighted walls overlap',
                    'count': 1,
                    'affected_count': 2,
                }],
            }),
            'worksets': _metric({'names': ['Shared Levels and Grids']}),
            'rooms': _metric({
                'count': 1,
                'by_state': {'placed': 1, 'unplaced': 0, 'not_enclosed': 0, 'redundant': 0},
                'items': [{'unique_id': 'room-1', 'number': '101', 'state': 'placed'}],
            }),
            'spaces': _metric({
                'count': 0,
                'by_state': {'placed': 0, 'unplaced': 0, 'not_enclosed': 0, 'redundant': 0},
                'items': [],
            }),
            'coordinates': _metric({
                'project_base_point_mm': [1000.0, 200.0, 0.0],
                'survey_point_mm': [0.0, 0.0, 0.0],
                'true_north_deg': 0.0,
                'survey_clipped': True,
            }),
            'links': _metric({'items': []}),
            'walls': _metric({
                'items': [{
                    'unique_id': 'wall-1',
                    'start_mm': [100.0, 20.0, 0.0],
                    'end_mm': [200.0, 20.0, 0.0],
                    'attach_top': ['floor-1'],
                    'attach_base': [],
                }],
            }),
            'constraints': _metric({
                'locked_count': 1,
                'items': [{
                    'unique_id': 'dim-1',
                    'is_locked': True,
                    'classification': 'eligible',
                }],
            }),
            'groups': _metric({
                'instance_count': 1,
                'items': [{'unique_id': 'g-1', 'type_name': 'Core', 'member_count': 4, 'nested': False}],
            }),
            'doors': _metric({
                'items': [{
                    'unique_id': 'door-1',
                    'category': 'Doors',
                    'location_mm': [40.0, -8.0, 3.0],
                    'facing': [1.0, 0.0, 0.0],
                    'mirrored': False,
                    'hand_flipped': False,
                    'facing_flipped': False,
                    'workplane_flipped': False,
                    'from_room': 'room-a',
                    'to_room': 'room-b',
                }],
            }),
            'windows': _metric({'items': []}),
            'hosted_families': _metric({'items': []}),
            'mep_connectors': _metric({
                'unused_count': 1,
                'items': [{'unique_id': 'pipe-1', 'category': 'Pipes', 'unused': 1}],
            }),
            'extents': _metric({
                'min_mm': [0.0, 0.0, 0.0],
                'max_mm': [200.0, 50.0, 9.0],
                'level_count': 2,
            }),
            'samples': _metric({
                'points': [
                    {'id': 'pbp', 'point_mm': [1000.0, 200.0, 0.0]},
                    {'id': 'wall-1', 'point_mm': [100.0, 20.0, 0.0]},
                    {'id': 'door-1', 'point_mm': [40.0, -8.0, 3.0]},
                ],
            }),
        }

    def _mirrored_metrics(self, source):
        import copy
        after = copy.deepcopy(source)
        after['coordinates']['data']['project_base_point_mm'] = [-1000.0, 200.0, 0.0]
        after['walls']['data']['items'][0]['start_mm'] = [-100.0, 20.0, 0.0]
        after['walls']['data']['items'][0]['end_mm'] = [-200.0, 20.0, 0.0]
        after['doors']['data']['items'][0]['location_mm'] = [-40.0, -8.0, 3.0]
        after['doors']['data']['items'][0]['facing'] = [-1.0, 0.0, 0.0]
        after['doors']['data']['items'][0]['mirrored'] = True
        after['samples']['data']['points'] = [
            {'id': 'pbp', 'point_mm': [-1000.0, 200.0, 0.0]},
            {'id': 'wall-1', 'point_mm': [-100.0, 20.0, 0.0]},
            {'id': 'door-1', 'point_mm': [-40.0, -8.0, 3.0]},
        ]
        after['extents']['data']['min_mm'] = [-200.0, 0.0, 0.0]
        after['extents']['data']['max_mm'] = [0.0, 50.0, 9.0]
        after['constraints']['data']['locked_count'] = 0
        after['constraints']['data']['items'][0]['is_locked'] = False
        return after

    def test_count_and_warning_delta(self):
        before = _snapshot(self._base_metrics())
        after_metrics = self._mirrored_metrics(self._base_metrics())
        after_metrics['counts']['data']['doors'] = 1
        after_metrics['counts']['data']['by_category']['Doors'] = 1
        after_metrics['warnings']['data']['items'].append({
            'failure_guid': 'guid-b',
            'description': 'Room not enclosed',
            'count': 2,
            'affected_count': 1,
        })
        after_metrics['warnings']['data']['total'] = 3
        after_metrics['warnings']['data']['group_count'] = 2
        after = _snapshot(after_metrics, title='A_MIRROR')
        result = compare_snapshots(before, after)
        checks = [item['check'] for item in result['findings']]
        self.assertIn('counts', checks)
        self.assertIn('warnings', checks)
        self.assertTrue(any(item['severity'] == schema.SEVERITY_BLOCKING and item['check'] == 'counts'
                            for item in result['findings']))
        self.assertTrue(any(item['element'] == 'guid-b' for item in result['findings']))

    def test_handedness_and_attachments(self):
        before = _snapshot(self._base_metrics())
        after_metrics = self._mirrored_metrics(self._base_metrics())
        after_metrics['walls']['data']['items'][0]['attach_top'] = []
        after = _snapshot(after_metrics)
        result = compare_snapshots(before, after)
        self.assertTrue(any(item['check'] == 'doors' and item['category'] == 'mirrored'
                            for item in result['findings']))
        self.assertTrue(any(item['check'] == 'walls' and 'attachment' in item['category']
                            for item in result['findings']))

    def test_fixed_survey_outlier_does_not_flood_location_findings(self):
        before_metrics = self._base_metrics()
        before_metrics['samples']['data']['points'].append({
            'id': 'survey',
            'point_mm': [-483059.0, -721594.0, -20250.0],
        })
        after_metrics = self._mirrored_metrics(before_metrics)
        after_metrics['samples']['data']['points'].append({
            'id': 'survey',
            'point_mm': [-483059.0, -721594.0, -20250.0],
        })
        result = compare_snapshots(
            _snapshot(before_metrics),
            _snapshot(after_metrics),
        )
        self.assertEqual(result['transform']['name'], 'mirror_x')
        self.assertEqual(result['transform']['confidence'], 'high')
        self.assertEqual(result['transform']['n_samples'], 2)
        self.assertFalse(any(
            item['check'] == 'walls' and item['category'] == 'location'
            for item in result['findings']))

    def test_wall_endpoint_order_may_reverse(self):
        before_metrics = self._base_metrics()
        after_metrics = self._mirrored_metrics(before_metrics)
        after_metrics['walls']['data']['items'][0]['start_mm'] = [-200.0, 20.0, 0.0]
        after_metrics['walls']['data']['items'][0]['end_mm'] = [-100.0, 20.0, 0.0]
        result = compare_snapshots(
            _snapshot(before_metrics),
            _snapshot(after_metrics),
        )
        self.assertFalse(any(
            item['check'] == 'walls' and item['category'] == 'location'
            for item in result['findings']))

    def test_mirror_facing_reversal_with_flip_is_expected(self):
        before_metrics = self._base_metrics()
        before_metrics['doors']['data']['items'][0]['facing'] = [0.0, 1.0, 0.0]
        after_metrics = self._mirrored_metrics(before_metrics)
        after_metrics['doors']['data']['items'][0]['facing'] = [0.0, -1.0, 0.0]
        result = compare_snapshots(
            _snapshot(before_metrics),
            _snapshot(after_metrics),
        )
        self.assertFalse(any(
            item['check'] == 'doors' and item['category'] == 'facing'
            for item in result['findings']))

    def test_exact_from_to_swap_is_informational(self):
        before_metrics = self._base_metrics()
        after_metrics = self._mirrored_metrics(before_metrics)
        after_metrics['doors']['data']['items'][0]['from_room'] = 'room-b'
        after_metrics['doors']['data']['items'][0]['to_room'] = 'room-a'
        result = compare_snapshots(
            _snapshot(before_metrics),
            _snapshot(after_metrics),
        )
        swapped = [
            item for item in result['findings']
            if item['check'] == 'doors'
            and item['category'] == 'from_to_room_swapped'
        ]
        self.assertEqual(len(swapped), 1)
        self.assertEqual(swapped[0]['severity'], schema.SEVERITY_INFO)

    def test_orientation_changes_are_compacted_with_occurrences(self):
        before_metrics = self._base_metrics()
        second = dict(before_metrics['doors']['data']['items'][0])
        second['unique_id'] = 'door-2'
        second['location_mm'] = [60.0, 5.0, 0.0]
        before_metrics['doors']['data']['items'].append(second)
        after_metrics = self._mirrored_metrics(before_metrics)
        after_metrics['doors']['data']['items'][0]['mirrored'] = True
        after_metrics['doors']['data']['items'][1]['location_mm'] = [-60.0, 5.0, 0.0]
        after_metrics['doors']['data']['items'][1]['facing'] = [-1.0, 0.0, 0.0]
        after_metrics['doors']['data']['items'][1]['mirrored'] = True
        result = compare_snapshots(
            _snapshot(before_metrics),
            _snapshot(after_metrics),
        )
        mirrored = [
            item for item in result['findings']
            if item['check'] == 'doors' and item['category'] == 'mirrored'
        ]
        self.assertEqual(len(mirrored), 1)
        self.assertEqual(mirrored[0]['occurrences'], 2)
        self.assertEqual(mirrored[0]['severity'], schema.SEVERITY_INFO)
        self.assertGreaterEqual(result['summary']['informational_occurrences'], 2)

    def test_mep_disconnect_increase(self):
        before = _snapshot(self._base_metrics())
        after_metrics = self._mirrored_metrics(self._base_metrics())
        after_metrics['mep_connectors']['data']['unused_count'] = 4
        after_metrics['mep_connectors']['data']['items'][0]['unused'] = 4
        after = _snapshot(after_metrics)
        result = compare_snapshots(before, after)
        self.assertTrue(any(item['check'] == 'mep_connectors' and item['severity'] == schema.SEVERITY_BLOCKING
                            for item in result['findings']))
        self.assertFalse(result['summary']['passed'])

    def test_expected_constraint_unlock(self):
        before = _snapshot(self._base_metrics())
        after_metrics = self._mirrored_metrics(self._base_metrics())
        after = _snapshot(after_metrics)
        manifest = {
            'schema_version': schema.MANIFEST_SCHEMA_VERSION,
            'unlock_results': {
                'committed': [{'unique_id': 'dim-1'}],
            },
        }
        result = compare_snapshots(before, after, manifest=manifest)
        unrecorded = [item for item in result['findings']
                      if item['check'] == 'constraints' and item['category'] == 'unrecorded']
        residual_lock = [item for item in result['findings']
                         if item['category'] == 'residual_lock']
        self.assertEqual(unrecorded, [])
        self.assertEqual(residual_lock, [])

    def test_preexisting_unlock_recorded_by_prepare_manifest(self):
        before = _snapshot(self._base_metrics())
        after_metrics = self._mirrored_metrics(self._base_metrics())
        after = _snapshot(after_metrics)
        manifest = {
            'schema_version': schema.MANIFEST_SCHEMA_VERSION,
            'unlock_results': {'committed': []},
            'classified': {
                'preexisting_unlocked_since_baseline': [
                    {'unique_id': 'dim-1'},
                ],
            },
        }
        result = compare_snapshots(before, after, manifest=manifest)
        self.assertFalse(any(
            item['check'] == 'constraints' and item['category'] == 'unrecorded'
            for item in result['findings']))

    def test_unrecorded_constraint_unlock(self):
        before = _snapshot(self._base_metrics())
        after_metrics = self._mirrored_metrics(self._base_metrics())
        after = _snapshot(after_metrics)
        result = compare_snapshots(before, after, manifest=None)
        self.assertTrue(any(item['category'] == 'unrecorded' for item in result['findings']))

    def test_schema_mismatch_blocking(self):
        before = _snapshot(self._base_metrics())
        after = _snapshot(self._mirrored_metrics(self._base_metrics()))
        after['schema_version'] = 'mirror-project/0.1'
        result = compare_snapshots(before, after)
        self.assertTrue(any(item['check'] == 'schema' and item['severity'] == schema.SEVERITY_BLOCKING
                            for item in result['findings']))

    def test_csv_export(self):
        before = _snapshot(self._base_metrics())
        after = _snapshot(self._mirrored_metrics(self._base_metrics()), title='B')
        result = compare_snapshots(before, after)
        handle, path = tempfile.mkstemp(suffix='.csv')
        os.close(handle)
        try:
            export_findings_csv(path, result['findings'])
            with open(path, 'rb') as handle_in:
                text = handle_in.read().decode('utf-8')
            self.assertIn('severity', text)
            self.assertIn('check', text)
        finally:
            os.remove(path)


if __name__ == '__main__':
    unittest.main()
