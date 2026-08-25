# -*- coding: utf-8 -*-
"""Versioned JSON schema constants for Mirror Project snapshots."""

from __future__ import print_function

SCHEMA_ID = 'mirror-project'
SCHEMA_VERSION = 'mirror-project/1.0'
MANIFEST_SCHEMA_VERSION = 'mirror-project-manifest/1.0'
POINTER_SCHEMA_VERSION = 'mirror-project-pointer/1.0'
REVIEW_SCHEMA_VERSION = 'mirror-project-review/1.0'

UNITS_MM = 'mm'
COORDINATE_FRAME_INTERNAL = 'internal'

STATUS_OK = 'ok'
STATUS_ERROR = 'error'
STATUS_NOT_AVAILABLE = 'not_available'

SEVERITY_BLOCKING = 'blocking'
SEVERITY_REVIEW = 'review'
SEVERITY_INFO = 'informational'

DEFAULT_POSITION_TOL_MM = 2.0
DEFAULT_ANGLE_TOL_DEG = 0.5

NAMED_TRANSFORMS = (
    'identity',
    'mirror_x',
    'mirror_y',
    'rotate_z_90',
    'rotate_z_180',
    'rotate_z_270',
)

CONSTRAINT_CLASSES = (
    'eligible',
    'grouped',
    'multi_segment',
    'labeled',
    'unsupported_type',
    'missing',
    'unmodifiable',
)

ROOM_STATES = (
    'placed',
    'unplaced',
    'not_enclosed',
    'redundant',
)
