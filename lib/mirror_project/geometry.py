# -*- coding: utf-8 -*-
"""Pure-Python geometry helpers for mirror/rotation inference and compare.

No Revit imports. Points and vectors are 3-tuples of floats in millimetres
unless noted. Named transforms match native Mirror Project axes through the
Internal Origin and 90-degree Project North rotations about Z.
"""

from __future__ import print_function

import math

from .schema import (
    DEFAULT_ANGLE_TOL_DEG,
    DEFAULT_POSITION_TOL_MM,
    NAMED_TRANSFORMS,
)

FEET_TO_MM = 304.8
DEG_TO_RAD = math.pi / 180.0
RAD_TO_DEG = 180.0 / math.pi


def as_xyz(value):
    """Coerce list/tuple/.NET XYZ-like to a 3-tuple of floats."""
    if value is None:
        return None
    if isinstance(value, (list, tuple)):
        if len(value) < 3:
            return None
        return (float(value[0]), float(value[1]), float(value[2]))
    try:
        return (float(value.X), float(value.Y), float(value.Z))
    except Exception:
        return None


def xyz_mm_from_feet(value):
    xyz = as_xyz(value)
    if xyz is None:
        return None
    return (xyz[0] * FEET_TO_MM, xyz[1] * FEET_TO_MM, xyz[2] * FEET_TO_MM)


def round_xyz(xyz, digits=6):
    xyz = as_xyz(xyz)
    if xyz is None:
        return None
    return [round(xyz[0], digits), round(xyz[1], digits), round(xyz[2], digits)]


def dot(a, b):
    a = as_xyz(a)
    b = as_xyz(b)
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def sub(a, b):
    a = as_xyz(a)
    b = as_xyz(b)
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def add(a, b):
    a = as_xyz(a)
    b = as_xyz(b)
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def scale(v, s):
    v = as_xyz(v)
    return (v[0] * s, v[1] * s, v[2] * s)


def length(v):
    return math.sqrt(dot(v, v))


def distance(a, b):
    return length(sub(a, b))


def normalize(v):
    v = as_xyz(v)
    mag = length(v)
    if mag < 1e-12:
        return (0.0, 0.0, 0.0)
    return scale(v, 1.0 / mag)


def angle_deg(a, b):
    """Unsigned angle in degrees between two vectors."""
    a = normalize(a)
    b = normalize(b)
    mag = length(a) * length(b)
    if mag < 1e-12:
        return 0.0
    cosang = max(-1.0, min(1.0, dot(a, b)))
    return math.acos(cosang) * RAD_TO_DEG


def reflect_point(point, origin, normal):
    """Reflect point about a plane: p' = p - 2n((p-o)·n)."""
    n = normalize(normal)
    d = dot(sub(point, origin), n)
    return sub(point, scale(n, 2.0 * d))


def reflect_vector(vec, normal):
    """Reflect a direction: v' = v - 2n(v·n)."""
    n = normalize(normal)
    return sub(vec, scale(n, 2.0 * dot(vec, n)))


def rotate_point_z(point, origin, angle_rad):
    p = sub(point, origin)
    c = math.cos(angle_rad)
    s = math.sin(angle_rad)
    return add(origin, (p[0] * c - p[1] * s, p[0] * s + p[1] * c, p[2]))


def rotate_vector_z(vec, angle_rad):
    c = math.cos(angle_rad)
    s = math.sin(angle_rad)
    v = as_xyz(vec)
    return (v[0] * c - v[1] * s, v[0] * s + v[1] * c, v[2])


def bbox_corners(min_xyz, max_xyz, local_origin=(0.0, 0.0, 0.0),
                 basis_x=(1.0, 0.0, 0.0), basis_y=(0.0, 1.0, 0.0),
                 basis_z=(0.0, 0.0, 1.0)):
    """Return eight corners of an oriented bounding box."""
    mn = as_xyz(min_xyz)
    mx = as_xyz(max_xyz)
    o = as_xyz(local_origin)
    bx = as_xyz(basis_x)
    by = as_xyz(basis_y)
    bz = as_xyz(basis_z)
    corners = []
    for x in (mn[0], mx[0]):
        for y in (mn[1], mx[1]):
            for z in (mn[2], mx[2]):
                p = add(o, add(scale(bx, x), add(scale(by, y), scale(bz, z))))
                corners.append(p)
    return corners


def aabb_from_corners(corners):
    xs = [c[0] for c in corners]
    ys = [c[1] for c in corners]
    zs = [c[2] for c in corners]
    return {
        'min_mm': [min(xs), min(ys), min(zs)],
        'max_mm': [max(xs), max(ys), max(zs)],
        'corners_mm': [round_xyz(c) for c in corners],
    }


def apply_named_point(name, point, origin=(0.0, 0.0, 0.0)):
    point = as_xyz(point)
    origin = as_xyz(origin)
    if name == 'identity':
        return point
    if name == 'mirror_x':
        return reflect_point(point, origin, (1.0, 0.0, 0.0))
    if name == 'mirror_y':
        return reflect_point(point, origin, (0.0, 1.0, 0.0))
    if name == 'rotate_z_90':
        return rotate_point_z(point, origin, math.pi / 2.0)
    if name == 'rotate_z_180':
        return rotate_point_z(point, origin, math.pi)
    if name == 'rotate_z_270':
        return rotate_point_z(point, origin, 3.0 * math.pi / 2.0)
    raise ValueError('Unknown transform {}'.format(name))


def apply_named_vector(name, vec):
    vec = as_xyz(vec)
    if name == 'identity':
        return vec
    if name == 'mirror_x':
        return reflect_vector(vec, (1.0, 0.0, 0.0))
    if name == 'mirror_y':
        return reflect_vector(vec, (0.0, 1.0, 0.0))
    if name == 'rotate_z_90':
        return rotate_vector_z(vec, math.pi / 2.0)
    if name == 'rotate_z_180':
        return rotate_vector_z(vec, math.pi)
    if name == 'rotate_z_270':
        return rotate_vector_z(vec, 3.0 * math.pi / 2.0)
    raise ValueError('Unknown transform {}'.format(name))


def rms_mm(pairs, name, origin=(0.0, 0.0, 0.0)):
    """RMS distance in mm after applying named transform to before-points."""
    if not pairs:
        return None
    acc = 0.0
    n = 0
    for before, after in pairs:
        predicted = apply_named_point(name, before, origin=origin)
        d = distance(predicted, after)
        acc += d * d
        n += 1
    if n == 0:
        return None
    return math.sqrt(acc / float(n))


def _transform_errors(pairs, name, origin=(0.0, 0.0, 0.0)):
    errors = []
    for before, after in pairs:
        predicted = apply_named_point(name, before, origin=origin)
        errors.append(distance(predicted, after))
    return errors


def _median(values):
    ordered = sorted(values)
    count = len(ordered)
    if not count:
        return None
    middle = count // 2
    if count % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2.0


def infer_best_transform(pairs, position_tol_mm=DEFAULT_POSITION_TOL_MM,
                         origin=(0.0, 0.0, 0.0)):
    """Pick named transform using inlier consensus, then inlier RMS.

    Returns confidence ``high``, ``low``, ``ambiguous``, or ``not_available``.
    Coordinate controls can remain fixed during native Mirror Project, so a
    small number of outliers must not overwhelm correctly mirrored model points.
    """
    usable = []
    for before, after in pairs or []:
        b = as_xyz(before)
        a = as_xyz(after)
        if b is None or a is None:
            continue
        usable.append((b, a))
    if len(usable) < 1:
        return {
            'name': None,
            'rms_mm': None,
            'all_rms_mm': None,
            'median_mm': None,
            'confidence': 'not_available',
            'n_samples': 0,
            'n_inliers': 0,
            'inlier_ratio': 0.0,
            'candidates': [],
        }

    inlier_tol = max(float(position_tol_mm) * 5.0, 10.0)
    ranked = []
    for name in NAMED_TRANSFORMS:
        errors = _transform_errors(usable, name, origin=origin)
        inliers = [error for error in errors if error <= inlier_tol]
        inlier_rms = None
        if inliers:
            inlier_rms = math.sqrt(
                sum(error * error for error in inliers) / float(len(inliers)))
        all_rms = math.sqrt(
            sum(error * error for error in errors) / float(len(errors)))
        ranked.append({
            'name': name,
            'rms_mm': None if inlier_rms is None else round(inlier_rms, 4),
            'all_rms_mm': round(all_rms, 4),
            'median_mm': round(_median(errors), 4),
            'n_inliers': len(inliers),
            'inlier_ratio': round(len(inliers) / float(len(errors)), 4),
        })
    ranked.sort(key=lambda item: (
        -item['n_inliers'],
        item['rms_mm'] if item['rms_mm'] is not None else 1e12,
        item['median_mm'],
    ))
    best = ranked[0]
    second = ranked[1] if len(ranked) > 1 else None
    confidence = 'high'
    minimum_inliers = min(len(usable), 3)
    minimum_ratio = 0.6 if len(usable) >= 5 else 0.5
    if best['rms_mm'] is None or best['n_inliers'] < minimum_inliers:
        confidence = 'not_available'
    elif best['inlier_ratio'] < minimum_ratio or best['rms_mm'] > inlier_tol:
        confidence = 'low'
    elif second is not None and second['rms_mm'] is not None:
        if (second['n_inliers'] == best['n_inliers']
                and abs(second['rms_mm'] - best['rms_mm']) <= position_tol_mm):
            confidence = 'ambiguous'
    return {
        'name': best['name'],
        'rms_mm': best['rms_mm'],
        'all_rms_mm': best['all_rms_mm'],
        'median_mm': best['median_mm'],
        'confidence': confidence,
        'n_samples': len(usable),
        'n_inliers': best['n_inliers'],
        'inlier_ratio': best['inlier_ratio'],
        'candidates': ranked,
    }


def within_tol(a, b, tol=DEFAULT_POSITION_TOL_MM):
    if a is None or b is None:
        return False
    return distance(a, b) <= tol


def angle_within_tol(a, b, tol_deg=DEFAULT_ANGLE_TOL_DEG):
    if a is None or b is None:
        return False
    return angle_deg(a, b) <= tol_deg
