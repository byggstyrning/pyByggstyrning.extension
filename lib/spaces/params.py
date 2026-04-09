# -*- coding: utf-8 -*-
"""Capture and restore MEP Space parameters across recreation.

Uses linked Room UniqueId (from Space.Room) as the stable key between
old and new spaces.
"""

from __future__ import print_function

import clr
clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BuiltInParameter,
    StorageType,
)
from Autodesk.Revit.DB.Mechanical import Space

from pyrevit import script

logger = script.get_logger()


def _build_skip_builtin_set():
    """Built-in parameter ids to never copy (identity, level, phase, name/number)."""
    skip_params = set()
    skip_param_names = [
        'INVALID',
        'ELEM_TYPE_PARAM',
        'ELEM_CATEGORY_PARAM',
        'ELEM_FAMILY_PARAM',
        'ELEM_FAMILY_AND_TYPE_PARAM',
        'ELEM_TYPE_NAME_PARAM',
        'ELEM_TYPE_ID_PARAM',
        'ELEM_ID_PARAM',
        'ELEM_LEVEL_PARAM',
        'ELEM_LEVEL_ID_PARAM',
        'ELEM_PHASE_CREATED_PARAM',
        'ELEM_PHASE_DEMOLISHED_PARAM',
        'ROOM_NUMBER',
        'ROOM_NAME',
    ]
    for param_name in skip_param_names:
        try:
            if hasattr(BuiltInParameter, param_name):
                skip_params.add(getattr(BuiltInParameter, param_name))
        except Exception:
            pass
    return skip_params


_SKIP_BUILTINS = None


def _get_skip_builtins():
    global _SKIP_BUILTINS
    if _SKIP_BUILTINS is None:
        _SKIP_BUILTINS = _build_skip_builtin_set()
    return _SKIP_BUILTINS


def _should_skip_parameter(source_param):
    """Return True if this parameter should not be captured."""
    if not source_param:
        return True
    if source_param.IsReadOnly:
        return True
    if not source_param.HasValue:
        return True
    try:
        if source_param.StorageType == StorageType.ElementId:
            return True
    except Exception:
        return True
    defn = source_param.Definition
    if not defn:
        return True
    if hasattr(defn, 'BuiltInParameter'):
        try:
            bip = defn.BuiltInParameter
            invalid_bip = getattr(BuiltInParameter, 'INVALID', None)
            skip_set = _get_skip_builtins()
            if invalid_bip is not None and bip != invalid_bip and bip in skip_set:
                return True
            if invalid_bip is None and bip in skip_set:
                return True
        except Exception:
            pass
    return False


def _serialize_value(source_param):
    """Return (storage_type, value_tuple) for snapshot, or None if unsupported."""
    st = source_param.StorageType
    try:
        if st == StorageType.String:
            return (st, source_param.AsString() or '')
        if st == StorageType.Integer:
            return (st, source_param.AsInteger())
        if st == StorageType.Double:
            return (st, source_param.AsDouble())
    except Exception:
        return None
    return None


def capture_space_parameters(spaces):
    """Capture writable parameter values from spaces, keyed by linked Room UniqueId.

    Uses Space.Room to resolve the associated Room. Spaces without a Room link
    are not captured; their ElementIds are returned for exclusion from deletion.

    Args:
        spaces: Iterable of Space elements in the host document.

    Returns:
        tuple: (param_cache, unlinked_space_ids)
            param_cache: dict mapping room UniqueId (str) -> {param_name: (StorageType, value)}
            unlinked_space_ids: set of int (space Id.IntegerValue) for spaces with no Space.Room
    """
    param_cache = {}
    unlinked_space_ids = set()

    for space in spaces:
        if not isinstance(space, Space):
            continue
        try:
            linked_room = space.Room
        except Exception as ex:
            logger.debug("Space.Room failed for {}: {}".format(space.Id, str(ex)))
            linked_room = None

        if linked_room is None:
            unlinked_space_ids.add(space.Id.IntegerValue)
            continue

        try:
            room_uid = linked_room.UniqueId
        except Exception:
            unlinked_space_ids.add(space.Id.IntegerValue)
            continue

        snapshot = {}
        for source_param in space.Parameters:
            try:
                if _should_skip_parameter(source_param):
                    continue
                defn = source_param.Definition
                if not defn:
                    continue
                param_name = defn.Name
                if not param_name:
                    continue
                packed = _serialize_value(source_param)
                if packed is None:
                    continue
                snapshot[param_name] = packed
            except Exception as ex:
                logger.debug("capture param skip: {}".format(str(ex)))

        if room_uid in param_cache:
            logger.debug(
                "Duplicate Room UniqueId in capture; overwriting snapshot for {}".format(
                    room_uid))
        param_cache[room_uid] = snapshot

    return (param_cache, unlinked_space_ids)


def restore_space_parameters(space, param_snapshot):
    """Apply a captured parameter snapshot to a new Space.

    Args:
        space: New Space element in the host document.
        param_snapshot: dict param_name -> (StorageType, value)

    Returns:
        int: Number of parameter values successfully written.
    """
    if not param_snapshot:
        return 0

    restored = 0
    for param_name, packed in param_snapshot.items():
        try:
            st, value = packed
            target_param = space.LookupParameter(param_name)
            if not target_param or target_param.IsReadOnly:
                continue
            if target_param.StorageType != st:
                continue

            if st == StorageType.String:
                if target_param.HasValue:
                    cur = target_param.AsString() or ''
                    if cur == (value or ''):
                        continue
                target_param.Set(value or '')
                restored += 1
            elif st == StorageType.Integer:
                if target_param.HasValue and target_param.AsInteger() == value:
                    continue
                target_param.Set(value)
                restored += 1
            elif st == StorageType.Double:
                if target_param.HasValue:
                    cur = target_param.AsDouble()
                    if abs(cur - value) < 1e-9:
                        continue
                target_param.Set(value)
                restored += 1
        except Exception as ex:
            logger.debug("restore param '{}' failed: {}".format(param_name, str(ex)))

    return restored
