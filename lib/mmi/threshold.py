# -*- coding: utf-8 -*-
"""Pure decisions for MMI monitor warning thresholds.

No Revit or pyRevit imports, so these helpers can be tested outside Revit.
"""


# Extensible storage int32 fields cannot store a larger value.
INT32_MAX = 2147483647


def parse_limit(value):
    """Return an integer limit of 1 or more, or None when the value is not valid."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, float):
        if int(value) != value:
            return None
        number = int(value)
    else:
        text = str(value).strip()
        if not text:
            return None
        try:
            number = int(text)
        except (TypeError, ValueError):
            return None
    if number < 1 or number > INT32_MAX:
        return None
    return number


def normalize_limit(value, default):
    """Return a valid limit, or the default when value is missing or invalid."""
    parsed = parse_limit(value)
    if parsed is None:
        return int(default)
    return parsed


def at_or_above(value, limit):
    """True when value is greater than or equal to limit."""
    try:
        return int(value) >= int(limit)
    except (TypeError, ValueError):
        return False


def should_warn_on_count(enabled, count, limit):
    """True when the warning is on and count is strictly over the limit."""
    if not enabled:
        return False
    try:
        return int(count) > int(limit)
    except (TypeError, ValueError):
        return False


def include_in_instance_param_count(type_changed, location_changed, regenerated_by_type_edit):
    """True when a modified instance looks like an instance-parameter edit.

    Type-selector changes, moves, and instances that only regenerated because
    their type was edited in the same change are left out of the count.
    """
    if type_changed or location_changed or regenerated_by_type_edit:
        return False
    return True
