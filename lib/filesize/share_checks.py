# -*- coding: utf-8 -*-
"""Pure checks for the optional nested-family sharing experiment.

No Revit imports. The button uses these to decide when a reload must be
rolled back, and to report how many elements each category gained or lost.
"""


def reload_loss(before_types, after_types, before_placed, after_placed):
    """Return (missing type names, placed count dropped).

    A reload is kept only when every previous type name is still there and
    the placed count did not fall. Extra types and extra placements are fine.
    """
    missing = sorted(set(before_types).difference(after_types))
    placed_lost = after_placed < before_placed
    return missing, placed_lost


def category_delta(before, after):
    """Map category name to count change, including categories that disappear."""
    names = set(before.keys())
    names.update(after.keys())
    delta = {}
    for name in names:
        change = after.get(name, 0) - before.get(name, 0)
        if change:
            delta[name] = change
    return delta
