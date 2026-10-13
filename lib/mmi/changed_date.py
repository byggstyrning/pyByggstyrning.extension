# -*- coding: utf-8 -*-
"""When an MMI button should stamp a changed date. No Revit imports."""

DATE_FORMAT = "%Y-%m-%d"


def format_changed_date(when):
    """Return a calendar date as YYYY-MM-dd."""
    return when.strftime(DATE_FORMAT)


def _clean_name(value):
    if value is None:
        return ""
    return str(value).strip()


def should_stamp_changed_date(enabled, parameter_name, mmi_parameter_name=None):
    """True when a button click should write today's date.

    The switch must be on, a text parameter must be chosen, and that parameter
    must not be the MMI parameter itself.
    """
    if not enabled:
        return False
    name = _clean_name(parameter_name)
    if not name:
        return False
    mmi_name = _clean_name(mmi_parameter_name)
    if mmi_name and name == mmi_name:
        return False
    return True
