# -*- coding: utf-8 -*-
"""Text formulas for 3D View Reference labels.

No Revit imports, so the rules can be tested outside Revit.
A formula is read left to right. Each part is one value, and the text
between two parts is stored on the part before the gap.
"""

try:
    unicode
except NameError:
    unicode = str


SOURCE_SHEET_NUMBER = "sheet_number"
SOURCE_SHEET = "sheet"
SOURCE_PROJECT = "project"
SOURCE_VIEW_NAME = "view_name"
SOURCE_VIEW = "view"

SHEET_NUMBER_LABEL = "Sheet number"
VIEW_NAME_LABEL = "View name"

DEFAULT_SHEET_FORMULA = {
    "parts": [{"source": SOURCE_SHEET_NUMBER}],
}
DEFAULT_VIEW_FORMULA = {
    "parts": [{"source": SOURCE_VIEW_NAME}],
}

_NAMED_SOURCES = (SOURCE_SHEET, SOURCE_PROJECT, SOURCE_VIEW)
_BARE_SOURCES = (SOURCE_SHEET_NUMBER, SOURCE_VIEW_NAME)


def _as_text(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except Exception:
        return u"{}".format(value)


def _formula_part(part):
    """One formula part, or None. Keeps a per-gap separator when the part has one."""
    if not isinstance(part, dict):
        return None
    source = _as_text(part.get("source"))
    if source in _BARE_SOURCES:
        item = {"source": source}
    elif source in _NAMED_SOURCES and part.get("name"):
        item = {"source": source, "name": _as_text(part.get("name"))}
    else:
        return None
    if "separator" in part:
        item["separator"] = _as_text(part.get("separator"))
    return item


def normalize_formula(formula, default_source):
    """Parts in order. The text between two parts is stored on the part before the gap.

    An older formula had one separator for every gap. That value is copied onto
    each gap. An empty or broken formula is a single default part.
    """
    parts = []
    shared = None
    own = False
    if isinstance(formula, dict):
        if formula.get("separator"):
            shared = _as_text(formula.get("separator"))
        for part in formula.get("parts") or []:
            item = _formula_part(part)
            if item is None:
                continue
            if "separator" in item:
                own = True
            parts.append(item)
    if not parts:
        return {"parts": [{"source": default_source}]}
    if not own and shared:
        for part in parts[:-1]:
            part["separator"] = shared
    for part in parts[:-1]:
        part["separator"] = _as_text(part.get("separator"))
    parts[-1].pop("separator", None)
    return {"parts": parts}


def normalize_sheet_number_formula(formula):
    return normalize_formula(formula, SOURCE_SHEET_NUMBER)


def normalize_view_name_formula(formula):
    return normalize_formula(formula, SOURCE_VIEW_NAME)


def part_label(part):
    """What a formula part is called in the dropdown."""
    source = part.get("source")
    if source == SOURCE_PROJECT:
        return u"Project: {}".format(part.get("name") or "")
    if source == SOURCE_SHEET:
        return u"Sheet: {}".format(part.get("name") or "")
    if source == SOURCE_VIEW:
        return u"View: {}".format(part.get("name") or "")
    if source == SOURCE_VIEW_NAME:
        return VIEW_NAME_LABEL
    return SHEET_NUMBER_LABEL


def compose_parts(formula, default_source, text_of):
    """Join the parts whose text is not empty. text_of(part) returns that text."""
    formula = normalize_formula(formula, default_source)
    values = []
    for part in formula["parts"]:
        text = _as_text(text_of(part)).strip()
        if text:
            values.append((text, _as_text(part.get("separator"))))
    if not values:
        return u""
    written = [values[0][0]]
    for index in range(len(values) - 1):
        written.append(values[index][1])
        written.append(values[index + 1][0])
    return u"".join(written)
