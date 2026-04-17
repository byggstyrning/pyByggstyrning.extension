# -*- coding: utf-8 -*-
"""Revit version compatibility helpers.

Central location for APIs that changed across Revit releases. Import from here
instead of branching on ``HOST_APP.version`` in every call site. All helpers
resolve the right implementation once at module load so the hot path is a
direct function call.

Covered deprecations:

* ``ElementId.IntegerValue`` -> ``ElementId.Value`` (Revit 2024, 64-bit ids).
  ``IntegerValue`` still exists but raises ``OverflowException`` when the id
  does not fit in Int32. Use :func:`get_element_id_value`.

* ``Definition.ParameterType`` / ``ParameterType`` enum (deprecated Revit 2022,
  ``Definition.ParameterType`` removed from the ``Definition`` class in Revit
  2026). Use :func:`get_param_data_type` or the typed predicates
  :func:`is_param_yesno`, :func:`is_param_text`, :func:`is_param_number`,
  :func:`is_param_integer`, :func:`is_param_acceptable_for_mapping`.

* ``DisplayUnitType`` (removed Revit 2025). Use
  :func:`convert_from_internal_units` / :func:`convert_to_internal_units` with
  a ``UnitTypeId`` ``ForgeTypeId``.

* ``ParameterFilterRuleFactory.CreateEqualsRule`` signature changed in Revit
  2024 (dropped ``caseSensitive`` argument). Use :func:`create_equals_rule`.

* ``ElementId(Int32)`` constructor removed in Revit 2026; only the Int64
  overload remains. IronPython usually widens Python ``int`` automatically
  but passing ``System.Int64`` explicitly makes overload resolution
  deterministic on every runtime. Use :func:`make_element_id`.
"""

from pyrevit import HOST_APP

try:
    REVIT_VERSION = int(HOST_APP.version)
except Exception:
    REVIT_VERSION = 2024

HAS_64BIT_ELEMENT_ID = REVIT_VERSION >= 2024
HAS_FORGE_PARAM_TYPE = REVIT_VERSION >= 2022
HAS_LEGACY_PARAM_TYPE = REVIT_VERSION < 2026
HAS_FORGE_UNITS = REVIT_VERSION >= 2021
HAS_CASE_SENSITIVE_EQUALS_RULE = REVIT_VERSION < 2023

from Autodesk.Revit.DB import ElementId, StorageType

try:
    from System import Int64 as _Int64
except ImportError:
    _Int64 = None


if HAS_64BIT_ELEMENT_ID and _Int64 is not None:
    def make_element_id(value):
        """Construct an :class:`ElementId` from an integer, version-safely.

        Revit 2026 removed the ``ElementId(Int32)`` constructor; only
        ``ElementId(Int64)`` remains. IronPython normally widens a Python
        ``int`` to Int64 for single-overload resolution, but passing
        ``System.Int64`` explicitly avoids any risk of the wrong overload
        being picked on pre-2026 builds and matches the runtime signature
        on 2026+.
        """
        if isinstance(value, ElementId):
            return value
        return ElementId(_Int64(value))
else:
    def make_element_id(value):
        if isinstance(value, ElementId):
            return value
        return ElementId(value)


if HAS_64BIT_ELEMENT_ID:
    def get_element_id_value(element_id):
        """Return the integer value of an :class:`ElementId`.

        Uses ``ElementId.Value`` on Revit 2024+ (Int64) and ``IntegerValue``
        on earlier versions. Always prefer this helper over ``.IntegerValue``
        which throws on Int64 overflow starting Revit 2024.
        """
        return element_id.Value
else:
    def get_element_id_value(element_id):
        return element_id.IntegerValue


def get_elementid_value_func():
    """Back-compat shim: return :func:`get_element_id_value`.

    Older code does ``_get = get_elementid_value_func(); _get(eid)``; keep
    that pattern working by returning the plain function directly.
    """
    return get_element_id_value


if HAS_FORGE_PARAM_TYPE:
    try:
        from Autodesk.Revit.DB import SpecTypeId

        _SPEC_YESNO = SpecTypeId.Boolean.YesNo
        _SPEC_TEXT = SpecTypeId.String.Text
        _SPEC_NUMBER = SpecTypeId.Number
        _SPEC_INT = SpecTypeId.Int.Integer
    except Exception:
        # Older 2022 drops may not yet expose all SpecTypeId members; fall
        # back to None and use StorageType only.
        _SPEC_YESNO = None
        _SPEC_TEXT = None
        _SPEC_NUMBER = None
        _SPEC_INT = None
else:
    _SPEC_YESNO = _SPEC_TEXT = _SPEC_NUMBER = _SPEC_INT = None


def get_param_data_type(definition):
    """Return a ``ForgeTypeId`` (2022+) or ``ParameterType`` (<=2021).

    Callers should compare the returned value against
    ``SpecTypeId.*`` / ``ParameterType.*`` as appropriate for their runtime.
    Most code should prefer the typed predicates in this module instead.
    """
    if HAS_FORGE_PARAM_TYPE:
        try:
            return definition.GetDataType()
        except Exception:
            pass
    if HAS_LEGACY_PARAM_TYPE:
        try:
            return definition.ParameterType
        except Exception:
            pass
    return None


def is_param_yesno(definition):
    """True if the parameter definition represents a YesNo/Boolean value."""
    if HAS_FORGE_PARAM_TYPE and _SPEC_YESNO is not None:
        try:
            return definition.GetDataType() == _SPEC_YESNO
        except Exception:
            pass
    if HAS_LEGACY_PARAM_TYPE:
        try:
            from Autodesk.Revit.DB import ParameterType
            return definition.ParameterType == ParameterType.YesNo
        except Exception:
            return False
    return False


def is_param_text(definition):
    """True if the parameter definition is a free-form text field."""
    try:
        if definition.StorageType != StorageType.String:
            return False
    except Exception:
        pass
    if HAS_FORGE_PARAM_TYPE and _SPEC_TEXT is not None:
        try:
            return definition.GetDataType() == _SPEC_TEXT
        except Exception:
            pass
    if HAS_LEGACY_PARAM_TYPE:
        try:
            from Autodesk.Revit.DB import ParameterType
            return definition.ParameterType == ParameterType.Text
        except Exception:
            pass
    # Fall back to storage-type check if richer classification unavailable.
    try:
        return definition.StorageType == StorageType.String
    except Exception:
        return False


def is_param_number(definition):
    """True if the parameter stores a generic Number/Double."""
    try:
        return definition.StorageType == StorageType.Double
    except Exception:
        return False


def is_param_integer(definition):
    """True if the parameter stores an Integer (including YesNo)."""
    try:
        return definition.StorageType == StorageType.Integer
    except Exception:
        return False


def is_param_acceptable_for_mapping(definition):
    """True if a parameter is a plain value type suitable for mapping.

    Matches the historical ``ParameterType in {Text, Number, Integer, YesNo}``
    intent using ``StorageType`` so it works on every Revit version.
    """
    try:
        return definition.StorageType in (
            StorageType.String,
            StorageType.Double,
            StorageType.Integer,
        )
    except Exception:
        return False


if HAS_FORGE_UNITS:
    try:
        from Autodesk.Revit.DB import UnitTypeId, UnitUtils
    except Exception:
        UnitTypeId = None
        UnitUtils = None
else:
    UnitTypeId = None
    try:
        from Autodesk.Revit.DB import UnitUtils
    except Exception:
        UnitUtils = None


def convert_from_internal_units(value, unit_type_id, legacy_display_unit=None):
    """Version-safe ``UnitUtils.ConvertFromInternalUnits`` wrapper.

    ``unit_type_id`` is a ``ForgeTypeId`` (``UnitTypeId.*``). On Revit <2021
    callers must pass the equivalent ``DisplayUnitType`` via
    ``legacy_display_unit``; otherwise the raw value is returned.
    """
    if HAS_FORGE_UNITS and UnitUtils is not None:
        return UnitUtils.ConvertFromInternalUnits(value, unit_type_id)
    if legacy_display_unit is not None and UnitUtils is not None:
        return UnitUtils.ConvertFromInternalUnits(value, legacy_display_unit)
    return value


def convert_to_internal_units(value, unit_type_id, legacy_display_unit=None):
    """Version-safe ``UnitUtils.ConvertToInternalUnits`` wrapper."""
    if HAS_FORGE_UNITS and UnitUtils is not None:
        return UnitUtils.ConvertToInternalUnits(value, unit_type_id)
    if legacy_display_unit is not None and UnitUtils is not None:
        return UnitUtils.ConvertToInternalUnits(value, legacy_display_unit)
    return value


def create_equals_rule(param_id, value, case_sensitive=True):
    """Version-safe ``ParameterFilterRuleFactory.CreateEqualsRule``.

    Revit 2023 and earlier require a ``caseSensitive`` third argument; the
    overload was removed in Revit 2024. This wrapper hides that difference.
    """
    from Autodesk.Revit.DB import ParameterFilterRuleFactory
    if HAS_CASE_SENSITIVE_EQUALS_RULE:
        return ParameterFilterRuleFactory.CreateEqualsRule(
            param_id, value, case_sensitive
        )
    return ParameterFilterRuleFactory.CreateEqualsRule(param_id, value)
