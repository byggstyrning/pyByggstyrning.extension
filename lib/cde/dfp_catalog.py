# -*- coding: utf-8 -*-
"""DFP (door-function programme) catalog — mirrors Hub ``dfp-catalog.ts``.

Each function maps to ``Pset_DFP.Func_<code>`` on the door or its linked IfcGroup.
"""
from __future__ import division

DFP_PSET_NAME = "Pset_DFP"

# Stable function codes and Swedish display labels (category omitted; group=Function).
RAW_FUNCTIONS = [
    ("1.01", "Nödutrymning mekanisk"),
    ("1.02", "Nödutrymning 2-vägs"),
    ("1.03", "Nödut- och återinrymning"),
    ("1.04", "Nödutrymning, elektrisk tryckknapp"),
    ("1.05", "Branddörrstängning"),
    ("1.07", "Brandcellsgräns"),
    ("1.08", "Skyddsklass"),
    ("1.09", "Gastät"),
    ("1.10", "Ventilation"),
    ("1.11", "Ventilation öppenarea"),
    ("2.1", "Dörrautomatik"),
    ("2.2", "Dörrautomatik, beröringsfri öppning"),
    ("2.3", "Dörrautomatik, brandsäker utrymning"),
    ("2.4", "Freeswing"),
    ("2.6", "Förberedd för dörrautom."),
    ("2.7", "Dörruppställning elektronisk"),
    ("2.8", "Dörruppställning mekanisk"),
    ("2.9", "Dörrstängare vanlig"),
    ("3.1", "Mekaniskt lås"),
    ("3.2", "Lås A"),
    ("3.3", "Lås B"),
    ("4.1", "Kortläsare in"),
    ("4.2", "Kortläsare in och ut"),
    ("4.3", "Dörrbladsläsare"),
    ("4.7", "Intelligent låscylinder"),
    ("5.1", "Öppnaknapp mekanisk"),
    ("5.2", "Öppnaknapp beröringsfri"),
    ("5.3", "Tryckknapp i trycke"),
    ("5.4", "Trycke/Handtag"),
    ("6.1", "Tidsstyrd"),
    ("6.2", "Daglarm"),
    ("6.3", "Slussfunktion"),
    ("6.4", "Förregling"),
    ("6.5", "Bokningssystem"),
    ("7.1", "Entresignal"),
    ("7.2", "Porttelefon"),
    ("7.3", "Manöverpanel"),
    ("7.4", "Dörrind. för ventilation"),
]

_NUMBER_CODES = frozenset(["1.11"])


def dfp_property_name(code):
    """``\"5.4\"`` -> ``\"Func_5_4\"``."""
    return "Func_{}".format(str(code).replace(".", "_"))


def dfp_value_key(code):
    """Full ``Pset_DFP.<propertyName>`` key for a function code."""
    return "{}.{}".format(DFP_PSET_NAME, dfp_property_name(code))


def build_dfp_parameter_defs(param_def_fn):
    """Return schedule ParameterDef rows for all catalogued DFP functions."""
    defs = []
    for code, label in RAW_FUNCTIONS:
        value_type = "number" if code in _NUMBER_CODES else "bool"
        defs.append(param_def_fn(
            dfp_value_key(code),
            u"{} ({})".format(label, code),
            value_type,
            group="Function",
            source="cde"))
    return defs
