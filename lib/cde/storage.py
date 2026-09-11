# -*- coding: utf-8 -*-
"""Extensible Storage mapping a Revit model to a CDE project + revision.

Mirrors the pattern used by ``lib/streambim/streambim_api.py`` (DataStorage
element + ``BaseSchema`` + pickle/base64 for the free-form prefs blob).
"""
import base64
import pickle

import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector, ExtensibleStorage

from pyrevit import script, revit

import sys
import os.path as op
_lib_dir = op.dirname(op.dirname(__file__))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

from extensible_storage import BaseSchema, simple_field

logger = script.get_logger()

# Current prefs blob version, so we can migrate later if the shape changes.
SCHEMA_VERSION = "1"


class CDEMappingSchema(BaseSchema):
    """Stores the link between this Revit model and a CDE project + revision."""

    # Freshly generated, unique to this schema.
    guid = "d6b1f4e2-7c3a-4e9b-9a1d-2f5c8e0b1a47"

    @simple_field(value_type="string")
    def schema_version():
        """Prefs blob schema version"""

    @simple_field(value_type="string")
    def cde_project_id():
        """The mapped CDE project id"""

    @simple_field(value_type="string")
    def cde_project_name():
        """Human-readable CDE project name (for display)"""

    @simple_field(value_type="string")
    def cde_revision_id():
        """The mapped CDE revision id"""

    @simple_field(value_type="string")
    def base_url():
        """CDE backend base url the mapping was created against"""

    @simple_field(value_type="string")
    def pickled_prefs():
        """Base64-encoded pickle of UI prefs (columns, sync flags, color rule, category)"""


def get_or_create_mapping_storage(doc):
    """Return the existing CDE mapping DataStorage element, creating one if needed."""
    if not doc:
        logger.error("CDE: no active document available")
        return None
    try:
        for ds in FilteredElementCollector(doc)\
                .OfClass(ExtensibleStorage.DataStorage).ToElements():
            try:
                entity = ds.GetEntity(CDEMappingSchema.schema)
                if entity.IsValid():
                    return ds
            except Exception:
                continue
        with revit.Transaction("Create CDE Mapping Storage", doc):
            return ExtensibleStorage.DataStorage.Create(doc)
    except Exception as ex:
        logger.error("CDE: error creating mapping storage: {}".format(ex))
        return None


def _find_mapping_storage(doc):
    """Return the mapping DataStorage element without creating one."""
    if not doc:
        return None
    try:
        for ds in FilteredElementCollector(doc)\
                .OfClass(ExtensibleStorage.DataStorage).ToElements():
            try:
                if ds.GetEntity(CDEMappingSchema.schema).IsValid():
                    return ds
            except Exception:
                continue
    except Exception as ex:
        logger.error("CDE: error locating mapping storage: {}".format(ex))
    return None


def load_mapping(doc):
    """Load the mapping as a plain dict, or None if the model is unmapped.

    Returned keys: ``project_id``, ``project_name``, ``revision_id``,
    ``base_url``, ``prefs`` (dict).
    """
    storage = _find_mapping_storage(doc)
    if not storage:
        return None
    schema = CDEMappingSchema(storage)
    if not schema.is_valid:
        return None
    project_id = schema.get("cde_project_id")
    if not project_id:
        return None

    prefs = {}
    pickled = schema.get("pickled_prefs")
    if pickled:
        try:
            prefs = pickle.loads(base64.b64decode(pickled))
        except Exception as ex:
            logger.warn("CDE: could not decode prefs: {}".format(ex))
            prefs = {}

    return {
        "project_id": project_id,
        "project_name": schema.get("cde_project_name") or "",
        "revision_id": schema.get("cde_revision_id") or "",
        "base_url": schema.get("base_url") or "",
        "prefs": prefs,
    }


def save_mapping(doc, project_id, project_name, revision_id, base_url, prefs=None):
    """Persist (or overwrite) the CDE mapping for this model."""
    storage = get_or_create_mapping_storage(doc)
    if not storage:
        return False
    try:
        encoded_prefs = ""
        if prefs:
            encoded_prefs = base64.b64encode(pickle.dumps(prefs))
        with CDEMappingSchema(storage) as entity:
            entity.set("schema_version", SCHEMA_VERSION)
            entity.set("cde_project_id", project_id or "")
            entity.set("cde_project_name", project_name or "")
            entity.set("cde_revision_id", revision_id or "")
            entity.set("base_url", base_url or "")
            entity.set("pickled_prefs", encoded_prefs)
        return True
    except Exception as ex:
        logger.error("CDE: error saving mapping: {}".format(ex))
        return False


def save_prefs(doc, prefs):
    """Update only the prefs blob, preserving the project/revision mapping."""
    mapping = load_mapping(doc)
    if not mapping:
        return False
    return save_mapping(
        doc,
        mapping["project_id"],
        mapping["project_name"],
        mapping["revision_id"],
        mapping["base_url"],
        prefs,
    )
