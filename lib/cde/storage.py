# -*- coding: utf-8 -*-
"""Extensible Storage mapping a Revit model to a CDE project + revision.

Mirrors the pattern used by ``lib/streambim/streambim_api.py`` (DataStorage
element + ``BaseSchema`` + pickle/base64 for the free-form prefs blob).

Schema v2 adds a dedicated ``cde_revision_etag`` field. Revit ES schemas are
immutable once registered, so v1 (legacy GUID) is read for migration; writes use
v2 and ``BaseSchema`` auto-transfers matching fields from the old entity.
"""
import base64
import pickle

import clr
clr.AddReference("RevitAPI")
import System
from Autodesk.Revit.DB import FilteredElementCollector, ExtensibleStorage

from pyrevit import script, revit

import sys
import os.path as op
_lib_dir = op.dirname(op.dirname(__file__))
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

from extensible_storage import BaseSchema, simple_field
from extensible_storage.entity import Entity

logger = script.get_logger()

# Current prefs blob version, so we can migrate later if the shape changes.
SCHEMA_VERSION = "2"

# v1 schema GUID (no dedicated revision_etag field — etag lived in pickled_prefs).
LEGACY_CDE_MAPPING_GUID = System.Guid("d6b1f4e2-7c3a-4e9b-9a1d-2f5c8e0b1a47")


class CDEMappingSchema(BaseSchema):
    """Stores the link between this Revit model and a CDE project + revision."""

    # v2 — new GUID because ES schemas cannot gain fields after registration.
    guid = "a8c3e1f4-9b2d-4f6a-8e7c-1d5b0a3f6e92"

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
    def cde_revision_etag():
        """Cached NgModelRevision etag for If-Match on mutation commits"""

    @simple_field(value_type="string")
    def base_url():
        """CDE backend base url the mapping was created against"""

    @simple_field(value_type="string")
    def pickled_prefs():
        """Base64-encoded pickle of UI prefs (columns, sync flags, color rule, category)"""


def _legacy_schema():
    """Return the registered v1 schema, if any."""
    try:
        return ExtensibleStorage.Schema.Lookup(LEGACY_CDE_MAPPING_GUID)
    except Exception:
        return None


def _storage_has_mapping_entity(ds):
    """True when *ds* carries a v1 or v2 CDE mapping entity."""
    try:
        if ds.GetEntity(CDEMappingSchema.schema).IsValid():
            return True
        legacy = _legacy_schema()
        if legacy is not None and ds.GetEntity(legacy).IsValid():
            return True
    except Exception:
        pass
    return False


def _open_mapping_reader(storage):
    """Return ``(reader, is_v2)`` where *reader* supports ``.get(name)``."""
    wrapper = CDEMappingSchema(storage)
    if wrapper.is_valid:
        return wrapper, True
    legacy = _legacy_schema()
    if legacy is not None:
        entity = storage.GetEntity(legacy)
        if entity.IsValid():
            return Entity(entity), False
    return None, False


def get_or_create_mapping_storage(doc):
    """Return the existing CDE mapping DataStorage element, creating one if needed."""
    if not doc:
        logger.error("CDE: no active document available")
        return None
    try:
        for ds in FilteredElementCollector(doc)\
                .OfClass(ExtensibleStorage.DataStorage).ToElements():
            if _storage_has_mapping_entity(ds):
                return ds
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
            if _storage_has_mapping_entity(ds):
                return ds
    except Exception as ex:
        logger.error("CDE: error locating mapping storage: {}".format(ex))
    return None


def _revision_etag_from_prefs(prefs):
    if isinstance(prefs, dict):
        return prefs.get("revision_etag") or ""
    return ""


def load_mapping(doc):
    """Load the mapping as a plain dict, or None if the model is unmapped.

    Returned keys: ``project_id``, ``project_name``, ``revision_id``,
    ``revision_etag``, ``base_url``, ``prefs`` (dict).
    """
    storage = _find_mapping_storage(doc)
    if not storage:
        return None
    reader, is_v2 = _open_mapping_reader(storage)
    if reader is None:
        return None
    project_id = reader.get("cde_project_id")
    if not project_id:
        return None

    prefs = {}
    pickled = reader.get("pickled_prefs")
    if pickled:
        try:
            prefs = pickle.loads(base64.b64decode(pickled))
        except Exception as ex:
            logger.warn("CDE: could not decode prefs: {}".format(ex))
            prefs = {}

    revision_etag = ""
    if is_v2:
        revision_etag = reader.get("cde_revision_etag") or ""
    if not revision_etag:
        revision_etag = _revision_etag_from_prefs(prefs)

    stored_base_url = reader.get("base_url") or ""
    try:
        from cde import config as cde_config
        config_base_url = cde_config.get_base_url()
    except Exception:
        config_base_url = ""
    effective_base_url = (stored_base_url or config_base_url).rstrip("/")
    logger.debug(
        "CDE mapping URLs — stored: '{}', config default: '{}', effective: '{}'".format(
            stored_base_url, config_base_url, effective_base_url))

    return {
        "project_id": project_id,
        "project_name": reader.get("cde_project_name") or "",
        "revision_id": reader.get("cde_revision_id") or "",
        "revision_etag": revision_etag,
        "base_url": stored_base_url,
        "prefs": prefs,
    }


def save_mapping(doc, project_id, project_name, revision_id, base_url,
                 prefs=None, revision_etag=None):
    """Persist (or overwrite) the CDE mapping for this model."""
    storage = get_or_create_mapping_storage(doc)
    if not storage:
        return False
    try:
        encoded_prefs = ""
        prefs_dict = dict(prefs) if prefs else {}
        existing_etag = ""
        existing_revision = ""
        try:
            existing = load_mapping(doc)
            if existing:
                existing_etag = existing.get("revision_etag") or ""
                existing_revision = existing.get("revision_id") or ""
                if not prefs_dict and existing.get("prefs"):
                    prefs_dict = dict(existing.get("prefs") or {})
        except Exception:
            pass
        if revision_etag is not None:
            effective_etag = revision_etag or ""
        elif str(revision_id or "") != str(existing_revision or ""):
            effective_etag = ""
        else:
            effective_etag = existing_etag
        prefs_dict.pop("revision_etag", None)
        if prefs_dict:
            encoded_prefs = base64.b64encode(pickle.dumps(prefs_dict))
        with CDEMappingSchema(storage) as entity:
            entity.set("schema_version", SCHEMA_VERSION)
            entity.set("cde_project_id", project_id or "")
            entity.set("cde_project_name", project_name or "")
            entity.set("cde_revision_id", revision_id or "")
            entity.set("cde_revision_etag", effective_etag)
            entity.set("base_url", base_url or "")
            entity.set("pickled_prefs", encoded_prefs)
        return True
    except Exception as ex:
        logger.error("CDE: error saving mapping: {}".format(ex))
        return False


def update_revision_etag(doc, revision_etag):
    """Update only the cached revision etag, preserving other mapping fields."""
    storage = _find_mapping_storage(doc)
    if not storage:
        return False
    mapping = load_mapping(doc)
    if not mapping:
        return False
    try:
        prefs = dict(mapping.get("prefs") or {})
        prefs.pop("revision_etag", None)
        encoded_prefs = ""
        if prefs:
            encoded_prefs = base64.b64encode(pickle.dumps(prefs))
        with CDEMappingSchema(storage) as entity:
            entity.set("cde_revision_etag", revision_etag or "")
            if encoded_prefs or mapping.get("prefs"):
                entity.set("pickled_prefs", encoded_prefs)
        return True
    except Exception as ex:
        logger.error("CDE: error updating revision etag: {}".format(ex))
        return False


def save_prefs(doc, prefs):
    """Update only the prefs blob, preserving the project/revision mapping."""
    mapping = load_mapping(doc)
    if not mapping:
        return False
    clean_prefs = dict(prefs) if prefs else {}
    clean_prefs.pop("revision_etag", None)
    return save_mapping(
        doc,
        mapping["project_id"],
        mapping["project_name"],
        mapping["revision_id"],
        mapping["base_url"],
        clean_prefs,
        revision_etag=mapping.get("revision_etag"),
    )
