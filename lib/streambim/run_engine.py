# -*- coding: utf-8 -*-
"""Shared StreamBIM import engine.

Single home for the config-run logic that was previously duplicated across
the ChecklistImporter, Edit Configs and Run Everything tools.

The engine is split in two stages so modeless callers (the dockable CDE
panel) can keep the Revit UI responsive:

* fetch stage  - HTTP only, safe on a background thread:
                 ``fetch_config_values()`` resolves a config to a
                 ``{ifc_guid: value}`` dict (including grouped checklists).
* apply stage  - Revit API only, must run in API context (a pyRevit command
                 or an ExternalEvent): ``build_element_lookup()`` +
                 ``apply_config_values()``.

``run_config()`` chains both stages for synchronous callers such as the
Run Everything script and the Batch Importer.
"""

import json

import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    StorageType,
    Transaction,
)

from pyrevit import script

logger = script.get_logger()

# Parameter names probed for the IFC GUID match, in order.
IFC_GUID_PARAM_NAMES = ("IFCGuid", "IfcGUID", "IFC GUID")


class ConfigItem(object):
    """A saved StreamBIM import configuration plus runtime counters."""

    def __init__(self, id=None, checklist_id=None, checklist_name=None,
                 streambim_property=None, revit_parameter=None,
                 mapping_enabled=False, mapping_config=None):
        self.id = id
        self.checklist_id = checklist_id
        self.checklist_name = checklist_name
        self.streambim_property = streambim_property
        self.revit_parameter = revit_parameter

        # mapping_enabled arrives as bool or legacy 'True'/'False' string
        if isinstance(mapping_enabled, bool):
            self.mapping_enabled = mapping_enabled
        elif isinstance(mapping_enabled, str):
            self.mapping_enabled = mapping_enabled.lower() == "true"
        else:
            self.mapping_enabled = bool(mapping_enabled)

        self.mapping_config = mapping_config
        self.elements_total = 0
        self.elements_processed = 0
        self.elements_updated = 0
        self.mapping_count = 0
        if mapping_config:
            try:
                self.mapping_count = len(json.loads(mapping_config))
            except Exception:
                self.mapping_count = 0

    @property
    def DisplayName(self):
        return "{} -> {}".format(self.streambim_property, self.revit_parameter)

    @property
    def Status(self):
        if self.elements_total == 0:
            return "Not processed"
        return "{}/{} elements processed, {} updated".format(
            self.elements_processed, self.elements_total, self.elements_updated)

    @property
    def ChecklistName(self):
        return self.checklist_name or "Unknown Checklist"


def config_from_dict(config_dict):
    """Build a ConfigItem from a stored config dict."""
    return ConfigItem(
        id=None,
        checklist_id=config_dict.get('checklist_id'),
        checklist_name=config_dict.get('checklist_name', 'Unknown Checklist'),
        streambim_property=config_dict.get('streambim_property'),
        revit_parameter=config_dict.get('revit_parameter'),
        mapping_enabled=config_dict.get('mapping_enabled'),
        mapping_config=config_dict.get('mapping_config'),
    )


class ChecklistMetadataCache(object):
    """Caches checklist records so grouped-checklist metadata is fetched once per run."""

    def __init__(self, api_client):
        self._client = api_client
        self._records = {}
        self._fetched = False

    def get(self, checklist_id):
        """Return (group_by, building_id) for a checklist; ('', None) when unknown."""
        if checklist_id not in self._records and not self._fetched:
            checklists = self._client.get_checklists()
            if checklists:
                # Only mark as fetched on success so a transient failure is
                # retried by the next config in the same run.
                self._fetched = True
                for checklist in checklists:
                    cid = checklist.get('id')
                    if cid:
                        self._records[cid] = checklist
            else:
                logger.warning("Failed to fetch checklists for metadata lookup: {}".format(
                    self._client.last_error or "empty response"))

        record = self._records.get(checklist_id)
        if not record:
            logger.warning("Checklist {} not found in fetched records".format(checklist_id))
            return ('', None)

        attrs = record.get('attributes', {})
        group_by = attrs.get('group-by', '') or ''

        buildings = record.get('relationships', {}).get('buildings', {}).get('data', [])
        building_id = buildings[0].get('id') if buildings else None

        return (group_by, building_id)


def get_property_value(checklist_item, property_name):
    """Get property value from checklist item.

    Checks both the attributes.properties and items paths in the JSON structure.
    """
    try:
        props = checklist_item.get('attributes', {}).get('properties', {})
        if props and property_name in props:
            return props.get(property_name)

        if 'items' in checklist_item and property_name in checklist_item['items']:
            return checklist_item['items'][property_name]

        return None
    except Exception as e:
        logger.error("Error getting property value: {}".format(str(e)))
        return None


def build_value_mapping(config):
    """Parse the config's mapping table into a {checklist_value: revit_value} dict."""
    value_mapping = {}
    if config.mapping_enabled and config.mapping_config:
        try:
            for mapping in json.loads(config.mapping_config):
                checklist_value = mapping.get('ChecklistValue')
                revit_value = mapping.get('RevitValue')
                if checklist_value and revit_value:
                    value_mapping[checklist_value] = revit_value
        except Exception as e:
            logger.error("Error parsing mapping config: {}".format(str(e)))
    return value_mapping


def set_parameter_value(param, value, storage_type):
    """Set parameter value based on storage type. Returns True when the value changed."""
    try:
        if param.IsReadOnly:
            return False

        if storage_type == StorageType.String:
            str_value = str(value)
            if param.AsString() != str_value:
                param.Set(str_value)
                return True

        elif storage_type == StorageType.Integer:
            try:
                int_value = int(value)
                if param.AsInteger() != int_value:
                    param.Set(int_value)
                    return True
            except (ValueError, TypeError):
                logger.debug("Could not convert '{}' to integer".format(value))
                return False

        elif storage_type == StorageType.Double:
            try:
                double_value = float(value)
                if param.AsDouble() != double_value:
                    param.Set(double_value)
                    return True
            except (ValueError, TypeError):
                logger.debug("Could not convert '{}' to double".format(value))
                return False

        elif storage_type == StorageType.ElementId:
            try:
                element_id = ElementId(int(value))
                if param.AsElementId() != element_id:
                    param.Set(element_id)
                    return True
            except (ValueError, TypeError):
                logger.debug("Could not convert '{}' to ElementId".format(value))
                return False

        return False
    except Exception as e:
        logger.error("Error setting parameter value: {}".format(str(e)))
        return False


def fetch_config_values(api_client, config, metadata_cache, status_cb=None):
    """FETCH STAGE (HTTP only - safe on a background thread).

    Resolves a configuration to a ``{ifc_guid: value}`` dict, applying the
    config's value mapping and, for grouped checklists, resolving group keys
    to IFC GUIDs through the StreamBIM API.

    Args:
        api_client (StreamBIMClient): logged-in client
        config (ConfigItem): configuration to resolve
        metadata_cache (ChecklistMetadataCache): shared per-run metadata cache
        status_cb (callable, optional): status_cb(message) progress reporter

    Returns:
        dict or None: {ifc_guid: value}; None on fetch failure
        (an empty dict means the checklist simply had no matching values).
    """
    def _status(msg):
        if status_cb:
            status_cb(msg)

    if not config.checklist_id:
        logger.info("Skipping configuration - no checklist ID")
        return {}

    _status("Fetching items: {}".format(config.ChecklistName))
    api_client.last_error = None  # clear stale errors from earlier calls
    checklist_items = api_client.get_checklist_items(config.checklist_id, limit=0)
    if not checklist_items:
        if api_client.last_error:
            logger.error("Error retrieving checklist items: {}".format(api_client.last_error))
            return None
        logger.info("No checklist items found for checklist ID: {}".format(config.checklist_id))
        return {}

    logger.info("Retrieved {} checklist items for {}".format(
        len(checklist_items), config.ChecklistName))

    value_mapping = build_value_mapping(config)

    group_by, building_id = metadata_cache.get(config.checklist_id)
    is_grouped = bool(group_by)

    guid_to_value = {}

    if is_grouped:
        if not building_id:
            logger.warning("Cannot process grouped checklist {}: no building ID found".format(
                config.checklist_id))
            return {}

        logger.info("Processing grouped checklist with group-by: {}".format(group_by))
        for item in checklist_items:
            try:
                group_key = item.get('object')
                if not group_key:
                    continue

                property_value = get_property_value(item, config.streambim_property)
                if property_value is None:
                    continue

                if config.mapping_enabled:
                    if property_value in value_mapping:
                        property_value = value_mapping[property_value]
                    else:
                        continue  # mapping enabled but value not mapped -> skip

                _status("Resolving group: {}".format(group_key))
                ifc_guids = api_client.resolve_group_key_to_ifc_guids(
                    config.checklist_id, building_id, group_key)

                for guid in ifc_guids:
                    guid_to_value[guid] = property_value
            except Exception as e:
                logger.error("Error resolving group key: {}".format(str(e)))
                continue

        logger.info("Resolved groups to {} IFC GUID mappings".format(len(guid_to_value)))
    else:
        for item in checklist_items:
            element_id = item.get('object')
            if not element_id:
                element_id = item.get('attributes', {}).get('elementId')
            if not element_id:
                continue

            property_value = get_property_value(item, config.streambim_property)
            if property_value is None:
                continue

            if config.mapping_enabled:
                if property_value in value_mapping:
                    property_value = value_mapping[property_value]
                else:
                    continue

            guid_to_value[element_id] = property_value

    return guid_to_value


def build_element_lookup(doc, wanted_guids):
    """APPLY STAGE helper (Revit API - requires API context).

    Scan the document once and return {ifc_guid: element} for the wanted GUIDs.
    """
    element_lookup = {}
    wanted = set(wanted_guids)
    if not wanted:
        return element_lookup

    all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    for element in all_elements:
        try:
            ifc_guid_param = None
            for param_name in IFC_GUID_PARAM_NAMES:
                ifc_guid_param = element.LookupParameter(param_name)
                if ifc_guid_param:
                    break

            if ifc_guid_param and ifc_guid_param.HasValue:
                guid_value = ifc_guid_param.AsString()
                if guid_value in wanted:
                    element_lookup[guid_value] = element
        except Exception:
            continue

    logger.info("Found {} elements with matching GUIDs".format(len(element_lookup)))
    return element_lookup


def apply_config_values(doc, config, guid_to_value, element_lookup=None):
    """APPLY STAGE (Revit API - requires API context).

    Write the resolved values to the matching elements inside one transaction.

    Args:
        doc: Revit document
        config (ConfigItem): configuration being applied (counters updated)
        guid_to_value (dict): {ifc_guid: value} from fetch_config_values
        element_lookup (dict, optional): prebuilt {ifc_guid: element};
            built on demand when omitted.

    Returns:
        tuple: (processed_count, updated_count)
    """
    processed_count = 0
    updated_count = 0

    if not guid_to_value:
        config.elements_processed = 0
        config.elements_updated = 0
        return (0, 0)

    if element_lookup is None:
        element_lookup = build_element_lookup(doc, guid_to_value.keys())

    t = Transaction(doc, "Batch Import: " + config.DisplayName)
    t.Start()
    try:
        config.elements_total = len(guid_to_value)

        for guid, property_value in guid_to_value.items():
            processed_count += 1
            try:
                element = element_lookup.get(guid)
                if not element:
                    continue

                param = element.LookupParameter(config.revit_parameter)
                if not param or param.IsReadOnly:
                    continue

                if set_parameter_value(param, property_value, param.StorageType):
                    updated_count += 1
            except Exception as e:
                logger.error("Error processing element: {}".format(str(e)))

            if processed_count % 100 == 0:
                logger.debug("Processed {}/{} mappings, updated {} so far".format(
                    processed_count, len(guid_to_value), updated_count))

        t.Commit()

        config.elements_processed = processed_count
        config.elements_updated = updated_count
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        logger.error("Error processing configuration: {}".format(str(e)))

    return (processed_count, updated_count)


def run_config(api_client, doc, config, metadata_cache, status_cb=None, element_lookup=None):
    """Fetch + apply one configuration synchronously (requires API context).

    Returns:
        tuple: (processed_count, updated_count)
    """
    try:
        guid_to_value = fetch_config_values(api_client, config, metadata_cache, status_cb)
        if guid_to_value is None:
            config.elements_processed = 0
            config.elements_updated = 0
            return (0, 0)
        if not guid_to_value:
            logger.info("No element IDs found in checklist items")
            config.elements_processed = 0
            config.elements_updated = 0
            return (0, 0)

        logger.info("Found {} unique element IDs in checklist items".format(len(guid_to_value)))
        return apply_config_values(doc, config, guid_to_value, element_lookup)
    except Exception as e:
        logger.error("Error in run_config: {}".format(str(e)))
        return (0, 0)
