# -*- coding: utf-8 -*-
"""Domain-level API abstraction over the CDE backend.

The schedule UI talks only to this layer, never to raw REST/GraphQL. Endpoints
that are confirmed (auth, GraphQL ``element``, projects) are wired directly;
endpoints that are still TODO on the backend (element-list-by-IfcClass,
per-category parameter definitions and their CRUD, symbols) are isolated behind
clearly marked methods so they can be filled in without touching the UI.

``MockCDEService`` mirrors the same interface with in-memory sample data so the
table / coloring UI (Phases 2-3) can be developed without a live backend.
"""
from collections import namedtuple
import hashlib
import json
import os
import time

from pyrevit import script

from cde import config
from cde.api import CDEClient, CDEApiError
from cde.dfp_catalog import DFP_PSET_NAME

logger = script.get_logger()

# --- Domain types ---------------------------------------------------------

Project = namedtuple("Project", ["id", "name"])
Revision = namedtuple("Revision", ["id", "name", "is_current"])
# id = NgModelRevision id (GraphQL revisionId); name = IFC file name.
Model = namedtuple("Model", ["id", "name", "version_id", "file_id", "is_projected"])
# value_type: "string" | "bool" | "number" | "enum"
ParameterDef = namedtuple(
    "ParameterDef",
    ["key", "label", "value_type", "allowed_values", "group", "source"])
# values: {param_key: value}
CDEElement = namedtuple("CDEElement", ["global_id", "ifc_class", "name", "values"])

MutationOutcome = namedtuple(
    "MutationOutcome",
    ["dry_run", "success", "mutation_id", "plan", "status_data", "etag"])

MUTATION_TERMINAL_STATES = frozenset([
    "completed", "succeeded", "failed", "rejected", "cancelled", "error"])
MUTATION_SUCCESS_STATES = frozenset(["completed", "succeeded"])

DEFAULT_IFC_CLASS = "IfcDoor"
# Door schedule: match both plain doors and standard-case doors in the graph.
DOOR_IFC_CLASSES = ("IFCDOOR", "IFCDOORSTANDARDCASE")
# GraphQL value roles merged onto each element (later roles win on conflict).
_VALUE_MERGE_ORDER = ("authoredValues", "derivedValues", "effectiveValues")
# Relationship types that link a door to its DFP-carrying IfcGroup.
_DFP_RELATIONSHIP_TYPES = frozenset(["GROUPS", "ASSIGNED_TO"])
# GraphQL PropertyValue fields per contract.graphql
_PROPERTY_VALUE_FIELDS = (
    " psetName propertyName value datatype unit state sourceKind")
# Node fields on element queries (elements + element).
_ELEMENT_NODE_FIELDS = (
    " globalId ifcClass name etag"
    " effectiveValues{" + _PROPERTY_VALUE_FIELDS + "}"
    " derivedValues{" + _PROPERTY_VALUE_FIELDS + "}"
    " authoredValues{" + _PROPERTY_VALUE_FIELDS + "}"
    " relationships{ type subjectGlobalId objectGlobalId }")


def _elements_graphql_query(include_filter):
    """Build elements() query matching contract.graphql."""
    if include_filter:
        return (
            "query($projectId: ID!, $revisionId: ID!, $first: Int, "
            "$after: String, $filter: ElementFilterInput) {"
            " elements(projectId: $projectId, revisionId: $revisionId, "
            "filter: $filter, first: $first, after: $after) {"
            " edges { cursor node {" + _ELEMENT_NODE_FIELDS + " } }"
            " pageInfo { hasNextPage endCursor } } }")
    return (
        "query($projectId: ID!, $revisionId: ID!, $first: Int, $after: String) {"
        " elements(projectId: $projectId, revisionId: $revisionId, "
        "first: $first, after: $after) {"
        " edges { cursor node {" + _ELEMENT_NODE_FIELDS + " } }"
        " pageInfo { hasNextPage endCursor } } }")


def _element_graphql_query():
    """Build element(globalId) query matching contract.graphql."""
    return (
        "query($projectId: ID!, $revisionId: ID!, $globalId: String!) {"
        " element(projectId: $projectId, revisionId: $revisionId, "
        "globalId: $globalId) {"
        + _ELEMENT_NODE_FIELDS
        + " } }")


def param_def(key, label, value_type="string", allowed_values=None, group="",
              source="cde"):
    return ParameterDef(
        key, label, value_type, allowed_values or [], group, source)


def normalize_ifc_class(value):
    """Canonical form for comparing IfcClass across sources.

    The live graph stores classes uppercase without separators
    (``IFCDOOR``); Revit/our categories use ``IfcDoor``. Compare on upper-case.
    """
    return (value or "").upper().replace("_", "")


def want_ifc_classes(ifc_class):
    """Normalized IfcClass set to match for a schedule category."""
    if normalize_ifc_class(ifc_class) == normalize_ifc_class(DEFAULT_IFC_CLASS):
        return frozenset(normalize_ifc_class(c) for c in DOOR_IFC_CLASSES)
    if ifc_class:
        return frozenset([normalize_ifc_class(ifc_class)])
    return frozenset()


def node_matches_ifc_classes(node, want_classes):
    if not want_classes:
        return True
    return normalize_ifc_class(node.get("ifcClass")) in want_classes


def coerce_dfp_cell_value(value):
    """Hub parity: backend stores IfcBoolean as ``\"true\"``/``\"false\"`` strings."""
    if value == "true":
        return True
    if value == "false":
        return False
    return value


def _extract_etag(data):
    """Best-effort NgModelRevision.etag from assorted backend payloads."""
    if not isinstance(data, dict):
        return ""
    for key in ("etag", "revisionEtag", "ngModelRevisionEtag"):
        val = data.get(key)
        if val:
            return str(val)
    for nested_key in ("ngModelRevision", "revision", "modelRevision"):
        nested = data.get(nested_key)
        if isinstance(nested, dict):
            for key in ("etag", "revisionEtag"):
                val = nested.get(key)
                if val:
                    return str(val)
    return ""


def _mutation_value(value):
    """Serialize a cell value for the mutations REST payload."""
    if isinstance(value, bool):
        return "true" if value else "false"
    if value is None:
        return None
    return value


def _build_mutation_operations(changes):
    """Build REST operations from {global_id: {property: value}}."""
    operations = []
    for gid in sorted(changes.keys()):
        props = changes.get(gid) or {}
        for prop in sorted(props.keys()):
            val = props[prop]
            if val is None or val == "":
                operations.append({
                    "op": "unset",
                    "property": prop,
                    "objectId": gid,
                })
            else:
                operations.append({
                    "op": "set",
                    "property": prop,
                    "value": _mutation_value(val),
                    "objectId": gid,
                })
    return operations


def _idempotency_key(revision_id, object_ids, operations, dry_run):
    payload = {
        "revision_id": revision_id,
        "object_ids": sorted(object_ids),
        "operations": operations,
        "dry_run": dry_run,
    }
    raw = json.dumps(payload, sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _parse_dry_run_plan(data):
    if not isinstance(data, dict):
        return {}
    plan = data.get("dryRunPlan") or data.get("plan") or data
    if not isinstance(plan, dict):
        return {}
    return plan


class CDEService(object):
    """Live implementation backed by :class:`CDEClient`."""

    def __init__(self, auth):
        self.auth = auth
        self.client = CDEClient(auth)
        # Set by list_elements: {"retrieved": n, "total": t} when the backend
        # returned fewer element nodes than the revision actually has, else None.
        self.last_truncation = None
        self.last_fetch_mode = None
        self.last_fetch_pages = 0
        self.last_fetch_nodes = 0

    # --- projects / revisions ------------------------------------------

    def list_projects(self):
        """Return active [Project]. GET /api/v1/projects (``is_active`` only)."""
        data = self.client.get(config.projects_url(self.client.base_url))
        items = data.get("projects", data) if isinstance(data, dict) else data
        logger.info("CDE: list_projects raw count={}".format(len(items or [])))
        if items:
            sample = items[0]
            if isinstance(sample, dict):
                logger.debug("CDE: list_projects sample keys={}".format(
                    sorted(sample.keys())))
        projects = []
        for item in (items or []):
            if not item.get("is_active", False):
                continue
            pid = item.get("id") or item.get("project_id")
            name = item.get("name") or item.get("display_name") or str(pid)
            if pid:
                projects.append(Project(str(pid), name))
        logger.info("CDE: list_projects active count={}".format(len(projects)))
        return projects

    def list_models(self, project_id):
        """Return [Model] for a project's uploaded IFC files.

        Files come from ifc-versions; each finalized version is resolved to its
        live-graph ``revisionId`` via ``/versions/{id}/ingest-status`` (the
        universal bridge - works whether the model was live-dropped or ingested
        through the normal upload/projection pipeline).

        ``Model.id`` is that revision id - the value :meth:`list_elements` and
        the Login mapping require - or "" when the model has not been ingested
        into the live graph yet (Login refuses to map an empty id and tells the
        user to wait for ingest). This makes the schedule resolve the *selected*
        model's own revision. ``Model.is_projected`` reflects graph-sync state.
        """
        versions = self._fetch_ifc_versions(project_id)
        if versions is None:
            versions = self._fetch_live_drops(project_id) or []

        models = []
        for item in versions:
            version_id = item.get("version_id")
            name = item.get("original_name") or item.get("name") or str(version_id or "")
            file_id = item.get("file_id")
            # Only finalized/projected versions can have a graph revision.
            finalized = bool(item.get("finalized_at") or item.get("fragments_generated_at"))
            rev, synced = ("", False)
            if version_id and finalized:
                rev, synced = self._revision_for_version(version_id)
            models.append(Model(
                str(rev or ""), name, str(version_id or ""),
                str(file_id or ""), synced))
        queryable = sum(1 for m in models if m.id)
        logger.info("CDE: list_models project={} files={} ingested(queryable)={}".format(
            project_id, len(models), queryable))
        return models

    def _revision_for_version(self, version_id):
        """Resolve a version's live-graph (revisionId, graph_synced) via
        ingest-status. Returns ("", False) when not ingested/unavailable."""
        try:
            data = self.client.get(
                config.ingest_status_url(version_id, self.client.base_url))
        except CDEApiError as ex:
            logger.debug("CDE: ingest-status unavailable for {}: {}".format(
                version_id, ex))
            return "", False
        if not isinstance(data, dict):
            return "", False
        rev = data.get("revisionId") or data.get("revision_id") or ""
        synced = bool(data.get("graphSynced") or data.get("projected"))
        completed = data.get("status") in (None, "completed")
        return (rev if (rev and synced and completed) else ""), synced

    def _fetch_ifc_versions(self, project_id):
        """GET /api/v1/projects/{id}/ifc-versions → list | None on error."""
        try:
            data = self.client.get(
                config.ifc_versions_url(project_id, self.client.base_url))
            items = data if isinstance(data, list) else (data or {}).get("items")
            if items is None:
                return None
            logger.info("CDE: ifc-versions project={} count={}".format(
                project_id, len(items)))
            if items:
                logger.debug("CDE: ifc-versions sample keys={}".format(
                    sorted(items[0].keys())))
            return items
        except CDEApiError as ex:
            logger.debug("CDE: ifc-versions unavailable for {}: {}".format(
                project_id, ex))
            return None

    def _fetch_live_drops(self, project_id):
        """GET /api/v1/projects/{id}/live-drops → list | None on error.

        Each drop carries ``version_id``, ``file_id`` and (when ingested into
        the live graph) ``ingest.revisionId`` / ``ingest.graphSynced``. Returned
        raw; list_models joins it to ifc-versions by version_id.
        """
        try:
            data = self.client.get(
                config.live_drops_url(project_id, self.client.base_url))
            items = data if isinstance(data, list) else (data or {}).get("items")
            if items is None:
                return None
            ingested = sum(1 for it in items if (it.get("ingest") or {}).get("revisionId"))
            logger.info("CDE: live-drops project={} count={} ingested={}".format(
                project_id, len(items), ingested))
            return items
        except CDEApiError as ex:
            logger.debug("CDE: live-drops unavailable for {}: {}".format(
                project_id, ex))
            return None

    def list_revisions(self, project_id):
        """Return [Revision] for a project.

        TODO(api): confirm the revisions endpoint. We try the graph overview/
        status surface which is revision-scoped; until a dedicated list exists
        the backend may expose revisions there.
        """
        try:
            data = self.client.get(
                config.graph_overview_url(self.client.base_url),
                params={"projectId": project_id})
        except CDEApiError as ex:
            logger.warn("CDE: list_revisions fallback: {}".format(ex))
            data = {}
        revisions = []
        for item in (data.get("revisions", []) if isinstance(data, dict) else []):
            rid = item.get("id") or item.get("revisionId")
            if rid:
                revisions.append(Revision(
                    str(rid),
                    item.get("name") or str(rid),
                    bool(item.get("isCurrent") or item.get("is_current"))))
        return revisions

    # --- elements -------------------------------------------------------

    # Filtered door fetch: POST /api/v2/graphql with
    # filter.ifcClass / filter.ifcClasses (see read-api.graphql). Page size 500
    # matches Hub nextgenGraph.allElements. Full-scan fallback pages ~1000 edges.
    DOOR_PAGE_SIZE = 500
    FULL_SCAN_PAGE_SIZE = 1000

    def list_elements(self, project_id, revision_id, ifc_class=DEFAULT_IFC_CLASS):
        """Return [CDEElement] of the given IfcClass for a project+revision.

        Live-graph reality (probed against the running backend):
        * ``elements`` is a Relay connection: ``{ edges{ node{...} } pageInfo }``.
        * Prefer ``ifcClass`` on the GraphQL query (server filter). If page 0
          contains other classes, fall back to a full revision scan and filter
          client-side (observed when the gateway ignores ``ifcClass``).
        * IfcClass is stored upper-case w/o separators (``IFCDOOR``).
        * ``revisionId`` MUST be the live-graph ingest revision id (from
          live-drops ``ingest.revisionId``), NOT the ifc-version ``version_id``.
        * Variables ``projectId``/``revisionId`` are read from the variables map
          by name (the gateway does not honour declared query variables).
        Returns [] on failure so the UI degrades gracefully.
        """
        # NOTE: the gateway reads 'first' from the variables map (the query-string
        # literal is ignored), so 'first' MUST be passed as a variable or it
        # silently defaults to 50.
        want_classes = want_ifc_classes(ifc_class)
        nodes, fetch_mode = self._fetch_element_nodes(
            project_id, revision_id, ifc_class)
        self.last_fetch_mode = fetch_mode
        self.last_fetch_nodes = len(nodes)
        if fetch_mode == "full_scan":
            self.last_truncation = self._truncation(project_id, revision_id, len(nodes))
        else:
            self.last_truncation = None
        node_by_gid = {
            n.get("globalId"): n for n in nodes if n.get("globalId")}
        self.last_node_by_gid = node_by_gid
        elements = []
        group_dfp_merged = 0
        for node in nodes:
            if fetch_mode == "full_scan" and want_classes:
                if not node_matches_ifc_classes(node, want_classes):
                    continue
            el = self._element_from_node(node)
            if want_classes and node_matches_ifc_classes(node, want_classes):
                merged = dict(el.values)
                if self._inherit_group_values(merged, node, node_by_gid):
                    group_dfp_merged += 1
                    el = el._replace(values=merged)
            elements.append(el)
        logger.info("CDE: list_elements {} of {} nodes match {}".format(
            len(elements), len(nodes), ifc_class))
        if elements:
            el0 = elements[0]
            all_keys = sorted(el0.values.keys())
            pset_prefixes = sorted(set(
                k.split(".", 1)[0] for k in all_keys if "." in k))
            logger.info(
                "CDE: sample door keys={} psets={}".format(
                    len(all_keys), pset_prefixes[:10]))
        return elements

    # Backend returns ~1000 edges per GraphQL page regardless of `first`.
    # Full-scan fallback: capped for UI responsiveness (was 15 pages / ~15k nodes).
    MAX_ELEMENT_PAGES = 15
    FULL_SCAN_MAX_PAGES = 3

    def _fetch_element_nodes(self, project_id, revision_id, ifc_class):
        """Fetch element nodes, preferring server-side ``filter.ifcClass(es)``."""
        want_classes = want_ifc_classes(ifc_class)
        if want_classes:
            filter_classes = list(DOOR_IFC_CLASSES) if (
                want_classes == want_ifc_classes(DEFAULT_IFC_CLASS)) else list(want_classes)
            for filter_key in ("ifcClasses", "ifcClass"):
                nodes, ok, pages = self._fetch_all_element_nodes(
                    project_id, revision_id,
                    page_size=self.DOOR_PAGE_SIZE,
                    filter_key=filter_key,
                    filter_classes=filter_classes,
                    want_classes=want_classes)
                if ok:
                    self.last_fetch_pages = pages
                    return nodes, "ifc_class_filter"
            logger.debug("CDE: server ignored filter.ifcClass(es); using capped revision scan")
        nodes, _ok, pages = self._fetch_all_element_nodes(
            project_id, revision_id,
            page_size=self.FULL_SCAN_PAGE_SIZE,
            filter_key=None, filter_classes=None, want_classes=None,
            max_pages=self.FULL_SCAN_MAX_PAGES)
        self.last_fetch_pages = pages
        return nodes, "full_scan"

    def _fetch_all_element_nodes(self, project_id, revision_id,
                                 page_size, filter_key=None,
                                 filter_classes=None, want_classes=None,
                                 max_pages=None):
        """Page ``elements(projectId, revisionId, filter?, first, after)``."""
        page_limit = max_pages if max_pages is not None else self.MAX_ELEMENT_PAGES
        filtered = bool(filter_key and filter_classes)
        query = _elements_graphql_query(filtered)
        nodes = []
        seen_cursors = set()
        after = None
        pages_fetched = 0
        filter_rejected = False
        for page_num in range(page_limit):
            pages_fetched += 1
            variables = {
                "projectId": project_id,
                "revisionId": revision_id,
                "first": page_size,
                "after": after,
            }
            if filtered:
                variables["filter"] = {filter_key: list(filter_classes)}
            prev_total = len(nodes)
            try:
                data = self.client.graphql(query, variables)
            except CDEApiError as ex:
                logger.warn("CDE: list_elements failed: {}".format(ex))
                break
            conn = data.get("elements") or {}
            edges = conn.get("edges") or []
            page_info = conn.get("pageInfo") or {}
            fresh = [e for e in edges if (e.get("cursor") not in seen_cursors)]
            if not fresh:
                break
            if filtered and want_classes and page_num == 0:
                mismatched = sum(
                    1 for edge in fresh
                    if not node_matches_ifc_classes((edge or {}).get("node") or {}, want_classes))
                if mismatched or len(fresh) >= 900:
                    filter_rejected = True
                    break
            for edge in fresh:
                cursor = edge.get("cursor")
                if cursor is not None:
                    seen_cursors.add(cursor)
                node = (edge or {}).get("node") or {}
                if node:
                    nodes.append(node)
            next_after = page_info.get("endCursor")
            if len(nodes) == prev_total:
                break
            if not page_info.get("hasNextPage") or not next_after or next_after == after:
                break
            after = next_after
        logger.info("CDE: list_elements rev={} raw nodes={}".format(
            revision_id, len(nodes)))
        if filtered:
            return nodes, not filter_rejected and bool(nodes), pages_fetched
        return nodes, True, pages_fetched

    def _truncation(self, project_id, revision_id, retrieved):
        """Compare retrieved node count to the authoritative graph total.

        Returns {"retrieved", "total"} when fewer were fetched than exist
        (the current ~200-node gateway cap), else None.
        """
        try:
            data = self.client.get(
                config.graph_status_url(self.client.base_url),
                params={"projectId": project_id, "revisionId": revision_id})
        except CDEApiError as ex:
            logger.debug("CDE: graph status unavailable: {}".format(ex))
            return None
        total = ((data or {}).get("postgres") or {}).get("elements")
        if isinstance(total, int) and retrieved < total:
            logger.warn("CDE: elements TRUNCATED - retrieved {} of {} nodes "
                        "(backend elements-query cap; schedule is incomplete "
                        "until paging/cap is fixed server-side)".format(
                            retrieved, total))
            return {"retrieved": retrieved, "total": total}
        return None

    def get_element(self, project_id, revision_id, global_id):
        """Return a single CDEElement by GlobalId (live-graph ``element`` field)."""
        node = self._fetch_element_node(project_id, revision_id, global_id)
        return self._element_from_node(node) if node else None

    def fetch_element_detail(self, project_id, revision_id, global_id,
                             node_by_gid=None):
        """Fetch one element and merge group-carried values when *node_by_gid* is set."""
        node = self._fetch_element_node(project_id, revision_id, global_id)
        if not node:
            return None
        el = self._element_from_node(node)
        lookup = node_by_gid if node_by_gid is not None else {}
        if lookup is None:
            lookup = {}
        merged = dict(el.values)
        if self._inherit_group_values(merged, node, lookup):
            el = el._replace(values=merged)
        return el

    def _fetch_element_node(self, project_id, revision_id, global_id):
        query = _element_graphql_query()
        variables = {"projectId": project_id, "revisionId": revision_id,
                     "globalId": global_id}
        try:
            data = self.client.graphql(query, variables)
        except CDEApiError as ex:
            logger.warn("CDE: get_element({}) failed: {}".format(global_id, ex))
            return None
        return data.get("element")

    def _merge_values_from_node(self, node):
        """Merge authored, derived, and effective values from a graph node."""
        values = {}
        for source in _VALUE_MERGE_ORDER:
            for ev in (node.get(source) or []):
                pset = ev.get("psetName") or ""
                name = ev.get("propertyName")
                if name is not None:
                    key = "{}.{}".format(pset, name) if pset else name
                    val = ev.get("value")
                    if pset == DFP_PSET_NAME:
                        val = coerce_dfp_cell_value(val)
                    values[key] = val
        return values

    def _element_from_node(self, node):
        if not node:
            return CDEElement(None, None, None, {})
        return CDEElement(
            node.get("globalId"),
            node.get("ifcClass"),
            node.get("name"),
            self._merge_values_from_node(node))

    def _group_global_ids_for_door(self, door_node):
        """Return GlobalIds of related groups that may carry ``Pset_DFP``."""
        door_gid = door_node.get("globalId")
        if not door_gid:
            return []
        group_gids = []
        for rel in (door_node.get("relationships") or []):
            rel_type = (rel.get("type") or "").upper()
            if rel_type not in _DFP_RELATIONSHIP_TYPES:
                continue
            subj = rel.get("subjectGlobalId")
            obj = rel.get("objectGlobalId")
            if obj == door_gid and subj:
                group_gids.append(subj)
            elif subj == door_gid and obj:
                group_gids.append(obj)
        return group_gids

    def _inherit_group_values(self, door_values, door_node, node_by_gid):
        """Copy effective property values from linked IfcGroup nodes onto a door."""
        merged_any = False
        for group_gid in self._group_global_ids_for_door(door_node):
            group_node = node_by_gid.get(group_gid)
            if not group_node:
                continue
            for key, val in self._merge_values_from_node(group_node).items():
                if val is None or val == "":
                    continue
                if key not in door_values or door_values.get(key) in (None, ""):
                    door_values[key] = val
                    merged_any = True
        return merged_any

    def _inherit_group_dfp(self, door_values, door_node, node_by_gid):
        """Deprecated alias — use :meth:`_inherit_group_values`."""
        return self._inherit_group_values(door_values, door_node, node_by_gid)

    # --- parameter definitions / values --------------------------------

    def get_parameter_defs(self, project_id, ifc_class=DEFAULT_IFC_CLASS):
        """Return [ParameterDef] for a category (DFP door-function catalog for doors).

        Values still come from GraphQL element payloads (derived/effective and
        linked IfcGroup ``Pset_DFP``). A Postgres-backed override endpoint can
        extend this later.
        """
        if normalize_ifc_class(ifc_class) != normalize_ifc_class(DEFAULT_IFC_CLASS):
            return []
        try:
            from cde.dfp_catalog import build_dfp_parameter_defs
            return build_dfp_parameter_defs(param_def)
        except Exception as ex:
            logger.debug("CDE: dfp parameter defs unavailable: {}".format(ex))
            return []

    def get_element_values(self, project_id, revision_id, global_ids, params=None):
        """Return {global_id: {param_key: value}} for the requested elements.

        Default implementation fetches per-element via the confirmed
        ``element`` GraphQL field. TODO(api): replace with a batch endpoint for
        large selections.
        """
        result = {}
        for gid in global_ids:
            try:
                element = self.get_element(project_id, revision_id, gid)
            except CDEApiError as ex:
                logger.warn("CDE: get_element_values({}) failed: {}".format(gid, ex))
                continue
            if element:
                result[gid] = element.values
        return result

    def set_element_values(self, project_id, revision_id, global_ids, values,
                           dry_run=True, etag=None, idempotency_key=None):
        """Write ``values`` (dict) across ``global_ids`` via REST mutations.

        Convenience wrapper around :meth:`apply_element_mutations` for the
        bulk-set toolbar (same value on every selected row).
        """
        changes = {}
        for gid in global_ids:
            changes[gid] = dict(values or {})
        return self.apply_element_mutations(
            project_id, revision_id, changes,
            dry_run=dry_run, etag=etag, idempotency_key=idempotency_key)

    def apply_element_mutations(self, project_id, revision_id, changes,
                                dry_run=True, etag=None, idempotency_key=None):
        """POST /api/v2/mutations (transactional lane).

        ``changes`` is ``{global_id: {property_key: value}}``. When
        ``dry_run`` is True (default) the backend returns a DryRunPlan; commit
        with ``dry_run=False`` and a fresh ``If-Match`` etag.
        """
        changes = changes or {}
        object_ids = sorted(changes.keys())
        if not object_ids:
            return MutationOutcome(
                dry_run=dry_run, success=True, mutation_id=None,
                plan={}, status_data=None, etag=etag or "")

        operations = _build_mutation_operations(changes)
        if not operations:
            return MutationOutcome(
                dry_run=dry_run, success=True, mutation_id=None,
                plan={}, status_data=None, etag=etag or "")

        payload = {
            "lane": "transactional",
            "projectId": project_id,
            "revisionId": revision_id,
            "objectIds": object_ids,
            "operations": operations,
            "dryRun": bool(dry_run),
        }
        headers = {
            "Idempotency-Key": idempotency_key or _idempotency_key(
                revision_id, object_ids, operations, dry_run),
        }
        if not dry_run:
            if not etag:
                etag = self.fetch_revision_etag(project_id, revision_id)
            if not etag:
                raise CDEApiError(
                    "Cannot commit without revision etag (If-Match required).")
            headers["If-Match"] = etag

        resp = self.client.post(
            config.mutations_url(self.client.base_url),
            payload,
            extra_headers=headers,
            return_meta=True)

        data = resp.data if isinstance(resp.data, dict) else {}
        new_etag = (
            resp.headers.get("ETag")
            or resp.headers.get("Etag")
            or resp.headers.get("etag")
            or _extract_etag(data)
            or etag
            or "")

        if dry_run:
            plan = _parse_dry_run_plan(data)
            return MutationOutcome(
                dry_run=True, success=True, mutation_id=None,
                plan=plan, status_data=data, etag=new_etag)

        mutation_id = (
            data.get("id") or data.get("mutationId")
            or data.get("mutation_id"))
        return MutationOutcome(
            dry_run=False, success=True,
            mutation_id=str(mutation_id) if mutation_id else None,
            plan=None, status_data=data, etag=new_etag)

    def poll_mutation(self, mutation_id, max_attempts=12, interval_seconds=0.5):
        """Poll GET /api/v2/mutations/{id} until a terminal state."""
        if not mutation_id:
            return {}
        last = {}
        for attempt in range(max_attempts):
            resp = self.client.get(
                config.mutation_status_url(mutation_id, self.client.base_url),
                return_meta=True)
            last = resp.data if isinstance(resp.data, dict) else {}
            state = (
                (last.get("status") or last.get("state") or "")
                .lower())
            if state in MUTATION_TERMINAL_STATES:
                return last
            if attempt < max_attempts - 1:
                time.sleep(interval_seconds)
        return last

    def fetch_revision_etag(self, project_id, revision_id):
        """Return NgModelRevision.etag for If-Match optimistic locking.

        TODO(api): confirm exact field name / endpoint when OpenAPI stabilizes.
        """
        try:
            data = self.client.get(
                config.graph_status_url(self.client.base_url),
                params={"projectId": project_id, "revisionId": revision_id})
            etag = _extract_etag(data)
            if etag:
                return etag
        except CDEApiError as ex:
            logger.debug("CDE: graph status etag lookup failed: {}".format(ex))

        try:
            data = self.client.get(
                config.graph_overview_url(self.client.base_url),
                params={"projectId": project_id, "revisionId": revision_id})
            etag = _extract_etag(data)
            if etag:
                return etag
            for item in (data.get("revisions") or []):
                rid = item.get("id") or item.get("revisionId")
                if str(rid) == str(revision_id):
                    etag = _extract_etag(item)
                    if etag:
                        return etag
        except CDEApiError as ex:
            logger.debug("CDE: graph overview etag lookup failed: {}".format(ex))

        return ""

    def get_symbols(self, project_id):
        """Return symbol metadata for legends/glyphs. TODO(api)."""
        return []


class MockCDEService(object):
    """In-memory sample data mirroring CDEService for offline UI development."""

    def __init__(self, auth=None):
        self.auth = auth

    def list_projects(self):
        return [Project("demo-1", "Demo Project A"),
                Project("demo-2", "Demo Project B")]

    def list_revisions(self, project_id):
        return [Revision("rev-2", "Revision 2", True),
                Revision("rev-1", "Revision 1", False)]

    def list_models(self, project_id):
        return [
            Model("rev-2", "Demo-Model-A.ifc", "ver-2", "file-2", True),
            Model("rev-1", "Demo-Model-B.ifc", "ver-1", "file-1", True),
        ]

    def get_parameter_defs(self, project_id, ifc_class=DEFAULT_IFC_CLASS):
        return [
            param_def("needs_handle", "Needs handle", "bool", group="Hardware"),
            param_def("needs_lock", "Needs lock", "bool", group="Hardware"),
            param_def("lock_type", "Lock type", "enum",
                      ["Cylinder", "Mortise", "Electronic"], group="Hardware"),
            param_def("fire_escape", "Fire escape zone", "bool", group="Function"),
            param_def("access_class", "Access class", "enum",
                      ["A", "B", "C"], group="Function"),
        ]

    def list_elements(self, project_id, revision_id, ifc_class=DEFAULT_IFC_CLASS):
        samples = []
        for i in range(1, 9):
            samples.append(CDEElement(
                "GID-{:04d}".format(i), ifc_class, "Door {}".format(i),
                {
                    "needs_handle": bool(i % 2),
                    "needs_lock": bool(i % 3),
                    "lock_type": ["Cylinder", "Mortise", "Electronic"][i % 3],
                    "fire_escape": bool(i % 4 == 0),
                    "access_class": ["A", "B", "C"][i % 3],
                }))
        return samples

    def get_element_values(self, project_id, revision_id, global_ids, params=None):
        elements = {e.global_id: e.values
                    for e in self.list_elements(project_id, revision_id)}
        return {gid: elements.get(gid, {}) for gid in global_ids}

    def set_element_values(self, project_id, revision_id, global_ids, values,
                           dry_run=True, etag=None, idempotency_key=None):
        return self.apply_element_mutations(
            project_id, revision_id,
            {gid: dict(values or {}) for gid in global_ids},
            dry_run=dry_run, etag=etag, idempotency_key=idempotency_key)

    def apply_element_mutations(self, project_id, revision_id, changes,
                                dry_run=True, etag=None, idempotency_key=None):
        logger.info("CDE(mock): apply mutations dry_run={} on {} element(s)".format(
            dry_run, len(changes or {})))
        object_ids = sorted((changes or {}).keys())
        if dry_run:
            plan = {
                "matchedElements": object_ids,
                "conflicts": [],
                "candidateValues": dict(changes or {}),
                "impactedRelationships": [],
            }
            return MutationOutcome(
                dry_run=True, success=True, mutation_id=None,
                plan=plan, status_data={"dryRunPlan": plan},
                etag=etag or "mock-etag")

        mutation_id = "mock-mutation-{}".format(int(time.time()))
        status = {"id": mutation_id, "status": "completed", "results": changes}
        return MutationOutcome(
            dry_run=False, success=True, mutation_id=mutation_id,
            plan=None, status_data=status, etag=etag or "mock-etag")

    def poll_mutation(self, mutation_id, max_attempts=12, interval_seconds=0.5):
        return {"id": mutation_id, "status": "completed"}

    def fetch_revision_etag(self, project_id, revision_id):
        return "mock-etag"

    def get_symbols(self, project_id):
        return []
