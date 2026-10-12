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
    "completed", "succeeded", "committed", "failed", "rejected",
    "cancelled", "error"])
MUTATION_SUCCESS_STATES = frozenset(["completed", "succeeded", "committed"])

DEFAULT_IFC_CLASS = "IfcDoor"
# Door schedule: match both plain doors and standard-case doors in the graph.
DOOR_IFC_CLASSES = ("IFCDOOR", "IFCDOORSTANDARDCASE")
# GraphQL value roles merged onto each element (later roles win on conflict).
_VALUE_MERGE_ORDER = ("authoredValues", "derivedValues", "effectiveValues")
_VALUE_ROLE_PREFIX = {
    "authoredValues": "authored",
    "derivedValues": "derived",
    "effectiveValues": "effective",
}
# Relationship types that link a door to its DFP-carrying IfcGroup.
_DFP_RELATIONSHIP_TYPES = frozenset(["GROUPS", "ASSIGNED_TO"])
# GraphQL PropertyValue fields per contract.graphql
_PROPERTY_VALUE_FIELDS = (
    " psetName propertyName value datatype unit state sourceKind")
_RULE_TRACE_FIELDS = (
    " ruleId ruleVersion sourceAuthority validationState executedAt")
_RELATIONSHIP_FIELDS = (
    " relationshipRef family type state sourceKind confidence "
    "subjectGlobalId objectGlobalId direction constraints evidence")
# Node fields on element queries (elements + element) — Hub parity.
_ELEMENT_NODE_FIELDS = (
    " globalId ifcClass name etag"
    " authoredValues{" + _PROPERTY_VALUE_FIELDS + "}"
    " derivedValues{" + _PROPERTY_VALUE_FIELDS + "}"
    " effectiveValues{" + _PROPERTY_VALUE_FIELDS + "}"
    " ruleTrace{" + _RULE_TRACE_FIELDS + "}"
    " relationships{" + _RELATIONSHIP_FIELDS + "}")
# Lighter list pass: property values only (detail fetch uses full node).
_ELEMENT_LIST_NODE_FIELDS = (
    " globalId ifcClass name etag"
    " authoredValues{" + _PROPERTY_VALUE_FIELDS + "}"
    " derivedValues{" + _PROPERTY_VALUE_FIELDS + "}"
    " effectiveValues{" + _PROPERTY_VALUE_FIELDS + "}")


def _elements_graphql_query(include_filter, list_pass=False):
    """Build elements() query matching contract.graphql."""
    node_fields = _ELEMENT_LIST_NODE_FIELDS if list_pass else _ELEMENT_NODE_FIELDS
    if include_filter:
        return (
            "query($projectId: ID!, $revisionId: ID!, $first: Int, "
            "$after: String, $filter: ElementFilterInput) {"
            " elements(projectId: $projectId, revisionId: $revisionId, "
            "filter: $filter, first: $first, after: $after) {"
            " edges { cursor node {" + node_fields + " } }"
            " pageInfo { hasNextPage endCursor } } }")
    return (
        "query($projectId: ID!, $revisionId: ID!, $first: Int, $after: String) {"
        " elements(projectId: $projectId, revisionId: $revisionId, "
        "first: $first, after: $after) {"
        " edges { cursor node {" + node_fields + " } }"
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


def _property_value_key(pset, name):
    if pset:
        return "{}.{}".format(pset, name)
    return name


def _merge_property_values_from_node(node):
    """Merge all GraphQL value roles; flat keys use effective-wins merge order."""
    values = {}
    flat = {}
    for source in _VALUE_MERGE_ORDER:
        role = _VALUE_ROLE_PREFIX.get(source, source)
        for ev in (node.get(source) or []):
            pset = ev.get("psetName") or ""
            name = ev.get("propertyName")
            if name is None:
                continue
            key_flat = _property_value_key(pset, name)
            val = ev.get("value")
            if pset == DFP_PSET_NAME:
                val = coerce_dfp_cell_value(val)
            values["{}.{}".format(role, key_flat)] = val
            flat[key_flat] = val
    values.update(flat)
    return values


def _merge_rule_trace_into_values(values, node):
    for i, rt in enumerate(node.get("ruleTrace") or []):
        rid = rt.get("ruleId") or str(i)
        prefix = "RuleTrace.{}".format(rid)
        for field in ("ruleVersion", "sourceAuthority", "validationState", "executedAt"):
            val = rt.get(field)
            if val is not None and val != "":
                values["{}.{}".format(prefix, field)] = val


def _merge_relationships_into_values(values, node):
    for i, rel in enumerate(node.get("relationships") or []):
        rtype = rel.get("type") or "rel"
        prefix = "Relationship.{}.{}".format(i, rtype)
        for field in (
                "family", "type", "state", "sourceKind", "confidence",
                "subjectGlobalId", "objectGlobalId", "direction",
                "relationshipRef"):
            val = rel.get(field)
            if val is not None and val != "":
                values["{}.{}".format(prefix, field)] = val
        evidence = rel.get("evidence")
        if evidence is not None and evidence != "":
            if isinstance(evidence, dict):
                # ensure_ascii=False avoids json's py_encode_basestring_ascii,
                # which decodes byte-strings via the system code page and throws
                # on non-ASCII bytes (0xF6 'ö') under IronPython — the root cause
                # of "skip element ... 0xf6", read-back failures, and the
                # uncaught CLR crash (0xe0434352) during refresh.
                values["{}.evidence".format(prefix)] = json.dumps(
                    evidence, ensure_ascii=False)
            else:
                values["{}.evidence".format(prefix)] = evidence


def values_from_graph_node(node):
    """Build the full inspector value map from a GraphQL element node."""
    if not node:
        return {}
    values = _merge_property_values_from_node(node)
    _merge_rule_trace_into_values(values, node)
    _merge_relationships_into_values(values, node)
    return values


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


def _property_to_mutation_path(prop_key):
    """Map flat inspector key to mutation ``path`` = ``psetName.propertyName``.

    The backend parses ``path`` as ``<psetName>.<propertyName>`` (it splits on
    ``.`` and writes to authored values by default). A leading role qualifier
    (``authored.`` / ``derived.`` / ``effective.``) must be STRIPPED — if left
    in, the backend treats the role token (e.g. ``authored``) as the pset name,
    fails to resolve a real IFC pset, and performs no write
    (``governanceDecisions: dry_run_no_ifc_write``). Evidence: capture
    ``exchange_0004`` showed ``path='authored.Pset_DFP.Func_3_1'`` parsed to
    ``psetName='authored', propertyName='Func_3_1'`` with no IFC write.
    """
    for role in ("authored.", "derived.", "effective."):
        if prop_key.startswith(role):
            return prop_key[len(role):]
    return prop_key


def _build_mutation_operations(changes):
    """Build REST operations from {global_id: {property: value}}."""
    operations = []
    for gid in sorted(changes.keys()):
        props = changes.get(gid) or {}
        for prop in sorted(props.keys()):
            val = props[prop]
            path = _property_to_mutation_path(prop)
            if val is None or val == "":
                operations.append({
                    "op": "unset",
                    "path": path,
                })
            else:
                operations.append({
                    "op": "set",
                    "path": path,
                    "value": _mutation_value(val),
                })
    return operations


def _build_mutation_target(project_id, revision_id, etag=None,
                           file_version_id=None):
    """Build MutationTarget for POST /api/v2/mutations.

    The backend resolves the commit revision via ``_current_revision`` which
    matches ``model_id`` and/or ``revision_ref`` — NOT the GraphQL revision id.
    Its ``revision_ref`` is ``"fv:<file_version_id>"``. Sending the revision id
    as ``modelId`` (and an element etag as ``expectedRevision``) matched nothing,
    so the commit fell back to the newest revision project-wide and wrote to the
    WRONG revision while reads used the selected one. We therefore send
    ``expectedRevision = "fv:<file_version_id>"`` so project + revision_ref
    uniquely resolves to the same revision the client reads from. The per-element
    etag remains the ``If-Match`` header (optimistic lock), not the target.
    """
    target = {"projectId": project_id}
    if file_version_id:
        target["expectedRevision"] = "fv:{}".format(file_version_id)
    elif revision_id:
        # Legacy fallback for mappings saved before file_version_id was stored.
        # Re-run CDE Login to populate it so commits scope correctly.
        target["modelId"] = revision_id
    return target


def compute_revision_etag(project_id, file_version_id, file_id):
    """Reproduce the backend NgModelRevision ETag for If-Match.

    The backend derives it deterministically from identity parts (projection.py
    ``etag_for_parts``): ``sha256("<project>|<revision_ref>|<model_id>|<fileVer>")``
    where ``revision_ref == "fv:<file_version_id>"`` and ``model_id == file_id``.
    No read endpoint exposes this etag, so the client computes it. It is stable
    for a revision (pure identity, not content), so this is safe and exact.
    """
    if not (project_id and file_version_id and file_id):
        return ""
    revision_ref = "fv:{}".format(file_version_id)
    raw = "|".join([str(project_id), revision_ref, str(file_id),
                    str(file_version_id)])
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


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
        self.last_fetch_error = None
        self.last_fetch_pagination_complete = True
        self.last_revision_etag = ""
        # (project_id, revision_id) -> (file_version_id, file_id)
        self._mutation_identity_cache = {}

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

    def resolve_mutation_identity(self, project_id, revision_id):
        """Map live-graph revisionId -> (file_version_id, file_id).

        Mutations need ``revision_ref = fv:<file_version_id>`` and the
        backend model_id (ifc-versions ``file_id``) to scope commits and compute
        the revision If-Match etag. Login may not have stored these on older
        mappings, so resolve via ifc-versions + ingest-status at apply time.
        Returns ("", "") when not found.
        """
        if not (project_id and revision_id):
            return "", ""
        cache_key = (str(project_id), str(revision_id))
        cached = self._mutation_identity_cache.get(cache_key)
        if cached:
            return cached

        versions = self._fetch_ifc_versions(project_id)
        if not versions:
            return "", ""

        for item in versions:
            version_id = item.get("version_id")
            if not version_id:
                continue
            finalized = bool(
                item.get("finalized_at") or item.get("fragments_generated_at"))
            if not finalized:
                continue
            rev, synced = self._revision_for_version(version_id)
            if not synced or str(rev) != str(revision_id):
                continue
            file_id = item.get("file_id") or ""
            result = (str(version_id), str(file_id or ""))
            self._mutation_identity_cache[cache_key] = result
            return result

        return "", ""

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

    # Filtered door fetch: POST /api/v2/graphql with filter.ifcClasses
    # (contract.graphql). Page size 500 matches Hub nextgenGraph.allElements.
    DOOR_PAGE_SIZE = 500

    def list_elements(self, project_id, revision_id, ifc_class=DEFAULT_IFC_CLASS):
        """Return [CDEElement] of the given IfcClass for a project+revision.

        Live-graph reality (probed against the running backend):
        * ``elements`` is a Relay connection: ``{ edges{ node{...} } pageInfo }``.
        * Always paginates with ``filter.ifcClasses`` when a class is requested;
          client-side class guard + ``globalId`` dedupe as belt-and-suspenders.
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
        from cde.request_log import (
            mark_session, finish_session, get_log_path,
            get_get_index_path, get_trace_path, get_capture_dir)
        self.last_fetch_error = None
        self.last_fetch_nodes = 0
        self.last_fetch_mode = None
        if hasattr(self.auth, "reload_from_disk"):
            self.auth.reload_from_disk()
        if not self.auth.is_authenticated():
            self.last_fetch_error = (
                "Not signed in (or session expired). Run CDE Login, then Refresh.")
            logger.warn("CDE: list_elements skipped — {}".format(self.last_fetch_error))
            return []
        mark_session(
            "list_elements",
            project_id=project_id, revision_id=revision_id,
            ifc_class=ifc_class)
        logger.info(
            "CDE: HTTP capture index={} GETs={} trace={} dir={}".format(
                get_log_path(), get_get_index_path(),
                get_trace_path(), get_capture_dir()))
        want_classes = want_ifc_classes(ifc_class)
        ifc_classes = self._ifc_classes_for_filter(ifc_class)
        elements = []
        try:
            nodes, pages = self._paginate_element_nodes(
                project_id, revision_id,
                ifc_classes=ifc_classes,
                want_classes=want_classes,
                page_size=self.DOOR_PAGE_SIZE)
            self.last_fetch_pagination_complete = getattr(
                self, "_last_pagination_complete", True)
            self.last_fetch_mode = (
                "ifc_class_filter" if ifc_classes else "unfiltered")
            self.last_fetch_pages = pages
            self.last_fetch_nodes = len(nodes)
            self.last_truncation = self._truncation(
                project_id, revision_id, len(nodes),
                filtered=bool(ifc_classes),
                pagination_complete=self.last_fetch_pagination_complete,
                ifc_class=ifc_class)
            node_by_gid = {
                n.get("globalId"): n for n in nodes if n.get("globalId")}
            self.last_node_by_gid = node_by_gid
            group_dfp_merged = 0
            skipped_nodes = 0
            for node in nodes:
                if not getattr(self, "last_revision_etag", None):
                    node_etag = node.get("etag")
                    if node_etag:
                        self.last_revision_etag = str(node_etag)
                if want_classes and not node_matches_ifc_classes(node, want_classes):
                    continue
                try:
                    el = self._element_from_node(node)
                    merged = dict(el.values)
                    if self._inherit_group_values(merged, node, node_by_gid):
                        group_dfp_merged += 1
                        el = el._replace(values=merged)
                    elements.append(el)
                except Exception as ex:
                    skipped_nodes += 1
                    logger.warn("CDE: skip element {}: {}".format(
                        node.get("globalId"), ex))
            if skipped_nodes:
                logger.warn("CDE: skipped {} element(s) during value merge".format(
                    skipped_nodes))
            logger.info("CDE: list_elements {} of {} nodes match {}".format(
                len(elements), len(nodes), ifc_class))
            if elements:
                # Key sampling must never crash the refresh: a byte-string key
                # (e.g. a pset name carrying a raw 0xF6 'ö') compared against a
                # unicode key triggers an implicit .NET code-page decode that, if
                # it escaped this try/finally (no except), surfaces as an
                # unhandled CLR exception (0xe0434352) and terminates Revit.
                try:
                    el0 = elements[0]
                    all_keys = sorted(el0.values.keys())
                    pset_prefixes = sorted(set(
                        k.split(".", 1)[0] for k in all_keys if "." in k))
                    logger.info(
                        "CDE: sample door keys={} psets={}".format(
                            len(all_keys), pset_prefixes[:10]))
                except Exception as _ex:
                    logger.warn("CDE: key sampling skipped: {}".format(_ex))
        finally:
            finish_session(
                "list_elements",
                element_count=len(elements),
                raw_nodes=self.last_fetch_nodes,
                fetch_mode=self.last_fetch_mode,
                fetch_error=self.last_fetch_error)
        return elements

    MAX_ELEMENT_PAGES = 15

    def _ifc_classes_for_filter(self, ifc_class):
        """GraphQL ``filter.ifcClasses`` values for a schedule category."""
        want_classes = want_ifc_classes(ifc_class)
        if not want_classes:
            return None
        if want_classes == want_ifc_classes(DEFAULT_IFC_CLASS):
            return list(DOOR_IFC_CLASSES)
        return list(want_classes)

    def _paginate_element_nodes(self, project_id, revision_id,
                                ifc_classes=None, want_classes=None,
                                page_size=500, max_pages=None):
        """Single paginated ``elements`` loop.

        When ``ifc_classes`` is set, every request uses the filtered GraphQL
        query and passes ``filter: { ifcClasses: [...] }``. ``after`` always
        advances from ``pageInfo.endCursor`` — never reset mid-loop.
        """
        page_limit = max_pages if max_pages is not None else self.MAX_ELEMENT_PAGES
        filter_sent = bool(ifc_classes)
        query = _elements_graphql_query(filter_sent, list_pass=True)
        seen_gids = set()
        nodes = []
        after = None
        pages_fetched = 0
        last_has_next = False

        while pages_fetched < page_limit:
            pages_fetched += 1
            variables = {
                "projectId": project_id,
                "revisionId": revision_id,
                "first": page_size,
                "after": after,
            }
            if filter_sent:
                variables["filter"] = {"ifcClasses": list(ifc_classes)}

            try:
                data = self.client.graphql(query, variables)
            except CDEApiError as ex:
                self.last_fetch_error = str(ex)
                logger.warn("CDE: list_elements failed: {}".format(ex))
                break

            conn = data.get("elements") or {}
            edges = conn.get("edges") or []
            page_info = conn.get("pageInfo") or {}

            if not edges:
                break

            for edge in edges:
                node = (edge or {}).get("node") or {}
                gid = node.get("globalId")
                if not gid or gid in seen_gids:
                    continue
                seen_gids.add(gid)
                if want_classes and not node_matches_ifc_classes(node, want_classes):
                    if filter_sent:
                        logger.warn(
                            "CDE: non-door row despite ifcClasses filter: {}".format(
                                node.get("ifcClass")))
                    continue
                nodes.append(node)

            if not page_info.get("hasNextPage"):
                last_has_next = False
                break
            next_after = page_info.get("endCursor")
            if not next_after or next_after == after:
                last_has_next = False
                break
            last_has_next = True
            after = next_after

        self._last_pagination_complete = not last_has_next
        logger.info("CDE: list_elements rev={} pages={} nodes={} complete={}".format(
            revision_id, pages_fetched, len(nodes), self._last_pagination_complete))
        return nodes, pages_fetched

    def _truncation(self, project_id, revision_id, retrieved,
                    filtered=False, pagination_complete=True, ifc_class=None):
        """Detect incomplete fetches (pagination cap), not filtered-vs-total mismatch.

        ``postgres.elements`` is the whole-revision element count (~7808). Comparing
        a door-filtered fetch (171) against that total is wrong — skip when filtered
        and pagination reached the natural end.
        """
        if filtered and pagination_complete:
            return None
        if filtered and not pagination_complete:
            logger.warn(
                "CDE: filtered {} fetch stopped early — retrieved {} (pagination cap)".format(
                    ifc_class or "elements", retrieved))
            return {
                "retrieved": retrieved,
                "total": None,
                "ifc_class": ifc_class,
                "reason": "pagination_cap",
            }
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
            return {"retrieved": retrieved, "total": total, "reason": "graph_total"}
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

    def graph_node_values(self, node):
        """Full value map (all pset roles, ruleTrace, relationships) from a node."""
        return values_from_graph_node(node)

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
        """Merge authored, derived, effective, ruleTrace, and relationships."""
        return values_from_graph_node(node)

    def _element_from_node(self, node):
        if not node:
            return CDEElement(None, None, None, {})
        return CDEElement(
            node.get("globalId"),
            node.get("ifcClass"),
            node.get("name"),
            values_from_graph_node(node))

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
        skip_prefixes = ("authored.", "derived.", "effective.", "RuleTrace.", "Relationship.")
        for group_gid in self._group_global_ids_for_door(door_node):
            group_node = node_by_gid.get(group_gid)
            if not group_node:
                continue
            for key, val in self._merge_values_from_node(group_node).items():
                if key.startswith(skip_prefixes):
                    continue
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
                                dry_run=True, etag=None, idempotency_key=None,
                                file_version_id=None, file_id=None):
        """POST /api/v2/mutations (transactional lane).

        ``changes`` is ``{global_id: {property_key: value}}``. When
        ``dry_run`` is True (default) the backend returns a DryRunPlan; commit
        with ``dry_run=False`` and a fresh ``If-Match`` etag. ``file_version_id``
        scopes the commit to the SAME revision reads use (revision_ref);
        ``file_id`` lets us compute the exact revision ETag for If-Match.
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

        if not file_version_id or not file_id:
            resolved_vid, resolved_fid = self.resolve_mutation_identity(
                project_id, revision_id)
            if not file_version_id:
                file_version_id = resolved_vid or None
            if not file_id:
                file_id = resolved_fid or None

        target = _build_mutation_target(
            project_id, revision_id, etag=etag, file_version_id=file_version_id)
        payload = {
            "lane": "transactional",
            "target": target,
            "objectIds": object_ids,
            "operations": operations,
            "dryRun": bool(dry_run),
        }
        headers = {
            "Idempotency-Key": idempotency_key or _idempotency_key(
                revision_id, object_ids, operations, dry_run),
        }
        # The backend's If-Match must equal NgModelRevision.etag, which it derives
        # from identity parts. Stale cached/element etags cause 412 etag_mismatch,
        # so when we have the identity parts compute the authoritative etag and
        # prefer it over any passed/cached value.
        computed_etag = compute_revision_etag(
            project_id, file_version_id, file_id)
        effective_etag = computed_etag or etag
        if not effective_etag:
            effective_etag = self.fetch_revision_etag(project_id, revision_id)
        if not effective_etag:
            raise CDEApiError(
                "Cannot mutate without revision etag (If-Match required).")
        headers["If-Match"] = effective_etag

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

    def _fetch_revision_etag_graphql(self, project_id, revision_id):
        """Fallback: read ``etag`` from the first element node in the revision."""
        query = (
            "query($projectId: ID!, $revisionId: ID!, $first: Int) {"
            " elements(projectId: $projectId, revisionId: $revisionId, "
            "first: $first) {"
            " edges { node { etag } } } }")
        try:
            data = self.client.graphql(
                query,
                {
                    "projectId": project_id,
                    "revisionId": revision_id,
                    "first": 1,
                })
        except CDEApiError as ex:
            logger.debug("CDE: graphql etag lookup failed: {}".format(ex))
            return ""
        for edge in ((data.get("elements") or {}).get("edges") or []):
            node = (edge or {}).get("node") or {}
            etag = node.get("etag")
            if etag:
                return str(etag)
        return ""

    def fetch_revision_etag(self, project_id, revision_id):
        """Return NgModelRevision.etag for If-Match optimistic locking."""
        cached = getattr(self, "last_revision_etag", None)
        if cached:
            return cached

        try:
            data = self.client.get(
                config.graph_status_url(self.client.base_url),
                params={"projectId": project_id, "revisionId": revision_id})
            etag = _extract_etag(data)
            if etag:
                self.last_revision_etag = etag
                return etag
        except CDEApiError as ex:
            logger.debug("CDE: graph status etag lookup failed: {}".format(ex))

        try:
            data = self.client.get(
                config.graph_overview_url(self.client.base_url),
                params={"projectId": project_id, "revisionId": revision_id})
            etag = _extract_etag(data)
            if etag:
                self.last_revision_etag = etag
                return etag
            for item in (data.get("revisions") or []):
                rid = item.get("id") or item.get("revisionId")
                if str(rid) == str(revision_id):
                    etag = _extract_etag(item)
                    if etag:
                        self.last_revision_etag = etag
                        return etag
        except CDEApiError as ex:
            logger.debug("CDE: graph overview etag lookup failed: {}".format(ex))

        etag = self._fetch_revision_etag_graphql(project_id, revision_id)
        if etag:
            self.last_revision_etag = etag
        return etag

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
                                dry_run=True, etag=None, idempotency_key=None,
                                file_version_id=None, file_id=None):
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
