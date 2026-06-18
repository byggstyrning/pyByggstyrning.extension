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

from pyrevit import script

from cde import config
from cde.api import CDEClient, CDEApiError

logger = script.get_logger()

# --- Domain types ---------------------------------------------------------

Project = namedtuple("Project", ["id", "name"])
Revision = namedtuple("Revision", ["id", "name", "is_current"])
# value_type: "string" | "bool" | "number" | "enum"
ParameterDef = namedtuple(
    "ParameterDef", ["key", "label", "value_type", "allowed_values", "group"])
# values: {param_key: value}
CDEElement = namedtuple("CDEElement", ["global_id", "ifc_class", "name", "values"])

DEFAULT_IFC_CLASS = "IfcDoor"


def param_def(key, label, value_type="string", allowed_values=None, group=""):
    return ParameterDef(key, label, value_type, allowed_values or [], group)


class CDEService(object):
    """Live implementation backed by :class:`CDEClient`."""

    def __init__(self, auth):
        self.auth = auth
        self.client = CDEClient(auth)

    # --- projects / revisions ------------------------------------------

    def list_projects(self):
        """Return [Project]. Confirmed-ish: GET /api/v1/projects."""
        data = self.client.get(config.projects_url(self.client.base_url))
        items = data.get("projects", data) if isinstance(data, dict) else data
        projects = []
        for item in (items or []):
            pid = item.get("id") or item.get("project_id")
            name = item.get("name") or item.get("display_name") or str(pid)
            if pid:
                projects.append(Project(str(pid), name))
        return projects

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

    def list_elements(self, project_id, revision_id, ifc_class=DEFAULT_IFC_CLASS):
        """Return [CDEElement] of the given IfcClass for a project+revision.

        TODO(api): wire the confirmed element-list-by-IfcClass query. Shape the
        GraphQL below to the backend's real root field once known. Returns []
        on failure so the UI degrades gracefully (Revit-only matching still
        works).
        """
        query = (
            "query($projectId:String!,$revisionId:String!,$ifcClass:String!){"
            " elements(projectId:$projectId,revisionId:$revisionId,ifcClass:$ifcClass){"
            "   globalId ifcClass name"
            "   effectiveValues{ psetName propertyName value } } }")
        variables = {"projectId": project_id, "revisionId": revision_id,
                     "ifcClass": ifc_class}
        try:
            data = self.client.graphql(query, variables)
        except CDEApiError as ex:
            logger.warn("CDE: list_elements not available yet: {}".format(ex))
            return []
        return [self._element_from_node(node)
                for node in (data.get("elements") or [])]

    def get_element(self, project_id, revision_id, global_id):
        """Return a single CDEElement by GlobalId (confirmed GraphQL field)."""
        query = (
            "query($projectId:String!,$revisionId:String!,$globalId:String!){"
            " element(projectId:$projectId,revisionId:$revisionId,globalId:$globalId){"
            "   globalId ifcClass name"
            "   effectiveValues{ psetName propertyName value } } }")
        variables = {"projectId": project_id, "revisionId": revision_id,
                     "globalId": global_id}
        data = self.client.graphql(query, variables)
        node = data.get("element")
        return self._element_from_node(node) if node else None

    def _element_from_node(self, node):
        values = {}
        for ev in (node.get("effectiveValues") or []):
            name = ev.get("propertyName")
            if name is not None:
                values[name] = ev.get("value")
        return CDEElement(
            node.get("globalId"), node.get("ifcClass"), node.get("name"), values)

    # --- parameter definitions / values --------------------------------

    def get_parameter_defs(self, project_id, ifc_class=DEFAULT_IFC_CLASS):
        """Return [ParameterDef] for a category (the ~50 door schedule params).

        TODO(api): wire the Postgres-backed parameter-definition endpoint.
        Returns [] until available; the UI then only shows Revit-derived and
        GraphQL ``effectiveValues`` columns.
        """
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

    def set_element_values(self, project_id, revision_id, global_ids, values):
        """Write ``values`` (dict) across ``global_ids``. Returns True on success.

        TODO(api): wire the Postgres-backed mutation. Raises NotImplementedError
        so write-back is never silently dropped before the endpoint is wired.
        """
        raise NotImplementedError(
            "CDE set_element_values endpoint is not wired yet (TODO: backend mutation).")

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

    def set_element_values(self, project_id, revision_id, global_ids, values):
        logger.info("CDE(mock): set {} on {} elements".format(values, len(global_ids)))
        return True

    def get_symbols(self, project_id):
        return []
