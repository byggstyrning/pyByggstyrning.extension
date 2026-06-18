# -*- coding: utf-8 -*-
"""CDE integration package for pyRevit.

Connects Revit to the Byggstyrning CDE graph database (Neo4j + Postgres,
exposed via REST + GraphQL behind the CDE backend). Provides:

- ``config``        - base URL / endpoints / token paths (per-user overridable)
- ``auth``          - interactive Azure-OIDC login brokered by the CDE backend
- ``oidc_listener`` - loopback HttpListener used to capture the login redirect
- ``api``           - thin REST/GraphQL transport with Bearer auth
- ``service``       - domain-level API abstraction (+ a mock for offline UI dev)
- ``storage``       - Extensible Storage schema mapping a Revit model to a CDE
                      project + revision
- ``matching``      - join CDE elements (by IFC GlobalId) to Revit elements
- ``viewmodels``    - WPF-bindable row/grid models for the schedule UI
- ``coloring``      - temporary view graphics coloring by a CDE value
- ``revit_events``  - ExternalEvent handlers + active-view wiring

The tool is element/IfcClass-agnostic; doors (IfcDoor / OST_Doors) are the
first supported category.
"""
