# -*- coding: utf-8 -*-
__title__ = "CDE\nPanel"
__author__ = "Byggstyrning AB"
__doc__ = """Toggle the CDE panel.

A dockable panel showing this document's saved CDE import configurations
(current connection: StreamBIM). Run them individually or all at once
while you keep working - Revit stays interactive during the import.

The StreamBIM checklist tools themselves stay as regular windows; the
panel launches them for sign-in and configuration setup.

The panel is registered during Revit startup; this button only shows/hides
it. After installing or updating the extension, a FULL Revit restart is
required once - Revit does not allow dockable panes to be registered
mid-session, so a pyRevit reload alone is not enough."""

from pyrevit import forms, script

logger = script.get_logger()

# Panel GUID - must match streambim.panel_ui.CDE_PANEL_ID
CDE_PANEL_ID = "3d7a6f6e-2c9b-4f8e-9a41-c5e8b1c9d0a2"

try:
    pane = forms.get_dockable_panel(CDE_PANEL_ID)
    if pane.IsShown():
        pane.Hide()
    else:
        pane.Show()
except Exception as e:
    logger.error("CDE panel is not registered yet: {}".format(str(e)))
    forms.alert(
        "The CDE panel is not available yet.\n\n"
        "Dockable panels can only be registered while Revit is starting "
        "up, so after installing or updating this extension you need to "
        "close and restart Revit completely once.\n\n"
        "(A pyRevit reload is not enough - Revit rejects panel "
        "registration mid-session.)",
        title="CDE Panel")
