# -*- coding: utf-8 -*-
"""Sign in to the CDE and map this Revit model to a CDE project + model.

Runs the interactive Azure-OIDC redirect login (brokered by the CDE backend),
lists the user's projects/models and persists the chosen mapping into the
model via Extensible Storage. An offline "sample data" mode lets the flow be
exercised without a live backend.
"""
__title__ = "CDE\nLogin"
__author__ = "Byggstyrning AB"
__doc__ = "Sign in to the CDE and map this model to a CDE project + model."

import json
import os.path as op
import sys
import time

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Action
from System.Threading import Thread, ThreadStart

from pyrevit import forms, revit, script

# --- lib bootstrap --------------------------------------------------------
_pushbutton_dir = op.dirname(__file__)
_extension_dir = op.dirname(op.dirname(op.dirname(_pushbutton_dir)))
_lib_path = op.join(_extension_dir, "lib")
if _lib_path not in sys.path:
    sys.path.insert(0, _lib_path)

from styles import load_styles_to_window
from cde import config as cde_config
import cde.service as cde_service
reload(cde_config)
reload(cde_service)
from cde import storage
from cde.auth import CDEAuthClient
from cde.service import CDEService, MockCDEService

logger = script.get_logger()
doc = revit.doc

class _ComboItem(forms.Reactive):
    """ComboBox row; ItemTemplate binds DisplayName (IronPython-safe)."""

    def __init__(self, display_name, value):
        super(_ComboItem, self).__init__()
        self._display_name = display_name or ""
        self.value = value

    @property
    def DisplayName(self):
        return self._display_name

    def ToString(self):
        return self._display_name


class LoginWindow(forms.WPFWindow):
    """Login + project mapping dialog."""

    def __init__(self):
        forms.WPFWindow.__init__(self, op.join(_pushbutton_dir, "LoginWindow.xaml"))
        load_styles_to_window(self)

        self.auth = CDEAuthClient()
        self.offline = False
        self._closing = False
        self.service = self._make_service()

        self.Closing += self.window_closing

        self._refresh_account()
        self._refresh_current_mapping()
        if self.auth.is_authenticated():
            self._load_projects_async()

    # --- service selection ---------------------------------------------

    def _make_service(self):
        return MockCDEService(self.auth) if self.offline else CDEService(self.auth)

    # --- async helper ---------------------------------------------------

    def _run_async(self, work, on_done=None):
        """Run ``work()`` on a background thread; ``on_done(result, error)`` on UI."""
        def runner():
            result = None
            error = None
            try:
                result = work()
            except Exception as ex:
                error = ex
            if on_done is not None:
                self.Dispatcher.Invoke(Action(lambda: on_done(result, error)))
        thread = Thread(ThreadStart(runner))
        thread.IsBackground = True
        thread.Start()

    def _set_status(self, message):
        self.statusText.Text = message

    # --- account --------------------------------------------------------

    def _refresh_account(self):
        if self.offline:
            self.accountText.Text = "Offline demo mode (no sign-in needed)."
            return
        if self.auth.is_authenticated():
            email = self.auth.user_email or "signed in"
            self.accountText.Text = "Signed in as {}.".format(email)
        else:
            self.accountText.Text = "Not signed in."

    def on_login_click(self, sender, args):
        if self.offline:
            self._set_status("Offline mode active.")
            return
        self.loginButton.IsEnabled = False
        self._set_status("Opening browser - complete sign-in there...")

        def work():
            return self.auth.login_interactive()

        def done(success, error):
            if self._closing:
                return
            self.loginButton.IsEnabled = True
            if error is not None:
                self._set_status("Login error: {}".format(error))
                return
            if success:
                self._set_status("Signed in.")
                self._refresh_account()
                self._load_projects_async()
            else:
                self._set_status(self.auth.last_error or "Login failed.")

        self._run_async(work, done)

    def on_signout_click(self, sender, args):
        self.auth.logout()
        self.projectCombo.ItemsSource = None
        self.modelCombo.ItemsSource = None
        self._refresh_account()
        self._set_status("Signed out.")

    def on_offline_toggled(self, sender, args):
        self.offline = bool(self.offlineCheck.IsChecked)
        self.service = self._make_service()
        self._refresh_account()
        if self.offline or self.auth.is_authenticated():
            self._load_projects_async()
        else:
            self.projectCombo.ItemsSource = None
            self.modelCombo.ItemsSource = None

    # --- projects / models ---------------------------------------------

    def _load_projects_async(self):
        self._set_status("Loading projects...")

        def work():
            return self.service.list_projects()

        def done(projects, error):
            if error is not None:
                self._set_status("Could not load projects: {}".format(error))
                return
            items = [_ComboItem(p.name, p) for p in (projects or [])]
            self.projectCombo.ItemsSource = items
            self._set_status("Loaded {} active project(s).".format(len(items)))
            self._preselect_mapping()

        self._run_async(work, done)

    def on_project_changed(self, sender, args):
        item = self.projectCombo.SelectedItem
        if item is None:
            return
        project = item.value
        self._set_status("Loading models...")

        def work():
            return self.service.list_models(project.id)

        def done(models, error):
            if error is not None:
                self._set_status("Could not load models: {}".format(error))
                return
            items = [_ComboItem(m.name, m) for m in (models or [])]
            self.modelCombo.ItemsSource = items
            projected = [m for m in (models or []) if m.is_projected]
            if projected:
                for i, it in enumerate(items):
                    if it.value.is_projected:
                        self.modelCombo.SelectedIndex = i
                        break
            self._preselect_model(mapping_revision_id=(
                storage.load_mapping(doc) or {}).get("revision_id"))
            self._set_status("Loaded {} model(s), {} projected.".format(
                len(items), len(projected)))

        self._run_async(work, done)

    def _preselect_model(self, mapping_revision_id=None):
        if not mapping_revision_id:
            return
        for i in range(self.modelCombo.Items.Count):
            item = self.modelCombo.Items[i]
            if item.value.id == mapping_revision_id:
                self.modelCombo.SelectedIndex = i
                break

    def _preselect_mapping(self):
        mapping = storage.load_mapping(doc)
        if not mapping:
            return
        for i in range(self.projectCombo.Items.Count):
            item = self.projectCombo.Items[i]
            if item.value.id == mapping["project_id"]:
                self.projectCombo.SelectedIndex = i
                break

    # --- mapping --------------------------------------------------------

    def _refresh_current_mapping(self):
        mapping = storage.load_mapping(doc)
        if mapping:
            self.currentMappingText.Text = "Mapped to '{}' (model {}).".format(
                mapping["project_name"] or mapping["project_id"],
                mapping["revision_id"] or "n/a")
        else:
            self.currentMappingText.Text = "This model is not mapped to a CDE project."

    def on_map_click(self, sender, args):
        project_item = self.projectCombo.SelectedItem
        if project_item is None:
            self._set_status("Select a project first.")
            return
        project = project_item.value
        model_item = self.modelCombo.SelectedItem
        if model_item is None:
            self._set_status("Select a model first.")
            return
        model = model_item.value
        if not model.id:
            self._set_status(
                "Selected model is not projected yet (no revision id). "
                "Wait for CDE ingest to finish.")
            return
        revision_id = model.id

        ok = storage.save_mapping(
            doc, project.id, project.name, revision_id, cde_config.get_base_url())
        if ok:
            self._refresh_current_mapping()
            self._set_status("Model mapped to '{}' / {}.".format(
                project.name, model.name))
        else:
            self._set_status("Failed to save mapping (see log).")

    # --- lifecycle ------------------------------------------------------

    def window_closing(self, sender, args):
        self._closing = True
        try:
            self.auth.cancel_login()
        except Exception as ex:
            logger.debug("CDE: error cancelling login on close: {}".format(ex))


if __name__ == "__main__":
    LoginWindow().ShowDialog()
