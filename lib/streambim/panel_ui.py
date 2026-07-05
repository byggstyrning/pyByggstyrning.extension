# -*- coding: utf-8 -*-
"""Dockable CDE panel.

The always-available home for the recurring Common Data Environment
workflow: see the active document's saved import configurations, re-run
them with one click while Revit stays interactive, and watch progress live.

StreamBIM is the CDE connection this panel currently drives; the StreamBIM
checklist tools themselves (Checklist Importer, Edit Configs, Run
Everything) remain ordinary windows and are launched from here for setup.

Architecture (modeless rules):
* HTTP (StreamBIM API) runs on a background thread - never on the UI thread.
* Revit API reads/writes run inside pyRevit's generic ExternalEvent via
  pyrevit.revit.events.execute_in_revit_context().
* This module MUST be first imported from startup.py (Revit API context):
  both the panel registration and the module-level ExternalEvent used by
  execute_in_revit_context require it.
"""

import sys
import threading
import traceback
import os.path as op

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System import Action, EventHandler
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Visibility
from System.Windows.Controls import ListBoxItem
from System.Windows.Input import MouseButton
from System.Windows.Media import Visual, VisualTreeHelper
from Autodesk.Revit.UI.Events import ViewActivatedEventArgs

from pyrevit import HOST_APP, forms, script
# NOTE: imported at module level on purpose - creates its ExternalEvent
# while we are still in Revit API context (startup).
from pyrevit.revit import events as revit_events

# lib path bootstrap (this module lives in lib/streambim)
_current_dir = op.dirname(__file__)
_lib_dir = op.dirname(_current_dir)
_extension_dir = op.dirname(_lib_dir)
if _lib_dir not in sys.path:
    sys.path.append(_lib_dir)

from styles import load_styles_to_window
from streambim import streambim_api
from streambim.run_engine import (
    ChecklistMetadataCache,
    config_from_dict,
    fetch_config_values,
    build_element_lookup,
    apply_config_values,
)

logger = script.get_logger()

CDE_PANEL_ID = "3d7a6f6e-2c9b-4f8e-9a41-c5e8b1c9d0a2"
# Backwards-compatible alias
STREAMBIM_PANEL_ID = CDE_PANEL_ID

_CHECKLIST_IMPORTER_BUNDLE = op.join(
    _extension_dir, 'pyBS.tab', 'StreamBIM.panel', 'ChecklistImporter.pushbutton')
_CONFIG_EDITOR_BUNDLE = op.join(
    _extension_dir, 'pyBS.tab', 'StreamBIM.panel', 'Edit.stack', 'Edit Configs.pushbutton')


class ConfigCard(object):
    """Display model for one saved configuration."""

    def __init__(self, config):
        self.config = config
        self.last_result = None
        self.Title = config.ChecklistName
        self.Detail = u"{} → {}".format(
            config.streambim_property or '?', config.revit_parameter or '?')
        self.Meta = self._build_meta()

    def _build_meta(self):
        parts = []
        if self.config.mapping_enabled:
            parts.append("{} value mappings".format(self.config.mapping_count))
        else:
            parts.append("direct values")
        if self.last_result is not None:
            processed, updated = self.last_result
            parts.append(u"✓ {} of {} updated".format(updated, processed))
        return u" · ".join(parts)

    def set_result(self, processed, updated):
        self.last_result = (processed, updated)
        self.Meta = self._build_meta()


class CDEPanel(forms.WPFPanel):
    """Dockable CDE panel: run and monitor CDE imports while working.

    Current connection: StreamBIM.
    """

    panel_id = CDE_PANEL_ID
    panel_source = op.join(_current_dir, 'CDEPanel.xaml')
    panel_title = "CDE Panel"

    def __init__(self):
        forms.WPFPanel.__init__(self)
        load_styles_to_window(self)

        self._cards = []
        self._project_id = None
        self._doc = None
        self._doc_title = None
        self._is_running = False

        self._wire_ui()
        self._refresh_auth_ui()
        self._render_cards()
        self._set_status("Open a document to load its configurations.")

        # Revit document tracking: reload configs whenever a view (and thus
        # possibly another document) is activated. Runs in API context.
        try:
            HOST_APP.uiapp.ViewActivated += \
                EventHandler[ViewActivatedEventArgs](self._on_view_activated)
        except Exception as e:
            logger.warning("CDE panel: could not subscribe ViewActivated: {}"
                           .format(str(e)))

    # ------------------------------------------------------------------ UI

    def _wire_ui(self):
        self.refreshButton.Click += self._on_refresh_click
        self.signInButton.Click += self._on_new_config_click
        self.runSelectedButton.Click += self._on_run_selected_click
        self.runAllButton.Click += self._on_run_all_click
        self.newConfigButton.Click += self._on_new_config_click
        self.editConfigsButton.Click += self._on_edit_configs_click
        self.configsList.SelectionChanged += self._on_selection_changed
        self.configsList.MouseDoubleClick += self._on_list_double_click

    def _dispatch(self, func):
        """Run func on the panel's UI thread."""
        try:
            self.Dispatcher.BeginInvoke(Action(func))
        except Exception as e:
            logger.error("CDE panel dispatch error: {}".format(str(e)))

    def _set_status(self, message):
        self.statusText.Text = message or ""

    def _set_running(self, running):
        self._is_running = running
        self.progressPanel.Visibility = \
            Visibility.Visible if running else Visibility.Collapsed
        self.runAllButton.IsEnabled = (not running) and bool(self._cards)
        self.runSelectedButton.IsEnabled = \
            (not running) and self.configsList.SelectedItem is not None
        self.newConfigButton.IsEnabled = not running
        self.editConfigsButton.IsEnabled = not running
        self.refreshButton.IsEnabled = not running

    def _refresh_auth_ui(self):
        """Reflect cached sign-in state (token file read only - no network)."""
        try:
            client = streambim_api.StreamBIMClient()
            if client.idToken:
                self.accountText.Text = u"StreamBIM · {}".format(
                    client.username or "signed in")
                self.accountDot.Fill = self.Resources['SuccessBrush']
                self.signInButton.Visibility = Visibility.Collapsed
            else:
                self.accountText.Text = u"StreamBIM · not signed in"
                self.accountDot.Fill = self.Resources['TextLightBrush']
                self.signInButton.Visibility = Visibility.Visible
        except Exception as e:
            logger.error("CDE panel auth refresh error: {}".format(str(e)))

    def _render_cards(self):
        collection = ObservableCollection[object]()
        for card in self._cards:
            collection.Add(card)
        self.configsList.ItemsSource = collection
        self.configCountText.Text = str(len(self._cards))
        self.emptyState.Visibility = \
            Visibility.Collapsed if self._cards else Visibility.Visible
        self.runAllButton.IsEnabled = bool(self._cards) and not self._is_running
        self.runSelectedButton.IsEnabled = False

        context_parts = []
        if self._doc_title:
            context_parts.append(self._doc_title)
        if self._project_id:
            context_parts.append("project {}".format(self._project_id))
        self.projectText.Text = u" · ".join(context_parts)

    # ------------------------------------------------ document tracking

    def _on_view_activated(self, sender, args):
        """Revit event (API context): reload configs when the document changes."""
        try:
            doc = args.Document
            if doc is None or getattr(doc, 'IsFamilyDocument', False):
                return
            if self._is_running:
                return  # never rebuild cards mid-run; _finish_run re-renders
            same_doc = False
            try:
                same_doc = self._doc is not None and doc.Equals(self._doc)
            except Exception:
                same_doc = False  # cached doc closed/invalid -> treat as changed
            if same_doc:
                return  # same document: keep cards, selection and results
            self._load_configs_from_doc(doc)
        except Exception as e:
            logger.error("CDE panel ViewActivated error: {}".format(str(e)))

    def _load_configs_from_doc(self, doc):
        """Read-only config load. Must run in Revit API context."""
        try:
            config_dicts = streambim_api.load_configs_readonly(doc)
            self._cards = [ConfigCard(config_from_dict(d)) for d in config_dicts]
            self._project_id = streambim_api.get_saved_project_id(doc)
            self._doc = doc
            self._doc_title = doc.Title
            self._render_cards()
            if not self._is_running:
                if self._cards:
                    self._set_status("Ready. Select a configuration and run it, "
                                     "or run all.")
                else:
                    self._set_status("")
        except Exception as e:
            logger.error("CDE panel config load error: {}".format(str(e)))

    def _reload_active_doc_in_context(self):
        """Scheduled through ExternalEvent by the refresh button."""
        uidoc = HOST_APP.uiapp.ActiveUIDocument
        if uidoc and uidoc.Document:
            self._load_configs_from_doc(uidoc.Document)
        else:
            self._doc = None
            self._doc_title = None
            self._project_id = None
            self._cards = []
            self._render_cards()
            self._set_status("Open a document to load its configurations.")

    # ------------------------------------------------------------ actions

    def _on_selection_changed(self, sender, args):
        self.runSelectedButton.IsEnabled = \
            (not self._is_running) and self.configsList.SelectedItem is not None

    def _on_refresh_click(self, sender, args):
        self._refresh_auth_ui()
        revit_events.execute_in_revit_context(self._reload_active_doc_in_context)

    def _on_new_config_click(self, sender, args):
        self._launch_bundle(_CHECKLIST_IMPORTER_BUNDLE, "Checklist Importer")

    def _on_edit_configs_click(self, sender, args):
        self._launch_bundle(_CONFIG_EDITOR_BUNDLE, "Edit Configs")

    def _launch_bundle(self, bundle_path, display_name):
        """Execute an existing pyRevit command bundle exactly as a button click."""
        def _do_launch():
            launch_failed = False
            try:
                from pyrevit.extensions.genericcomps import GenericUIComponent
                from pyrevit.loader import sessionmgr
                unique_name = GenericUIComponent.make_unique_name(bundle_path)
                # execute_command returns None (without raising) when the
                # command cannot be found - detect that explicitly.
                cmd_class = sessionmgr.find_pyrevitcmd(unique_name)
                if cmd_class:
                    sessionmgr.execute_command_cls(cmd_class)
                else:
                    launch_failed = True
                    logger.error("CDE panel: command not found: {}".format(unique_name))
            except Exception as e:
                launch_failed = True
                logger.error("CDE panel: could not launch {}: {}".format(
                    display_name, str(e)))
            # Whatever happened in the dialog, reflect the latest state.
            self._refresh_auth_ui()
            self._reload_active_doc_in_context()
            if launch_failed:
                self._dispatch(lambda: self._set_status(
                    "Could not open {} from the panel - use its ribbon button."
                    .format(display_name)))

        self._set_status("Opening {}...".format(display_name))
        revit_events.execute_in_revit_context(_do_launch)

    def _on_run_selected_click(self, sender, args):
        card = self.configsList.SelectedItem
        if card is not None:
            self._start_run([card])

    def _on_list_double_click(self, sender, args):
        # MouseDoubleClick also fires for right-clicks and for double-clicks
        # on the scrollbar/empty area; only run when a left double-click
        # actually landed on a config card.
        if args.ChangedButton != MouseButton.Left:
            return
        source = args.OriginalSource
        while source is not None and not isinstance(source, ListBoxItem):
            source = (VisualTreeHelper.GetParent(source)
                      if isinstance(source, Visual) else None)
        if source is None:
            return
        self._on_run_selected_click(sender, args)

    def _on_run_all_click(self, sender, args):
        self._start_run(list(self._cards))

    # ---------------------------------------------------------- run flow

    def _start_run(self, cards):
        if self._is_running or not cards:
            return

        client = streambim_api.StreamBIMClient()
        if not client.idToken:
            self._set_status("Not signed in - use the \"Sign in...\" link above "
                             "(it opens the Checklist Importer's sign-in form).")
            return
        if not self._project_id:
            self._set_status("No StreamBIM project is linked to this document "
                             "yet - create a configuration first.")
            return
        run_doc = self._doc
        if run_doc is None:
            self._set_status("No document context - click refresh and try again.")
            return
        client.set_current_project(self._project_id)

        self._set_running(True)
        self._set_status("Fetching checklist data...")

        worker = threading.Thread(target=self._fetch_worker,
                                  args=(client, cards, run_doc))
        worker.daemon = True
        worker.start()

    def _fetch_worker(self, client, cards, run_doc):
        """Background thread: HTTP only - no Revit API access here."""
        results = []
        errors = []
        try:
            metadata_cache = ChecklistMetadataCache(client)
            for i, card in enumerate(cards):
                message = "Fetching {}/{}: {}".format(i + 1, len(cards), card.Title)
                self._dispatch(lambda m=message: self._set_status(m))

                guid_to_value = fetch_config_values(
                    client, card.config, metadata_cache,
                    status_cb=lambda m: self._dispatch(
                        lambda mm=m: self._set_status(mm)))

                if guid_to_value is None:
                    errors.append("{}: {}".format(
                        card.Title, client.last_error or "fetch failed"))
                else:
                    results.append((card, guid_to_value))
        except Exception as e:
            logger.error("CDE panel fetch error: {}\n{}".format(
                str(e), traceback.format_exc()))
            errors.append(str(e))

        def _continue_on_ui():
            if not results:
                self._set_running(False)
                self._set_status("Nothing to apply. " + ("; ".join(errors) if errors
                                                         else "No values matched."))
                return
            self._set_status("Applying values to elements...")
            try:
                revit_events.execute_in_revit_context(
                    self._apply_in_context, results, errors, run_doc)
            except Exception as e:
                # Scheduling failed: recover the UI instead of hanging.
                logger.error("CDE panel: could not schedule apply: {}".format(str(e)))
                self._finish_run(0, 0, list(errors) + [str(e)])

        self._dispatch(_continue_on_ui)

    def _apply_in_context(self, results, errors, run_doc):
        """ExternalEvent (API context): write values inside transactions."""
        total_processed = 0
        total_updated = 0
        try:
            uidoc = HOST_APP.uiapp.ActiveUIDocument
            doc = uidoc.Document if uidoc else None
            # Guard against a document switch during the (potentially long)
            # fetch: values fetched for one model must never be written into
            # another one that happens to be active when the event fires.
            doc_matches = False
            try:
                doc_matches = (doc is not None and run_doc is not None
                               and run_doc.IsValidObject and doc.Equals(run_doc))
            except Exception:
                doc_matches = False
            if not doc_matches:
                self._dispatch(lambda: self._finish_run(
                    0, 0, list(errors) + [
                        "Active document changed during the run - no values "
                        "were applied. Reactivate the original document and "
                        "run again."]))
                return

            # One model scan for all configs in this run.
            all_guids = set()
            for _, guid_to_value in results:
                all_guids.update(guid_to_value.keys())
            element_lookup = build_element_lookup(doc, all_guids)

            for card, guid_to_value in results:
                processed, updated = apply_config_values(
                    doc, card.config, guid_to_value, element_lookup)
                card.set_result(processed, updated)
                total_processed += processed
                total_updated += updated
        except Exception as e:
            logger.error("CDE panel apply error: {}\n{}".format(
                str(e), traceback.format_exc()))
            errors = list(errors) + [str(e)]

        self._dispatch(lambda: self._finish_run(
            total_processed, total_updated, errors))

    def _finish_run(self, total_processed, total_updated, errors):
        self._set_running(False)
        self._render_cards()
        summary = u"✓ Updated {} of {} matched elements.".format(
            total_updated, total_processed)
        if errors:
            summary += "  Issues: " + "; ".join(errors)
        self._set_status(summary)
