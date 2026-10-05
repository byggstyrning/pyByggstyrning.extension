# -*- coding: utf-8 -*-
"""ExternalEvent plumbing + active-view wiring for the modeless schedule window.

A modeless WPF window cannot touch the Revit API directly; it must marshal work
onto Revit's thread via ``ExternalEvent``. ``ExternalEventRunner`` lets the UI
queue arbitrary callables, and ``ActiveViewWatcher`` notifies the UI when the
user switches views (for the "active view only" toggle).
"""
import time

import clr
clr.AddReference("RevitAPIUI")
from Autodesk.Revit import UI

from pyrevit import script

logger = script.get_logger()

_VIEW_THROTTLE_MS = 400


class _RevitTaskHandler(UI.IExternalEventHandler):
    """Runs queued callables ``func(uiapp)`` on the Revit API thread."""

    def __init__(self):
        self._queue = []

    def enqueue(self, func):
        self._queue.append(func)

    def clear_queue(self):
        """Drop queued (not yet executed) tasks. Returns count cleared."""
        n = len(self._queue)
        self._queue = []
        return n

    def Execute(self, uiapp):
        pending, self._queue = self._queue, []
        for func in pending:
            try:
                func(uiapp)
            except Exception as ex:
                logger.error("CDE: external event task failed: {}".format(ex))

    def GetName(self):
        return "CDE Revit Task"


class ExternalEventRunner(object):
    """Queues UI-thread work for execution on the Revit API thread."""

    def __init__(self):
        self._handler = _RevitTaskHandler()
        self._event = UI.ExternalEvent.Create(self._handler)

    def run(self, func):
        """Queue ``func(uiapp)`` and request execution. Returns immediately."""
        self._handler.enqueue(func)
        self._event.Raise()

    def clear_pending(self):
        """Discard queued Revit tasks (e.g. before Apply). Returns count cleared."""
        n = self._handler.clear_queue()
        return n


class ActiveViewWatcher(object):
    """Calls ``on_view_changed(view)`` (throttled) when the active view changes."""

    def __init__(self, uiapp, on_view_changed):
        self._uiapp = uiapp
        self._callback = on_view_changed
        self._handler = None
        self._last_ms = 0
        self._enabled = False

    def start(self, uiapp=None):
        """Subscribe to ViewActivated. Must be called on the Revit API thread."""
        if self._enabled:
            return
        if uiapp is not None:
            self._uiapp = uiapp
        try:
            from Autodesk.Revit.UI.Events import ViewActivatedEventArgs
            from System import EventHandler
            self._handler = EventHandler[ViewActivatedEventArgs](
                self._on_view_activated)
        except Exception:
            self._handler = self._on_view_activated
        try:
            self._uiapp.ViewActivated += self._handler
            self._enabled = True
        except Exception as ex:
            inner = getattr(ex, "InnerException", None)
            clr_type = ex.GetType().FullName if hasattr(ex, "GetType") else type(ex).__name__
            logger.error(
                "CDE: could not subscribe to ViewActivated: {} (type={}, inner={})".format(
                    ex, clr_type, inner))

    def stop(self):
        """Unsubscribe from ViewActivated. Must be called on the Revit API thread."""
        if not self._enabled:
            return
        try:
            self._uiapp.ViewActivated -= self._handler
        except Exception:
            pass
        self._enabled = False

    def _on_view_activated(self, sender, args):
        now_ms = int(time.time() * 1000)
        if now_ms - self._last_ms < _VIEW_THROTTLE_MS:
            return
        self._last_ms = now_ms
        try:
            view = getattr(args, "CurrentActiveView", None)
            self._callback(view)
        except Exception as ex:
            logger.error("CDE: active-view callback failed: {}".format(ex))
