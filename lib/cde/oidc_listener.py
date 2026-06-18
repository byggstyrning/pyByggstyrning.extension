# -*- coding: utf-8 -*-
"""Loopback redirect listener for the interactive OIDC login.

Starts a ``System.Net.HttpListener`` on a free loopback port, hands its URL to
the CDE backend as ``return_to``, and captures the ``?code=`` (or ``?error=``)
that the backend redirects back with. Runs the blocking listen loop on a
background .NET thread; nothing here touches the Revit API, so it is safe off
the UI thread.
"""
from urlparse import urlparse, parse_qs

import clr
clr.AddReference("System")
from System.Net import HttpListener
from System.Text import Encoding
from System.Threading import Thread, ThreadStart, AutoResetEvent

from pyrevit import script

from cde import config

logger = script.get_logger()

_SUCCESS_HTML = (
    "<html><head><title>Signed in</title></head>"
    "<body style='font-family:Segoe UI,Arial,sans-serif;text-align:center;"
    "margin-top:80px;color:#333'>"
    "<h2>Signed in to the CDE</h2>"
    "<p>You can close this tab and return to Revit.</p>"
    "</body></html>"
)


class LoopbackRedirectListener(object):
    """One-shot loopback listener that captures the OIDC redirect code."""

    def __init__(self):
        self._listener = None
        self._port = None
        self._thread = None
        self._done = AutoResetEvent(False)
        self.code = None
        self.error = None

    @property
    def redirect_uri(self):
        if self._port is None:
            return None
        return "http://{}:{}{}".format(
            config.LOOPBACK_HOST, self._port, config.LOOPBACK_CALLBACK_PATH)

    def start(self):
        """Bind to the first available loopback port and start listening."""
        last_ex = None
        for port in config.LOOPBACK_PORT_RANGE:
            listener = HttpListener()
            # HttpListener prefixes must end with '/'.
            prefix = "http://{}:{}{}/".format(
                config.LOOPBACK_HOST, port,
                config.LOOPBACK_CALLBACK_PATH.rstrip("/"))
            listener.Prefixes.Add(prefix)
            try:
                listener.Start()
            except Exception as ex:
                last_ex = ex
                continue
            self._listener = listener
            self._port = port
            self._thread = Thread(ThreadStart(self._listen_loop))
            self._thread.IsBackground = True
            self._thread.Start()
            logger.debug("CDE: loopback listening on {}".format(prefix))
            return self.redirect_uri
        raise IOError(
            "Could not bind a loopback port for OIDC redirect: {}".format(last_ex))

    def _listen_loop(self):
        try:
            while self._listener is not None and self._listener.IsListening:
                try:
                    context = self._listener.GetContext()
                except Exception:
                    # Listener stopped while blocked in GetContext().
                    break
                query = parse_qs(urlparse(context.Request.Url.AbsoluteUri).query)
                code = query.get("code", [None])[0]
                error = query.get("error", [None])[0]

                self._respond(context, _SUCCESS_HTML)

                if code or error:
                    self.code = code
                    self.error = error
                    self._done.Set()
                    break
        except Exception as ex:
            self.error = str(ex)
            self._done.Set()

    def _respond(self, context, html):
        try:
            buf = Encoding.UTF8.GetBytes(html)
            response = context.Response
            response.ContentType = "text/html; charset=utf-8"
            response.ContentLength64 = buf.Length
            response.OutputStream.Write(buf, 0, buf.Length)
            response.OutputStream.Close()
        except Exception as ex:
            logger.debug("CDE: error writing loopback response: {}".format(ex))

    def wait_for_code(self, timeout_seconds=None):
        """Block until a code/error arrives or the timeout elapses.

        Returns the captured code (``self.code``); ``self.error`` is set when
        the backend reported an error. Returns None on timeout.
        """
        if timeout_seconds is None:
            timeout_seconds = config.LOGIN_TIMEOUT_SECONDS
        signalled = self._done.WaitOne(int(timeout_seconds * 1000))
        if not signalled:
            self.error = "Login timed out after {}s".format(timeout_seconds)
        return self.code

    def stop(self):
        listener = self._listener
        self._listener = None
        if listener is not None:
            try:
                listener.Stop()
                listener.Close()
            except Exception:
                pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()
