# -*- coding: utf-8 -*-
"""Interactive Azure-OIDC login, brokered by the CDE backend.

The add-in never holds Azure secrets: it opens the browser to the backend's
``/auth/azure/start`` endpoint, captures the redirect ``code`` on a loopback
listener, then exchanges it at ``/auth/azure/exchange`` for a CDE-issued JWT.
The JWT (and optional per-project API key) is cached under ``%APPDATA%/pyBS``.
"""
import os
import json
import time
import urllib
import urllib2

import clr
clr.AddReference("System")

from pyrevit import script

from cde import config
from cde.oidc_listener import LoopbackRedirectListener, stop_active_listener

logger = script.get_logger()


def _post_json(url, payload, headers=None):
    """POST a json body and return the decoded json response.

    Raises urllib2.HTTPError / URLError on transport failures.
    """
    data = json.dumps(payload).encode("utf-8")
    req = urllib2.Request(url, data=data)
    config.apply_http_headers(req, extra=headers)
    req.add_header("Content-Type", "application/json; charset=utf-8")
    response = urllib2.urlopen(req)
    body = response.read()
    if not body:
        return {}
    return json.loads(body)


def _read_http_error_body(ex):
    """Return a truncated response body from an HTTPError, if any."""
    try:
        raw = ex.read()
        if not raw:
            return ""
        return raw.decode("utf-8", errors="replace")[:500]
    except Exception:
        return ""


def _open_browser(url):
    """Open the system browser at ``url`` (Windows-first, with fallbacks)."""
    try:
        from System.Diagnostics import Process, ProcessStartInfo
        psi = ProcessStartInfo(url)
        psi.UseShellExecute = True
        Process.Start(psi)
        return True
    except Exception as ex:
        logger.debug("CDE: Process.Start failed ({}); trying webbrowser".format(ex))
    try:
        import webbrowser
        webbrowser.open(url)
        return True
    except Exception as ex:
        logger.error("CDE: could not open browser: {}".format(ex))
        return False


class CDEAuthClient(object):
    """Holds and refreshes the CDE auth state for a session."""

    def __init__(self, base_url=None):
        self.base_url = (base_url or config.get_base_url()).rstrip("/")
        self.access_token = None
        self.expires_at = 0
        self.user = None
        self.api_key = None
        self.last_error = None
        self._login_listener = None
        self._load()

    # --- state ----------------------------------------------------------

    def is_expired(self):
        if not self.access_token:
            return True
        return time.time() >= (self.expires_at - config.TOKEN_EXPIRY_SKEW_SECONDS)

    def is_authenticated(self):
        return bool(self.api_key) or (bool(self.access_token) and not self.is_expired())

    def authorized_header(self):
        """Return the auth header dict for API calls, or {} if unauthenticated."""
        if self.api_key:
            return {"Authorization": "Bearer {}".format(self.api_key)}
        if self.access_token and not self.is_expired():
            return {"Authorization": "Bearer {}".format(self.access_token)}
        return {}

    @property
    def user_email(self):
        if isinstance(self.user, dict):
            return self.user.get("email")
        return None

    # --- interactive login ---------------------------------------------

    def cancel_login(self):
        """Abort an in-flight browser login and release the loopback listener."""
        listener = self._login_listener
        if listener is not None:
            listener.stop()
            self._login_listener = None

    def login_interactive(self, timeout_seconds=None):
        """Run the full browser redirect login. Returns True on success."""
        self.last_error = None
        self.cancel_login()
        stop_active_listener()

        listener = LoopbackRedirectListener()
        self._login_listener = listener
        try:
            redirect_uri = listener.start()
        except Exception as ex:
            self.last_error = "Could not start local login listener: {}".format(ex)
            logger.error("CDE: {}".format(self.last_error))
            self._login_listener = None
            return False

        try:
            start_url = "{}?{}".format(
                config.auth_start_url(self.base_url),
                urllib.urlencode({"return_to": redirect_uri}))
            if not _open_browser(start_url):
                self.last_error = "Could not open the browser for login."
                return False

            code = listener.wait_for_code(timeout_seconds)
            if listener.error:
                if listener.error == "Login cancelled.":
                    self.last_error = listener.error
                    logger.debug("CDE: interactive login cancelled")
                else:
                    self.last_error = "Login failed: {}".format(listener.error)
                return False
            if not code:
                self.last_error = "No authorization code received."
                return False

            return self._exchange_code(code)
        finally:
            listener.stop()
            if self._login_listener is listener:
                self._login_listener = None

    def _exchange_code(self, code):
        try:
            result = _post_json(
                config.auth_exchange_url(self.base_url), {"code": code})
        except urllib2.HTTPError as ex:
            body = _read_http_error_body(ex)
            self.last_error = "Token exchange failed: HTTP {} {}{}".format(
                ex.code, ex.reason,
                " - {}".format(body) if body else "")
            logger.error("CDE: {}".format(self.last_error))
            return False
        except Exception as ex:
            self.last_error = "Token exchange failed: {}".format(ex)
            logger.error("CDE: {}".format(self.last_error))
            return False

        token = result.get("access_token")
        if not token:
            self.last_error = "Token exchange response missing access_token."
            return False

        self.access_token = token
        self.expires_at = float(result.get("expires_at") or 0)
        self.user = result.get("user")
        self._save()
        return True

    # --- API key (headless fallback) -----------------------------------

    def set_api_key(self, api_key):
        """Use a per-project CDE API key instead of the interactive token."""
        self.api_key = api_key or None
        self._save()

    # --- persistence ----------------------------------------------------

    def _save(self):
        try:
            token_dir = config.get_token_dir()
            if not os.path.isdir(token_dir):
                os.makedirs(token_dir)
            payload = {
                "base_url": self.base_url,
                "access_token": self.access_token,
                "expires_at": self.expires_at,
                "user": self.user,
                "api_key": self.api_key,
            }
            with open(config.get_token_file(), "w") as fh:
                json.dump(payload, fh)
        except Exception as ex:
            logger.warn("CDE: could not save token: {}".format(ex))

    def _load(self):
        path = config.get_token_file()
        if not os.path.isfile(path):
            return
        try:
            with open(path, "r") as fh:
                payload = json.load(fh)
            # Only adopt cached creds that match the active backend.
            if (payload.get("base_url") or "").rstrip("/") != self.base_url:
                return
            self.access_token = payload.get("access_token")
            self.expires_at = float(payload.get("expires_at") or 0)
            self.user = payload.get("user")
            self.api_key = payload.get("api_key")
        except Exception as ex:
            logger.debug("CDE: could not load token: {}".format(ex))

    def logout(self):
        self.access_token = None
        self.expires_at = 0
        self.user = None
        self.api_key = None
        try:
            path = config.get_token_file()
            if os.path.isfile(path):
                os.remove(path)
        except Exception as ex:
            logger.debug("CDE: could not remove token file: {}".format(ex))
