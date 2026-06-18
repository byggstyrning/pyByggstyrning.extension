# -*- coding: utf-8 -*-
"""Thin REST/GraphQL transport for the CDE backend.

Wraps ``urllib2`` with Bearer auth pulled from a :class:`CDEAuthClient`, json
encode/decode and IronPython-safe utf-8 handling. Higher layers
(:mod:`cde.service`) build domain calls on top of this.
"""
import json
import urllib
import urllib2

from pyrevit import script

from cde import config

logger = script.get_logger()


class CDEAuthError(Exception):
    """Raised when the backend rejects the credentials (HTTP 401/403)."""


class CDEApiError(Exception):
    """Raised for other transport / backend errors."""


class CDEClient(object):
    """Authenticated REST + GraphQL transport."""

    def __init__(self, auth):
        self.auth = auth

    @property
    def base_url(self):
        return self.auth.base_url

    # --- low level ------------------------------------------------------

    def _send(self, url, method="GET", payload=None):
        data = None
        headers = {"Accept": "application/json"}
        headers.update(self.auth.authorized_header())
        if payload is not None:
            data = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json; charset=utf-8"

        req = urllib2.Request(url, data=data)
        req.get_method = lambda: method
        for key, value in headers.items():
            req.add_header(key, value)

        try:
            response = urllib2.urlopen(req)
        except urllib2.HTTPError as ex:
            if ex.code in (401, 403):
                raise CDEAuthError("Not authorized (HTTP {}).".format(ex.code))
            body = ""
            try:
                body = ex.read()
            except Exception:
                pass
            raise CDEApiError("HTTP {} {} for {}: {}".format(
                ex.code, ex.reason, url, body))
        except Exception as ex:
            raise CDEApiError("Request to {} failed: {}".format(url, ex))

        raw = response.read()
        if not raw:
            return {}
        try:
            return self._decode_utf8(json.loads(raw))
        except ValueError:
            return raw

    def get(self, url, params=None):
        if params:
            url = "{}?{}".format(url, urllib.urlencode(params))
        return self._send(url, "GET")

    def post(self, url, payload=None):
        return self._send(url, "POST", payload if payload is not None else {})

    def graphql(self, query, variables=None):
        """Execute a GraphQL query/mutation; returns the ``data`` object.

        Raises CDEApiError if the response carries GraphQL ``errors``.
        """
        body = {"query": query, "variables": variables or {}}
        result = self.post(config.graphql_url(self.base_url), body)
        if isinstance(result, dict) and result.get("errors"):
            raise CDEApiError("GraphQL errors: {}".format(result["errors"]))
        if isinstance(result, dict):
            return result.get("data", result)
        return result

    # --- helpers --------------------------------------------------------

    def _decode_utf8(self, data):
        """Recursively coerce byte strings to unicode (IronPython 2.7)."""
        if isinstance(data, dict):
            return dict((self._decode_utf8(k), self._decode_utf8(v))
                        for k, v in data.items())
        if isinstance(data, list):
            return [self._decode_utf8(v) for v in data]
        if isinstance(data, str):
            try:
                return data.decode("utf-8")
            except (UnicodeError, AttributeError):
                return data
        return data
