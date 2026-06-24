# -*- coding: utf-8 -*-
"""Thin REST/GraphQL transport for the CDE backend.

Wraps ``urllib2`` with Bearer auth pulled from a :class:`CDEAuthClient`, json
encode/decode and IronPython-safe utf-8 handling. Higher layers
(:mod:`cde.service`) build domain calls on top of this.
"""
import json
import urllib
import urllib2

from collections import namedtuple

from pyrevit import script

from cde import config

logger = script.get_logger()

CDEResponse = namedtuple("CDEResponse", ["data", "status", "headers"])


class CDEAuthError(Exception):
    """Raised when the backend rejects the credentials (HTTP 401/403)."""


class CDEApiError(Exception):
    """Raised for other transport / backend errors."""

    def __init__(self, message, status=None, headers=None, body=None):
        super(CDEApiError, self).__init__(message)
        self.status = status
        self.headers = headers or {}
        self.body = body


class CDEPreconditionError(CDEApiError):
    """Raised on HTTP 412 (stale If-Match / etag)."""


class CDEConflictError(CDEApiError):
    """Raised on HTTP 409 (conflicting mutation)."""


class CDEClient(object):
    """Authenticated REST + GraphQL transport."""

    def __init__(self, auth):
        self.auth = auth

    @property
    def base_url(self):
        return self.auth.base_url

    # --- low level ------------------------------------------------------

    def _send(self, url, method="GET", payload=None, extra_headers=None,
              return_meta=False):
        data = None
        headers = {}
        headers.update(self.auth.authorized_header())
        if extra_headers:
            headers.update(extra_headers)
        if payload is not None:
            data = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json; charset=utf-8"

        req = urllib2.Request(url, data=data)
        req.get_method = lambda: method
        config.apply_http_headers(req, extra=headers)

        try:
            response = urllib2.urlopen(req)
        except urllib2.HTTPError as ex:
            resp_headers = self._headers_dict(ex.info())
            body = ""
            try:
                body = ex.read()
            except Exception:
                pass
            if ex.code in (401, 403):
                raise CDEAuthError("Not authorized (HTTP {}).".format(ex.code))
            decoded_body = self._decode_body(body)
            message = "HTTP {} {} for {}: {}".format(
                ex.code, ex.reason, url, body)
            if ex.code == 412:
                raise CDEPreconditionError(
                    message, status=ex.code, headers=resp_headers, body=decoded_body)
            if ex.code == 409:
                raise CDEConflictError(
                    message, status=ex.code, headers=resp_headers, body=decoded_body)
            raise CDEApiError(
                message, status=ex.code, headers=resp_headers, body=decoded_body)
        except Exception as ex:
            raise CDEApiError("Request to {} failed: {}".format(url, ex))

        raw = response.read()
        resp_headers = self._headers_dict(response.info())
        decoded = self._decode_body(raw)
        if return_meta:
            return CDEResponse(decoded, response.getcode(), resp_headers)
        return decoded

    def get(self, url, params=None, extra_headers=None, return_meta=False):
        if params:
            url = "{}?{}".format(url, urllib.urlencode(params))
        return self._send(url, "GET", extra_headers=extra_headers,
                          return_meta=return_meta)

    def post(self, url, payload=None, extra_headers=None, return_meta=False):
        return self._send(
            url, "POST", payload if payload is not None else {},
            extra_headers=extra_headers, return_meta=return_meta)

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

    def _headers_dict(self, info):
        if info is None:
            return {}
        try:
            return dict(info)
        except Exception:
            return {}

    def _decode_body(self, raw):
        if not raw:
            return {}
        try:
            return self._decode_utf8(json.loads(raw))
        except ValueError:
            if isinstance(raw, str):
                try:
                    return raw.decode("utf-8")
                except (UnicodeError, AttributeError):
                    return raw
            return raw

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
