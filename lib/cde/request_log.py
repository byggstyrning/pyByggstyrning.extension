# -*- coding: utf-8 -*-
"""Full HTTP capture for CDE traffic (network-engineer handoff).

Outputs under ``%APPDATA%/pyBS/``:

* ``cde_http_capture.jsonl`` — index (one line per exchange)
* ``cde_http_GETs.txt`` — every GET, one line each (grep-friendly)
* ``cde_http_trace.txt`` — raw request/response trace (all methods)
* ``cde_http_capture/exchange_NNNN.json`` — full request + response bodies

Authorization tokens are never stored. Large bodies are always saved to
``exchange_*.json``; trace file references or embeds them when small.
"""
import codecs
import json
import os
import re
import time

from pyrevit import script

from cde import config

logger = script.get_logger()


def _to_unicode(text):
    """Coerce bytes -> unicode (utf-8, latin-1 fallback) for safe file writes."""
    if isinstance(text, bytes) and not isinstance(text, unicode):
        try:
            return text.decode("utf-8")
        except Exception:
            return text.decode("latin-1", "replace")
    return text


def _write_text(path, mode, text):
    """Write unicode text as UTF-8. IronPython text-mode open() decodes raw
    byte-strings via the system code page and throws on non-ASCII (e.g. 0xF6
    'ö'); codecs.open with utf-8 + unicode text avoids that."""
    with codecs.open(path, mode, encoding="utf-8") as handle:
        handle.write(_to_unicode(text))

_CAPTURE_DIR_NAME = "cde_http_capture"
_INDEX_NAME = "cde_http_capture.jsonl"
_GET_INDEX_NAME = "cde_http_GETs.txt"
_TRACE_NAME = "cde_http_trace.txt"
_INLINE_BODY_MAX = 32768
_DOOR_CLASSES = frozenset(["IFCDOOR", "IFCDOORSTANDARDCASE"])

_exchange_counter = 0
_current_session = None
_session_get_urls = []

# Known CDE GET endpoints (for backend reference; runtime GETs also logged).
_KNOWN_GET_ENDPOINTS = [
    "/api/v1/projects",
    "/api/v1/projects/{projectId}/ifc-versions",
    "/api/v1/projects/{projectId}/live-drops",
    "/api/v1/projects/{projectId}/api-keys",
    "/api/v1/versions/{versionId}/ingest-status",
    "/api/v2/graph/status",
    "/api/v2/graph/overview",
    "/api/v2/graph/analysis",
    "/api/v2/mutations/{mutationId}",
]


def get_log_dir():
    return config.get_token_dir()


def get_capture_dir():
    return os.path.join(get_log_dir(), _CAPTURE_DIR_NAME)


def get_log_path():
    """Primary index file (NDJSON)."""
    return os.path.join(get_log_dir(), _INDEX_NAME)


def get_get_index_path():
    return os.path.join(get_log_dir(), _GET_INDEX_NAME)


def get_trace_path():
    return os.path.join(get_log_dir(), _TRACE_NAME)


def _ensure_dirs():
    capture = get_capture_dir()
    if not os.path.isdir(capture):
        os.makedirs(capture)
    log_dir = get_log_dir()
    if log_dir and not os.path.isdir(log_dir):
        os.makedirs(log_dir)


def _next_exchange_id():
    global _exchange_counter
    _exchange_counter += 1
    return _exchange_counter


def _normalize_ifc_class(value):
    return (value or "").upper().replace("_", "")


def _node_value_summary(node):
    if not node:
        return {}
    authored = node.get("authoredValues") or []
    derived = node.get("derivedValues") or []
    effective = node.get("effectiveValues") or []
    sample_effective = []
    for ev in effective[:5]:
        sample_effective.append({
            "psetName": ev.get("psetName"),
            "propertyName": ev.get("propertyName"),
            "value": ev.get("value"),
        })
    return {
        "globalId": node.get("globalId"),
        "ifcClass": node.get("ifcClass"),
        "name": (node.get("name") or "")[:120],
        "authoredCount": len(authored),
        "derivedCount": len(derived),
        "effectiveCount": len(effective),
        "ruleTraceCount": len(node.get("ruleTrace") or []),
        "relationshipCount": len(node.get("relationships") or []),
        "sampleEffectiveValues": sample_effective,
    }


def _summarize_response_body(body):
    if body is None:
        return None
    if not isinstance(body, dict):
        preview = unicode(body)
        if len(preview) > 2000:
            preview = preview[:2000] + "...<truncated>"
        return {"preview": preview}

    if body.get("errors"):
        return {
            "graphql_errors": body.get("errors"),
            "data_keys": list((body.get("data") or {}).keys()),
        }

    data = body.get("data")
    if not isinstance(data, dict):
        return {"top_level_keys": list(body.keys())[:30]}

    summary = {}
    elements = data.get("elements")
    if isinstance(elements, dict):
        edges = elements.get("edges") or []
        page_info = elements.get("pageInfo") or {}
        ifc_class_counts = {}
        mismatched_filter = 0
        door_samples = []
        for edge in edges:
            node = (edge or {}).get("node") or {}
            ic = node.get("ifcClass") or "?"
            ifc_class_counts[ic] = ifc_class_counts.get(ic, 0) + 1
            norm = _normalize_ifc_class(ic)
            if norm in _DOOR_CLASSES and len(door_samples) < 3:
                door_samples.append(_node_value_summary(node))
            if norm not in _DOOR_CLASSES:
                mismatched_filter += 1
        summary["elements"] = {
            "edgeCount": len(edges),
            "pageInfo": page_info,
            "ifcClassCounts": ifc_class_counts,
            "nonDoorOnPage": mismatched_filter,
            "doorSamples": door_samples,
        }

    element = data.get("element")
    if isinstance(element, dict):
        summary["element"] = _node_value_summary(element)

    if not summary:
        summary["data_keys"] = list(data.keys())[:30]

    return summary


def _redact_headers(headers):
    if not headers:
        return {}
    redacted = {}
    for key, value in headers.items():
        lk = (key or "").lower()
        if lk in ("authorization", "api-key", "x-api-key", "cookie"):
            redacted[key] = "<redacted>"
        elif lk == "idempotency-key":
            redacted[key] = (value or "")[:16] + "..." if value else value
        else:
            redacted[key] = value
    return redacted


def _json_text(obj):
    try:
        return _to_unicode(json.dumps(obj, sort_keys=True, indent=2))
    except Exception:
        return unicode(obj)


def _safe_filename_part(url):
    part = re.sub(r"[^\w\-.]+", "_", url)[:40]
    return part or "request"


def _write_capture_file(exchange_id, method, url, record):
    _ensure_dirs()
    fname = "exchange_{:04d}_{}_{}.json".format(
        exchange_id, method.upper(), _safe_filename_part(url))
    path = os.path.join(get_capture_dir(), fname)
    _write_text(path, "w", _json_text(record))
    return path


def _append_get_index(exchange_id, url, status, duration_ms, capture_path):
    line = "{ts} #{id:04d} GET {status} {ms}ms {url} -> {file}\n".format(
        ts=time.strftime("%Y-%m-%dT%H:%M:%S"),
        id=exchange_id,
        status=status if status is not None else "ERR",
        ms=duration_ms if duration_ms is not None else "?",
        url=url,
        file=os.path.basename(capture_path))
    try:
        _write_text(get_get_index_path(), "a", line)
    except Exception as ex:
        logger.debug("CDE: GET index write failed: {}".format(ex))


def _append_trace(exchange_id, method, url, req_headers, req_body,
                  status, resp_headers, resp_body, error, duration_ms,
                  capture_path):
    lines = [
        "",
        "=" * 80,
        "[{ts}] #{id:04d} {method} {status} {ms}ms".format(
            ts=time.strftime("%Y-%m-%dT%H:%M:%S"),
            id=exchange_id,
            method=method.upper(),
            status=status if status is not None else "ERROR",
            ms=duration_ms if duration_ms is not None else "?"),
        "{method} {url}".format(method=method.upper(), url=url),
        "",
        ">> Request Headers:",
    ]
    for key, val in sorted(_redact_headers(req_headers or {}).items()):
        lines.append("{}: {}".format(key, val))
    lines.append("")
    lines.append(">> Request Body:")
    if req_body is None:
        lines.append("(none)")
    else:
        lines.append(_json_text(req_body))
    lines.append("")
    if error:
        lines.append("!! Transport Error: {}".format(unicode(error)[:2000]))
    else:
        lines.append("<< Response Status: {}".format(status))
        lines.append("<< Response Headers:")
        for key, val in sorted(_redact_headers(resp_headers or {}).items()):
            lines.append("{}: {}".format(key, val))
        lines.append("")
        body_text = _json_text(resp_body)
        if len(body_text) <= _INLINE_BODY_MAX:
            lines.append("<< Response Body:")
            lines.append(body_text)
        else:
            lines.append("<< Response Body: (full file {})".format(capture_path))
            lines.append(body_text[:_INLINE_BODY_MAX] + "\n...<truncated in trace>")
    lines.append("=" * 80)
    try:
        _write_text(get_trace_path(), "a",
                    u"\n".join(_to_unicode(ln) for ln in lines) + u"\n")
    except Exception as ex:
        logger.debug("CDE: HTTP trace write failed: {}".format(ex))


def mark_session(label, **context):
    """Start a capture session (e.g. schedule refresh)."""
    global _current_session, _session_get_urls
    _current_session = "{}_{}".format(label, int(time.time()))
    _session_get_urls = []
    _ensure_dirs()
    entry = {
        "type": "session_start",
        "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "sessionId": _current_session,
        "label": label,
        "context": context,
        "paths": {
            "index": get_log_path(),
            "getIndex": get_get_index_path(),
            "trace": get_trace_path(),
            "captureDir": get_capture_dir(),
        },
        "knownGetEndpoints": list(_KNOWN_GET_ENDPOINTS),
    }
    _append_jsonl(entry)
    _append_known_gets_header()


def finish_session(label=None, **summary):
    """Close session; append GET roll-up for handoff."""
    global _current_session
    entry = {
        "type": "session_end",
        "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "sessionId": _current_session,
        "label": label,
        "summary": summary,
        "getRequestsThisSession": list(_session_get_urls),
        "getCount": len(_session_get_urls),
    }
    _append_jsonl(entry)
    try:
        _write_text(
            get_get_index_path(), "a",
            "# session_end {} get_count={}\n".format(
                _current_session or label, len(_session_get_urls)))
    except Exception:
        pass


def _append_known_gets_header():
    try:
        path = get_get_index_path()
        if os.path.isfile(path) and os.path.getsize(path) > 0:
            return
        header = [
            u"# CDE HTTP GET index — one line per GET request",
            u"# Format: TIMESTAMP #ID GET STATUS DURATION URL -> capture_file",
            u"# Known GET endpoint patterns used by pyRevit CDE:",
        ]
        for ep in _KNOWN_GET_ENDPOINTS:
            header.append(u"#   {}".format(ep))
        header.append(u"#")
        _write_text(path, "w", u"\n".join(header) + u"\n")
    except Exception as ex:
        logger.debug("CDE: GET index header failed: {}".format(ex))


def log_exchange(method, url, request_headers=None, request_body=None,
                 response_status=None, response_headers=None,
                 response_body=None, error=None, duration_ms=None):
    """Record one HTTP exchange (full capture + index + trace)."""
    exchange_id = _next_exchange_id()
    method_up = (method or "GET").upper()

    record = {
        "exchangeId": exchange_id,
        "sessionId": _current_session,
        "ts": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "timing": {"durationMs": duration_ms},
        "request": {
            "method": method_up,
            "url": url,
            "headers": _redact_headers(request_headers or {}),
            "body": request_body,
        },
        "response": {
            "status": response_status,
            "headers": _redact_headers(response_headers or {}),
            "body": response_body,
            "error": unicode(error)[:2000] if error else None,
        },
    }

    capture_path = _write_capture_file(exchange_id, method_up, url, record)

    index_entry = {
        "type": "exchange",
        "exchangeId": exchange_id,
        "sessionId": _current_session,
        "ts": record["ts"],
        "method": method_up,
        "url": url,
        "responseStatus": response_status,
        "durationMs": duration_ms,
        "captureFile": capture_path,
        "responseSummary": _summarize_response_body(response_body),
        "error": unicode(error)[:500] if error else None,
    }
    _append_jsonl(index_entry)

    _append_trace(
        exchange_id, method_up, url, request_headers, request_body,
        response_status, resp_headers=response_headers, resp_body=response_body,
        error=error, duration_ms=duration_ms, capture_path=capture_path)

    if method_up == "GET":
        get_line = {
            "exchangeId": exchange_id,
            "url": url,
            "status": response_status,
            "durationMs": duration_ms,
            "captureFile": capture_path,
        }
        _session_get_urls.append(get_line)
        _append_get_index(
            exchange_id, url, response_status, duration_ms, capture_path)


def _append_jsonl(entry):
    try:
        _ensure_dirs()
        _write_text(get_log_path(), "a",
                    _to_unicode(json.dumps(entry, sort_keys=True)) + u"\n")
    except Exception as ex:
        logger.debug("CDE: capture index write failed: {}".format(ex))
