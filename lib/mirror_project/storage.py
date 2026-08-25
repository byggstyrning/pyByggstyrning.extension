# -*- coding: utf-8 -*-
"""JSON snapshot/manifest IO with atomic replace and type sanitizing."""

from __future__ import print_function

import json
import os
import tempfile

try:
    unicode
except NameError:
    unicode = str  # noqa: A001  Python 3

try:
    long
except NameError:
    long = int  # noqa: A001  Python 3


def _to_unicode(value):
    """Decode Revit/Windows/CLR strings to plain Python unicode for json.dumps."""
    if value is None:
        return u''
    if isinstance(value, unicode):
        try:
            # IronPython: System.String may isinstance as unicode but break stdlib json.
            return unicode(value)
        except Exception:
            pass
    decode = getattr(value, 'decode', None)
    if decode is not None and not isinstance(value, unicode):
        for encoding in ('utf-8', 'cp1252', 'latin-1'):
            try:
                return decode(encoding)
            except (UnicodeDecodeError, LookupError, TypeError, ValueError, AttributeError):
                continue
        try:
            return unicode(value, 'latin-1', 'replace')
        except Exception:
            return unicode(repr(value))
    try:
        return unicode(value)
    except Exception:
        try:
            raw = str(value)
        except Exception:
            return unicode(repr(value))
        if isinstance(raw, unicode):
            return raw
        for encoding in ('utf-8', 'cp1252', 'latin-1'):
            try:
                return raw.decode(encoding)
            except (UnicodeDecodeError, LookupError, AttributeError):
                continue
        try:
            return raw.decode('latin-1', 'replace')
        except Exception:
            return unicode(repr(raw))


def _json_quote(text):
    """Escape one string as a JSON literal using unicode ord() only."""
    u = _to_unicode(text)
    parts = [u'"']
    for ch in u:
        o = ord(ch)
        if ch == u'"':
            parts.append(u'\\"')
        elif ch == u'\\':
            parts.append(u'\\\\')
        elif ch == u'\b':
            parts.append(u'\\b')
        elif ch == u'\f':
            parts.append(u'\\f')
        elif ch == u'\n':
            parts.append(u'\\n')
        elif ch == u'\r':
            parts.append(u'\\r')
        elif ch == u'\t':
            parts.append(u'\\t')
        elif o < 0x20 or o > 0x7e:
            parts.append(u'\\u{:04x}'.format(o))
        else:
            parts.append(ch)
    parts.append(u'"')
    return u''.join(parts)


def _format_json_float(value):
    if value != value:
        return u'null'
    text = repr(float(value))
    if text == 'nan' or text == 'inf' or text == '-inf':
        return u'null'
    return unicode(text)


def _dumps_json(obj, indent=2, level=0):
    """Serialize JSON without IronPython stdlib json.dumps encoding bugs."""
    if indent:
        sp = u' ' * (indent * level)
        sp1 = u' ' * (indent * (level + 1))
        nl = u'\n'
    else:
        sp = u''
        sp1 = u''
        nl = u''

    if obj is None:
        return u'null'
    if obj is True:
        return u'true'
    if obj is False:
        return u'false'
    if isinstance(obj, (int, long)):
        return unicode(int(obj))
    if isinstance(obj, float):
        return _format_json_float(obj)
    if isinstance(obj, (str, unicode)):
        return _json_quote(obj)
    if isinstance(obj, (list, tuple)):
        if not obj:
            return u'[]'
        items = []
        for item in obj:
            rendered = _dumps_json(item, indent, level + 1)
            items.append((sp1 + rendered) if indent else rendered)
        inner = (u',' + nl).join(items)
        if indent:
            return u'[' + nl + inner + nl + sp + u']'
        return u'[' + inner + u']'
    if isinstance(obj, dict):
        if not obj:
            return u'{}'
        keys = sorted(obj.keys(), key=lambda k: _to_unicode(k))
        items = []
        for key in keys:
            k_json = _json_quote(_to_unicode(key))
            v_json = _dumps_json(obj[key], indent, level + 1)
            if indent:
                items.append(sp1 + k_json + u': ' + v_json)
            else:
                items.append(k_json + u':' + v_json)
        inner = (u',' + nl).join(items)
        if indent:
            return u'{' + nl + inner + nl + sp + u'}'
        return u'{' + inner + u'}'
    return _json_quote(_to_unicode(obj))


def json_safe(obj):
    """Convert nested values to JSON-serializable Python primitives."""
    if obj is None or isinstance(obj, bool):
        return obj
    if isinstance(obj, float):
        if obj != obj:  # NaN
            return None
        return obj
    if isinstance(obj, (int, long)):
        return int(obj)
    if isinstance(obj, (str, unicode)):
        return _to_unicode(obj)
    if isinstance(obj, dict):
        safe = {}
        for key, value in obj.items():
            safe[_to_unicode(key)] = json_safe(value)
        return safe
    if isinstance(obj, (list, tuple)):
        return [json_safe(item) for item in obj]
    try:
        return _to_unicode(obj)
    except Exception:
        return repr(obj)


_JSON_EXTS = ('.json',)
_CSV_EXTS = ('.csv',)


def validate_user_data_path(path, allowed_exts=None):
    """Reject empty, NUL, traversal, and unexpected extensions."""
    if allowed_exts is None:
        allowed_exts = _JSON_EXTS
    if path is None:
        raise ValueError('Path is required')
    text = unicode(path)
    if not text or '\x00' in text:
        raise ValueError('Invalid path')
    normalized = os.path.normpath(text)
    abs_path = os.path.abspath(normalized)
    parts = abs_path.replace('\\', '/').split('/')
    if '..' in parts:
        raise ValueError('Path traversal is not allowed')
    ext = os.path.splitext(abs_path)[1].lower()
    if ext not in allowed_exts:
        raise ValueError('Unsupported file type: {}'.format(ext or '(none)'))
    return abs_path


def load_json(path):
    """Load UTF-8 JSON from ``path``."""
    safe_path = validate_user_data_path(path, _JSON_EXTS)
    with open(safe_path, 'rb') as handle:
        raw = handle.read()
    if raw.startswith(b'\xef\xbb\xbf'):
        raw = raw[3:]
    text = raw.decode('utf-8')
    return json.loads(text)


def save_json(path, data):
    """Write JSON atomically (temp file then replace)."""
    safe_path = validate_user_data_path(path, _JSON_EXTS)
    directory = os.path.dirname(safe_path)
    if directory and not os.path.isdir(directory):
        os.makedirs(directory)
    safe_data = json_safe(data)
    payload = _dumps_json(safe_data, indent=2)
    if not payload.endswith(u'\n'):
        payload = payload + u'\n'
    # IronPython: avoid str.encode() decode trap; dumper returns unicode.
    if isinstance(payload, unicode):
        raw = payload.encode('utf-8')
    else:
        raw = payload
    fd, tmp_path = tempfile.mkstemp(
        prefix='mirror_project_',
        suffix='.tmp',
        dir=directory or None,
    )
    try:
        with os.fdopen(fd, 'wb') as handle:
            handle.write(raw)
            handle.flush()
            try:
                os.fsync(handle.fileno())
            except (OSError, AttributeError):
                pass
        _replace_file(tmp_path, safe_path)
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        raise


def _replace_file(src, dest):
    if hasattr(os, 'replace'):
        os.replace(src, dest)
        return
    if os.path.exists(dest):
        bak = dest + '.bak'
        try:
            if os.path.exists(bak):
                os.remove(bak)
            os.rename(dest, bak)
            os.rename(src, dest)
            if os.path.exists(bak):
                os.remove(bak)
        except Exception:
            if os.path.exists(bak) and not os.path.exists(dest):
                os.rename(bak, dest)
            raise
    else:
        os.rename(src, dest)
