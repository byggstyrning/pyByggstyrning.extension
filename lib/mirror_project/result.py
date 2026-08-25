# -*- coding: utf-8 -*-
"""Collector result envelope (status, data, timing, errors)."""

from __future__ import print_function

import time

from .schema import STATUS_ERROR, STATUS_NOT_AVAILABLE, STATUS_OK


def new_result(status, data=None, timing_ms=0.0, errors=None):
    """Return a JSON-safe collector result dict."""
    if errors is None:
        errors = []
    return {
        'status': status,
        'data': data,
        'timing_ms': round(float(timing_ms), 2),
        'errors': list(errors),
    }


def run_collector(name, fn, *args):
    """Run one collector and wrap exceptions as an error result.

    ``fn`` may return a plain data payload or a dict with ``availability``
    set to ``not_available``.
    """
    started = time.time()
    try:
        data = fn(*args)
        status = STATUS_OK
        if isinstance(data, dict) and data.get('availability') == STATUS_NOT_AVAILABLE:
            status = STATUS_NOT_AVAILABLE
        return name, new_result(status, data=data, timing_ms=_elapsed_ms(started))
    except Exception as ex:
        return name, new_result(
            STATUS_ERROR,
            data=None,
            timing_ms=_elapsed_ms(started),
            errors=[_format_exc(ex)],
        )


def _elapsed_ms(started):
    return (time.time() - started) * 1000.0


def _format_exc(ex):
    try:
        text = str(ex)
    except Exception:
        text = repr(ex)
    cls = type(ex).__name__
    if text:
        return '{}: {}'.format(cls, text)
    return cls
