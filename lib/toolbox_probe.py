# -*- coding: utf-8 -*-
"""Locate optional pyByggstyrning.toolbox better_schedule implementation."""

import os

_BETTER_SCHEDULE_REL = os.path.join("better_schedule", "script.py")
_INSTALLED_TOOLBOX_DIRNAMES = (
    "pyByggstyrning.toolbox.lib",
    "pyByggstyrning.toolbox",
)


def _norm(path):
    return os.path.normpath(path) if path else None


def _extension_root_from_shim(shim_file):
    # pushbutton -> pulldown -> panel -> tab -> extension root
    pushbutton_dir = os.path.dirname(os.path.abspath(shim_file))
    return _norm(
        os.path.join(
            pushbutton_dir,
            "..",
            "..",
            "..",
            "..",
        )
    )


def find_better_schedule_script(shim_file=None):
    """Return path to better_schedule/script.py or None."""
    candidates = []

    env_root = os.environ.get("PYBYGGSTYRNING_TOOLBOX_ROOT")
    if env_root:
        candidates.append(_norm(env_root))

    if shim_file:
        ext_root = _extension_root_from_shim(shim_file)
        if ext_root:
            parent = os.path.dirname(ext_root)
            for name in _INSTALLED_TOOLBOX_DIRNAMES:
                candidates.append(_norm(os.path.join(parent, name)))

    seen = set()
    for root in candidates:
        if not root or root in seen:
            continue
        seen.add(root)
        script_path = _norm(os.path.join(root, _BETTER_SCHEDULE_REL))
        if os.path.isfile(script_path):
            return script_path
    return None
