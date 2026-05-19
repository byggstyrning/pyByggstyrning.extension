# -*- coding: utf-8 -*-
"""Hide Better Schedule when toolbox implementation is not installed locally."""

import os
import sys

__title__ = "Better Schedule"

_shim_dir = os.path.dirname(os.path.abspath(__file__))
_ext_lib = os.path.normpath(
    os.path.join(_shim_dir, "..", "..", "..", "..", "lib")
)
if _ext_lib not in sys.path:
    sys.path.insert(0, _ext_lib)

from toolbox_probe import find_better_schedule_script

_script = find_better_schedule_script(shim_file=os.path.join(_shim_dir, "script.py"))
__hideif__ = not (_script and os.path.isfile(_script))
