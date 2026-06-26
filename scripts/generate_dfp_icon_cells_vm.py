# -*- coding: utf-8 -*-
"""VM runner for DFP icon cell generation."""
from __future__ import print_function
import os
import sys

ROOT = "/tmp/dfp_gen"
sys.path.insert(0, ROOT)

# Patch paths then run generator logic inline
exec(open("/tmp/generate_dfp_icon_cells.py").read().replace(
    "ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))",
    "ROOT = '/tmp/dfp_gen'",
))
