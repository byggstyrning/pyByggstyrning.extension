# -*- coding: utf-8 -*-
"""Guided Mirror Project preflight, baseline, native command launch, and compare."""

__title__ = "Mirror\nProject"
__author__ = "Byggstyrning AB"
__doc__ = (
    "Guided Mirror Project workflow: preflight snapshot, baseline JSON, "
    "disable eligible locked constraints (not restored), launch native "
    "Mirror Project, compare, and create source-model review views."
)
__highlight__ = "new"

import sys
import os.path as op

from pyrevit import forms, revit, script

script_path = __file__
extension_dir = op.dirname(op.dirname(op.dirname(op.dirname(script_path))))
lib_path = op.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

logger = script.get_logger()


def main():
    try:
        from mirror_project.workflow import run_menu
    except Exception as ex:
        logger.error('Mirror Project import failed: {}'.format(ex))
        forms.alert(
            'Failed to import Mirror Project library. Reload the extension after lib/ changes.\n\n{}'.format(ex),
            title='Mirror Project',
        )
        return
    run_menu(revit.doc, __revit__)


main()
