# -*- coding: utf-8 -*-
"""Companion button so the Test stack has two items (pyRevit requirement)."""

__title__ = 'Ping'
__author__ = 'Byggstyrning AB'
__doc__ = """Minimal dialog to confirm the stack second slot works."""

from pyrevit import forms

forms.alert('Sandbox stack OK.', title='Ping')
