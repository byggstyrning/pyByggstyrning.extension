# -*- coding: utf-8 -*-
"""Temporary sandbox command to verify the Test panel appears on the ribbon."""

__title__ = 'Test\nFeature'
__author__ = 'Byggstyrning AB'
__highlight__ = 'new'
__doc__ = """Shows a dialog if this pushbutton loads (ribbon wiring OK)."""

from pyrevit import forms

forms.alert(
    'Test panel and Test Feature button load correctly.',
    title='Test Feature',
)
