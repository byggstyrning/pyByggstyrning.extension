# -*- coding: utf-8 -*-
"""Post native Revit commands. Revit owns dialogs and cancellation."""

from __future__ import print_function

import sys

from Autodesk.Revit.UI import PostableCommand, RevitCommandId

COMMANDS = {
    'mirror_project': PostableCommand.MirrorProject,
}

_SYS_KEY = '_mirror_project_deferred_posters'
_MAX_IDLE_ATTEMPTS = 100

class DeferredNativeCommand(object):
    """Post a built-in Revit command on Idling so modal pyRevit UI cannot skip steps."""

    def __init__(self, uiapp, command_key, logger=None):
        self._uiapp = uiapp
        self._command_key = command_key
        self._logger = logger
        self._handler = None
        self._posted = False
        self._attempts = 0
        self._command_id = None
        self._RevitCommandId = RevitCommandId

    def start(self):
        if self._command_key not in COMMANDS:
            raise ValueError('Unknown command {}'.format(self._command_key))
        self._command_id = self._RevitCommandId.LookupPostableCommandId(
            COMMANDS[self._command_key])
        if self._command_id is None:
            raise ValueError('Postable command id not found for {}'.format(self._command_key))
        self._handler = self._on_idling
        self._uiapp.Idling += self._handler
        if not hasattr(sys, _SYS_KEY):
            setattr(sys, _SYS_KEY, [])
        getattr(sys, _SYS_KEY).append(self)

    def stop(self):
        if self._handler is not None:
            try:
                self._uiapp.Idling -= self._handler
            except Exception:
                pass
            self._handler = None
        try:
            getattr(sys, _SYS_KEY).remove(self)
        except (AttributeError, ValueError):
            pass

    def _on_idling(self, sender, args):
        if self._posted:
            self.stop()
            return
        self._attempts += 1
        try:
            can_post = self._uiapp.CanPostCommand(self._command_id)
        except Exception:
            can_post = False
        if not can_post:
            if self._attempts >= _MAX_IDLE_ATTEMPTS:
                self.stop()
            return
        try:
            self._uiapp.PostCommand(self._command_id)
            self._posted = True
        except Exception as ex:
            if self._logger:
                self._logger.warning('Deferred PostCommand failed: {}'.format(ex))
        self.stop()


def schedule_native_command(uiapp, command_key, logger=None):
    """Queue native Revit command after this external command returns."""
    try:
        DeferredNativeCommand(uiapp, command_key, logger=logger).start()
    except Exception as ex:
        return False, str(ex)
    return True, 'Scheduled {}'.format(command_key)


def post_native_command(uiapp, command_key):
    """Immediate PostCommand (legacy). Prefer schedule_native_command in pyRevit UI."""
    if command_key not in COMMANDS:
        return False, 'Unknown command {}'.format(command_key)
    try:
        command_id = RevitCommandId.LookupPostableCommandId(COMMANDS[command_key])
    except Exception as ex:
        return False, 'LookupPostableCommandId failed: {}'.format(ex)
    if command_id is None:
        return False, 'Postable command id not found for {}'.format(command_key)
    try:
        uiapp.PostCommand(command_id)
    except Exception as ex:
        return False, 'PostCommand failed: {}'.format(ex)
    return True, 'Posted {}'.format(command_key)
