# -*- coding: utf-8 -*-
"""Switch local git branch for pyByggstyrning and reload pyRevit (WPF window)."""
# pylint: disable=import-error,invalid-name,broad-except

__title__ = 'Select Branch'
__author__ = 'Byggstyrning AB'
__highlight__ = 'new'
__doc__ = """Pick a branch from origin (remote-tracking refs), checkout, reload pyRevit."""

import sys
import os.path as op

from pyrevit import forms
from pyrevit import script
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo

logger = script.get_logger()

script_dir = op.dirname(__file__)
stack_dir = op.dirname(script_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

import extension_updater as updater  # noqa: E402


class BranchSwitcherWindow(forms.WPFWindow):
    """Pick a branch from origin (see ``list_remote_branches``), checkout, reload."""

    def __init__(self):
        self._repo_missing = False
        xaml_file = op.join(script_dir, 'BranchSwitcherWindow.xaml')
        forms.WPFWindow.__init__(self, xaml_file)

        if not updater.get_extension_dir():
            self._repo_missing = True
            self.currentBranchText.Text = 'Not a git clone.'
            self.branchList.IsEnabled = False
            self.switchReloadButton.IsEnabled = False
            return

        self.load_branch_list()

    def load_branch_list(self):
        cur = updater.get_current_branch()
        line1 = 'Current: {}'.format(cur) if cur else 'Current: (detached or unknown)'
        branches = updater.list_remote_branches()
        self._branches = list(branches)
        self.branchList.ItemsSource = self._branches
        if self._branches:
            line2 = 'Branches on origin (from last git fetch).'
        else:
            line2 = 'No origin/* branches in this clone — run git fetch first.'
        self.currentBranchText.Text = '{}\n{}'.format(line1, line2)
        if cur:
            for idx, name in enumerate(self._branches):
                if name == cur:
                    self.branchList.SelectedIndex = idx
                    break

    def SwitchReloadButton_Click(self, sender, e):
        if self._repo_missing:
            return
        selected = self.branchList.SelectedItem
        if selected is None:
            forms.alert('Select a branch.', title='Select Branch')
            return
        ok, message = updater.checkout_branch(str(selected))
        if not ok:
            forms.alert(message, title='Git checkout failed')
            return
        logger.info('Checked out branch {}, reloading pyRevit'.format(selected))
        self.Close()
        sessionmgr.reload_pyrevit()
        script.get_results().newsession = sessioninfo.get_session_uuid()

    def CancelButton_Click(self, sender, e):
        self.Close()


if __name__ == '__main__':
    try:
        BranchSwitcherWindow().ShowDialog()
    except Exception as err:
        logger.error('Switch branch failed: {}'.format(err))
        forms.alert(str(err), title='Select Branch')
