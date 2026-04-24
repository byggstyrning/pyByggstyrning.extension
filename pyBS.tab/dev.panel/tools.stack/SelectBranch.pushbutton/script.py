# -*- coding: utf-8 -*-
"""Switch local git branch for pyByggstyrning and reload pyRevit (WPF window)."""
# pylint: disable=import-error,invalid-name,broad-except

__title__ = 'Select Branch'
__author__ = 'Byggstyrning AB'
__highlight__ = 'updated'
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
from styles import load_styles_to_window  # noqa: E402

from System.Windows import Visibility  # noqa: E402
from System.Windows.Input import Key  # noqa: E402
from System.Windows.Media import SolidColorBrush, ColorConverter  # noqa: E402


# Branch type classification: prefix -> (label, hex color).
# Order matters: first match wins.
_BRANCH_TYPE_RULES = [
    ('feature/', ('FEATURE', '#2196F3')),
    ('feat/',    ('FEATURE', '#2196F3')),
    ('fix/',     ('FIX',     '#66BB6A')),
    ('bugfix/',  ('FIX',     '#66BB6A')),
    ('hotfix/',  ('HOTFIX',  '#EF5350')),
    ('release/', ('RELEASE', '#AB47BC')),
    ('chore/',   ('CHORE',   '#90A4AE')),
    ('docs/',    ('DOCS',    '#78909C')),
    ('refactor/', ('REFACTOR', '#FFA726')),
    ('experiment/', ('EXPERIMENT', '#26C6DA')),
    ('test/',    ('TEST',    '#BDBDBD')),
]
_DEFAULT_TYPE = ('OTHER', '#9E9E9E')
_MAIN_TYPE = ('MAIN', '#ffbb00')
_MAIN_NAMES = {'master', 'main', 'develop', 'dev', 'trunk'}


def _classify_branch(name):
    """Return (type_label, type_color_hex) for a branch name."""
    lower = name.lower()
    if lower in _MAIN_NAMES:
        return _MAIN_TYPE
    for prefix, info in _BRANCH_TYPE_RULES:
        if lower.startswith(prefix):
            return info
    return _DEFAULT_TYPE


def _hex_to_brush(hex_color):
    """Build a frozen SolidColorBrush from a #RRGGBB string."""
    color = ColorConverter.ConvertFromString(hex_color)
    brush = SolidColorBrush(color)
    brush.Freeze()
    return brush


class BranchItem(object):
    """View-model for a single branch row.

    Exposes attributes consumed by the XAML ItemTemplate bindings.
    """

    def __init__(self, name, is_current, has_local=True, has_remote=True):
        self.name = name
        self.is_current = is_current
        self.has_local = has_local
        self.has_remote = has_remote

        type_label, type_hex = _classify_branch(name)
        self.type_label = type_label
        self.type_color = _hex_to_brush(type_hex)

        self.current_visibility = (
            Visibility.Visible if is_current else Visibility.Collapsed
        )
        self.local_only_visibility = (
            Visibility.Visible if (has_local and not has_remote)
            else Visibility.Collapsed
        )
        self.remote_only_visibility = (
            Visibility.Visible if (has_remote and not has_local)
            else Visibility.Collapsed
        )

    def __str__(self):
        return self.name


class BranchSwitcherWindow(forms.WPFWindow):
    """Pick a branch from origin (see ``list_remote_branches``), checkout, reload."""

    def __init__(self):
        self._repo_missing = False
        self._all_items = []
        self._filter_text = ''

        xaml_file = op.join(script_dir, 'BranchSwitcherWindow.xaml')
        forms.WPFWindow.__init__(self, xaml_file)

        load_styles_to_window(self)

        if not updater.get_extension_dir():
            self._repo_missing = True
            self.currentBranchBadge.Text = 'not a git clone'
            self._set_status(
                'This extension folder is not a git repository. '
                'Branch switching is disabled.'
            )
            self.branchCountText.Text = '0 branches'
            self.branchList.IsEnabled = False
            self.searchBox.IsEnabled = False
            self.switchReloadButton.IsEnabled = False
            return

        self.load_branch_list()
        self.searchBox.Focus()

    def load_branch_list(self):
        cur = updater.get_current_branch()
        self.currentBranchBadge.Text = cur if cur else '(detached)'

        descriptors = updater.list_branches_with_source()

        # Guarantee the current branch is always represented, even if it's a
        # brand-new local branch that isn't on origin yet (would otherwise
        # leave the header badge referencing a row that isn't in the list).
        if cur and not any(d['name'] == cur for d in descriptors):
            descriptors.append({'name': cur, 'local': True, 'remote': False})
            descriptors.sort(key=lambda d: d['name'].lower())

        self._all_items = [
            BranchItem(
                name=d['name'],
                is_current=(d['name'] == cur),
                has_local=d['local'],
                has_remote=d['remote'],
            )
            for d in descriptors
        ]

        if self._all_items:
            # Happy path: primer already explains what the list is. Keep quiet.
            self._set_status('')
        else:
            self._set_status(
                'No branches found. Run git fetch (or reload pyRevit), '
                'then reopen this dialog.'
            )

        self._apply_filter()
        self._select_current()

    def _set_status(self, text):
        """Set the status line; hide it when there's nothing to say."""
        self.statusText.Text = text or ''
        self.statusText.Visibility = (
            Visibility.Visible if text else Visibility.Collapsed
        )

    def _apply_filter(self):
        needle = (self._filter_text or '').strip().lower()
        if not needle:
            filtered = list(self._all_items)
        else:
            filtered = [item for item in self._all_items
                        if needle in item.name.lower()]

        self.branchList.ItemsSource = filtered
        total = len(self._all_items)
        shown = len(filtered)
        if shown == total:
            self.branchCountText.Text = '{} branch{}'.format(
                total, '' if total == 1 else 'es')
        else:
            self.branchCountText.Text = '{} / {} shown'.format(shown, total)

    def _select_current(self):
        items = self.branchList.ItemsSource
        if not items:
            return
        # Prefer current branch as initial selection, otherwise pick first row.
        for idx, item in enumerate(items):
            if item.is_current:
                self.branchList.SelectedIndex = idx
                self.branchList.ScrollIntoView(item)
                return
        self.branchList.SelectedIndex = 0

    def _do_switch(self):
        if self._repo_missing:
            return
        selected = self.branchList.SelectedItem
        if selected is None:
            forms.alert('Select a branch.', title='Select Branch')
            return
        branch_name = selected.name if isinstance(selected, BranchItem) else str(selected)

        if isinstance(selected, BranchItem) and selected.is_current:
            forms.alert(
                'You are already on "{}".'.format(branch_name),
                title='Select Branch')
            return

        ok, message = updater.checkout_branch(branch_name)
        if not ok:
            forms.alert(message, title='Git checkout failed')
            return
        logger.info('Checked out branch {}, reloading pyRevit'.format(branch_name))
        self.Close()
        sessionmgr.reload_pyrevit()
        script.get_results().newsession = sessioninfo.get_session_uuid()

    # --- XAML event handlers -------------------------------------------------

    def SearchBox_TextChanged(self, sender, e):
        self._filter_text = sender.Text or ''
        self._apply_filter()
        # Keep a useful selection after filtering
        if self.branchList.Items.Count > 0 and self.branchList.SelectedIndex < 0:
            self.branchList.SelectedIndex = 0

    def BranchList_SelectionChanged(self, sender, e):
        pass

    def BranchList_MouseDoubleClick(self, sender, e):
        if self.branchList.SelectedItem is not None:
            self._do_switch()

    def BranchList_KeyDown(self, sender, e):
        if e.Key == Key.Enter:
            self._do_switch()
            e.Handled = True

    def SwitchReloadButton_Click(self, sender, e):
        self._do_switch()

    def CancelButton_Click(self, sender, e):
        self.Close()


if __name__ == '__main__':
    try:
        BranchSwitcherWindow().ShowDialog()
    except Exception as err:
        logger.error('Switch branch failed: {}'.format(err))
        forms.alert(str(err), title='Select Branch')
