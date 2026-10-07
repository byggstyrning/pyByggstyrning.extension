"""Show available updates, then update and reload."""
# -*- coding: utf-8 -*-
# pylint: disable=import-error,invalid-name,broad-except

__title__ = 'Reload'
__highlight__ = 'new'
__doc__ = """Show what is new, then update pyByggstyrning and reload."""

import json
import os.path as op

from pyrevit import forms, script
from pyrevit.loader import sessioninfo
from pyrevit.loader import sessionmgr
from System.Windows import (
    GridLength,
    GridUnitType,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    Visibility,
)
from System.Windows.Controls import ColumnDefinition, Grid, TextBlock
from System.Windows.Media import ColorConverter, SolidColorBrush

import extension_updater as updater


def _git_text(args):
    extension_dir = updater.get_extension_dir()
    if not extension_dir:
        return ''
    code, out, err = updater._run_git(args, extension_dir)
    if code != 0:
        return ''
    return (out or '').strip()


def _parse_github_slug(url):
    """Return (owner, repo) for a github remote URL, or ('', '')."""
    if not url:
        return '', ''
    text = url.strip()
    if text.endswith('.git'):
        text = text[:-4]
    marker = 'github.com'
    index = text.lower().find(marker)
    if index < 0:
        return '', ''
    tail = text[index + len(marker):].lstrip(':/')
    parts = [part for part in tail.split('/') if part]
    if len(parts) < 2:
        return '', ''
    return parts[0], parts[1]


def _fetch_json(path):
    """GET a path on https://api.github.com. The host is fixed."""
    allowed = set('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-._/~?=&%')
    if not path or not path.startswith('/repos/'):
        raise ValueError('GitHub path is not allowed')
    cleaned = ''.join(ch for ch in path if ch in allowed)
    if cleaned != path or any(part == '..' for part in path.split('/')):
        raise ValueError('GitHub path is not allowed')
    from System import Uri
    from System.IO import StreamReader
    from System.Net import HttpWebRequest
    request = HttpWebRequest.Create(Uri('https://api.github.com' + cleaned))
    request.Method = 'GET'
    request.AllowAutoRedirect = False
    request.Timeout = 8000
    request.UserAgent = 'pyByggstyrning'
    request.Accept = 'application/vnd.github+json'
    response = request.GetResponse()
    try:
        reader = StreamReader(response.GetResponseStream())
        try:
            raw = reader.ReadToEnd()
        finally:
            reader.Close()
    finally:
        response.Close()
    if not isinstance(raw, str):
        raw = str(raw)
    return json.loads(raw)


def _short_date(iso_text):
    """Turn a GitHub timestamp into a short date such as '7 Oct'."""
    text = (iso_text or '')[:10]
    parts = text.split('-')
    if len(parts) != 3:
        return ''
    try:
        year = int(parts[0])
        month = int(parts[1])
        day = int(parts[2])
    except ValueError:
        return ''
    if month < 1 or month > 12 or day < 1 or day > 31:
        return ''
    months = (
        'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
    )
    label = '{} {}'.format(day, months[month - 1])
    try:
        import time
        this_year = time.localtime().tm_year
    except Exception:
        this_year = year
    if year != this_year:
        label = '{} {}'.format(label, year)
    return label


def _incoming_changes(owner, repo_name, branch, sha):
    """Return (summary, messages) for updates that are not installed yet."""
    if not owner or not repo_name or not branch or not sha:
        return "You're up to date.", []
    slug_ok = set('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-._')
    if any(ch not in slug_ok for ch in owner + repo_name):
        return "You're up to date.", []
    try:
        from urllib import quote
    except ImportError:
        from urllib.parse import quote
    compare_url = '/repos/{}/compare/{}...{}'.format(
        '{}/{}'.format(owner, repo_name),
        quote(sha, safe=''),
        quote(branch, safe=''),
    )
    compared = _fetch_json(compare_url)
    lines = []
    for row in compared.get('commits') or []:
        commit = row.get('commit') or {}
        message = ((commit.get('message') or '').split('\n') or [''])[0].strip()
        if not message:
            continue
        dated = commit.get('committer') or commit.get('author') or {}
        lines.append((_short_date(dated.get('date')), message))
    if not lines:
        return "You're up to date.", []
    return "What's new:", lines


class ReloadStatusWindow(forms.WPFWindow):
    """Short list of updates that are not installed yet."""

    def __init__(self, summary, commits):
        self._do_reload = False
        xaml_file = op.join(op.dirname(__file__), 'ReloadWindow.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        self.summaryText.Text = summary
        self.commitList.Children.Clear()
        if commits:
            self.changeCard.Visibility = Visibility.Visible
            for index, item in enumerate(commits):
                when, message = item
                self.commitList.Children.Add(
                    self._change_row(when, message, index < len(commits) - 1)
                )
        else:
            self.changeCard.Visibility = Visibility.Collapsed

    def _change_row(self, when, message, gap_below):
        row = Grid()
        row.Margin = Thickness(0, 0, 0, 10 if gap_below else 0)
        date_col = ColumnDefinition()
        date_col.Width = GridLength(72)
        text_col = ColumnDefinition()
        text_col.Width = GridLength(1, GridUnitType.Star)
        row.ColumnDefinitions.Add(date_col)
        row.ColumnDefinitions.Add(text_col)
        date = TextBlock()
        date.Text = when or ''
        date.FontSize = 12
        date.Foreground = SolidColorBrush(ColorConverter.ConvertFromString('#8A93A0'))
        date.VerticalAlignment = VerticalAlignment.Top
        date.Margin = Thickness(0, 2, 0, 0)
        Grid.SetColumn(date, 0)
        body = TextBlock()
        body.Text = message
        body.FontSize = 13
        body.TextWrapping = TextWrapping.Wrap
        body.Foreground = SolidColorBrush(ColorConverter.ConvertFromString('#1F2933'))
        Grid.SetColumn(body, 1)
        row.Children.Add(date)
        row.Children.Add(body)
        return row

    def UpdateButton_Click(self, sender, args):
        self._do_reload = True
        self.Close()


def _collect():
    branch = updater.get_current_branch() or ''
    sha = _git_text(['rev-parse', 'HEAD'])
    remote = _git_text(['remote', 'get-url', 'origin'])
    owner, repo_name = _parse_github_slug(remote)
    return {
        'branch': branch,
        'sha': sha,
        'owner': owner,
        'repo': repo_name,
    }


def _pull_and_reload():
    repo = updater.get_repo_info()
    if not repo:
        forms.alert(
            'This extension cannot be updated from here.',
            title='pyByggstyrning',
        )
        return
    success, message = updater.pull_updates(repo)
    if not success:
        forms.alert(message or 'Update failed.', title='pyByggstyrning')
        return
    newsession = sessionmgr.reload_pyrevit()
    try:
        script.get_results().newsession = newsession or sessioninfo.get_session_uuid()
    except Exception:
        pass


def main():
    info = _collect()
    summary = "You're up to date."
    commits = []
    try:
        summary, commits = _incoming_changes(
            info['owner'], info['repo'], info['branch'], info['sha']
        )
    except Exception:
        summary = 'Could not check for updates.'
    window = ReloadStatusWindow(summary, commits)
    window.ShowDialog()
    if window._do_reload:
        _pull_and_reload()


if __name__ == '__main__':
    main()
