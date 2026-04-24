# -*- coding: utf-8 -*-
"""Extension updater module for pyByggstyrning.

Provides simple git update functionality for the extension.
"""
from __future__ import print_function
import os
import os.path as op
import subprocess


def get_extension_dir():
    """Get the extension root directory.
    
    Returns:
        str: Path to extension directory, or None if not found
    """
    # lib folder is directly under extension root
    lib_dir = op.dirname(op.abspath(__file__))
    extension_dir = op.dirname(lib_dir)
    
    if op.exists(op.join(extension_dir, '.git')):
        return extension_dir
    return None


def get_repo_info():
    """Get repository info for the extension.
    
    Returns:
        RepoInfo object or None
    """
    try:
        from pyrevit.coreutils import git as libgit
        
        extension_dir = get_extension_dir()
        if extension_dir:
            return libgit.get_repo(extension_dir)
        return None
    except Exception:
        return None


def _run_git(args, extension_dir):
    """Run a git command in extension_dir. Returns (returncode, stdout, stderr) as str."""
    popen_kw = dict(
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        stdin=subprocess.PIPE,
        cwd=extension_dir,
        shell=False,
    )
    if os.name == 'nt':
        try:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            si.wShowWindow = getattr(subprocess, 'SW_HIDE', 0)
            popen_kw['startupinfo'] = si
        except Exception:
            pass
        popen_kw['creationflags'] = 0x08000000

    try:
        proc = subprocess.Popen(['git'] + list(args), **popen_kw)
    except TypeError:
        popen_kw.pop('creationflags', None)
        try:
            proc = subprocess.Popen(['git'] + list(args), **popen_kw)
        except Exception as ex:
            return -1, '', str(ex)
    except Exception as ex:
        return -1, '', str(ex)

    try:
        out, err = proc.communicate()
        code = proc.returncode
    except Exception as ex:
        return -1, '', str(ex)

    def _to_str(data):
        if data is None:
            return ''
        if isinstance(data, unicode):  # noqa: F821  # IronPython / Py2
            return data
        try:
            return data.decode('utf-8', 'replace')
        except Exception:
            return str(data)

    return code, _to_str(out), _to_str(err)


def get_current_branch():
    """Return the current branch name, or None if not a repo or detached/unknown."""
    extension_dir = get_extension_dir()
    if not extension_dir:
        return None
    code, out, err = _run_git(['rev-parse', '--abbrev-ref', 'HEAD'], extension_dir)
    if code != 0:
        return None
    name = out.strip()
    if name == 'HEAD':
        return None
    return name


def list_local_branches():
    """Return sorted local branch names (refs/heads)."""
    extension_dir = get_extension_dir()
    if not extension_dir:
        return []
    code, out, err = _run_git(
        ['for-each-ref', 'refs/heads/', '--format=%(refname:short)'],
        extension_dir,
    )
    if code != 0:
        return []
    branches = []
    for line in out.splitlines():
        line = line.strip()
        if line:
            branches.append(line)
    return sorted(branches, key=lambda s: s.lower())


DEFAULT_REMOTE = 'origin'


def list_remote_branches(remote=DEFAULT_REMOTE):
    """Return sorted branch names for ``remote`` using local remote-tracking refs.

    Uses ``refs/remotes/<remote>`` (no network). Run ``git fetch`` to refresh the list.
    """
    extension_dir = get_extension_dir()
    if not extension_dir:
        return []
    remote_prefix = '{}/'.format(remote)
    ref_path = 'refs/remotes/{}/'.format(remote)
    code, out, err = _run_git(
        ['for-each-ref', ref_path, '--format=%(refname:short)'],
        extension_dir,
    )
    if code != 0:
        return []
    branches = []
    for line in out.splitlines():
        line = line.strip()
        if not line or not line.startswith(remote_prefix):
            continue
        name = line[len(remote_prefix):]
        if name == 'HEAD':
            continue
        branches.append(name)
    return sorted(set(branches), key=lambda s: s.lower())


def list_branches_with_source(remote=DEFAULT_REMOTE):
    """Return a combined, sorted list of branch descriptors.

    Each item is a dict:
        {
            'name':   'feature/foo',
            'local':  True|False,   # exists under refs/heads
            'remote': True|False,   # exists under refs/remotes/<remote>
        }

    The union includes branches that exist locally only (not yet pushed)
    as well as branches available on ``remote`` (from the last fetch).
    """
    local = set(list_local_branches())
    remote_names = set(list_remote_branches(remote))
    names = sorted(local | remote_names, key=lambda s: s.lower())
    return [
        {'name': n, 'local': n in local, 'remote': n in remote_names}
        for n in names
    ]


def checkout_branch(branch_name, remote=DEFAULT_REMOTE):
    """Check out a branch that exists locally or on ``remote``.

    Accepts branches from either ``list_local_branches`` or
    ``list_remote_branches`` to avoid blocking checkout of branches
    that were created locally but not yet pushed.

    Returns (success, message).
    """
    extension_dir = get_extension_dir()
    if not extension_dir:
        return False, 'Extension directory is not a git repository.'
    if not branch_name:
        return False, 'No branch selected.'

    local_set = set(list_local_branches())
    remote_set = set(list_remote_branches(remote))
    if branch_name not in local_set and branch_name not in remote_set:
        return False, (
            'Branch "{}" is not known locally or on remote "{}". '
            'Fetch the remote (e.g. pyRevit Reload) and try again.'
        ).format(branch_name, remote)

    code, out, err = _run_git(['checkout', branch_name], extension_dir)
    if code == 0:
        return True, 'OK'

    # Only attempt to pull the branch from the remote when it actually exists there.
    if branch_name in remote_set:
        _run_git(['fetch', remote, branch_name], extension_dir)
        code, out, err = _run_git(['checkout', branch_name], extension_dir)
        if code == 0:
            return True, 'OK'

        ref = '{}/{}'.format(remote, branch_name)
        code, out, err = _run_git(['checkout', '-b', branch_name, ref], extension_dir)
        if code == 0:
            return True, 'OK'

    detail = (err or out or '').strip()
    if not detail:
        detail = 'git checkout failed (exit code {})'.format(code)
    return False, detail


def get_version_string(repo_info):
    """Get a simple version string.
    
    Args:
        repo_info: RepoInfo object
        
    Returns:
        str: Version info like "main @ abc1234"
    """
    if not repo_info:
        return "Unknown"
    try:
        return "{} @ {}".format(
            repo_info.branch,
            repo_info.last_commit_hash[:7]
        )
    except Exception:
        return "Unknown"


def check_for_updates(repo_info):
    """Check if updates are available.
    
    Args:
        repo_info: RepoInfo object
        
    Returns:
        tuple: (has_updates, commits_behind) or (False, 0) on error
    """
    if not repo_info:
        return False, 0
    
    try:
        from pyrevit.coreutils import git as libgit
        
        # Fetch latest from remote
        libgit.git_fetch(repo_info)
        
        # Compare heads
        hist_div = libgit.compare_branch_heads(repo_info)
        if hist_div and hist_div.BehindBy > 0:
            return True, hist_div.BehindBy
        return False, 0
    except Exception:
        return False, 0


def pull_updates(repo_info):
    """Pull latest changes from remote.
    
    Args:
        repo_info: RepoInfo object
        
    Returns:
        tuple: (success, message)
    """
    if not repo_info:
        return False, "No repository found"
    
    try:
        from pyrevit.coreutils import git as libgit
        from pyrevit.compat import safe_strtype
        
        old_hash = repo_info.last_commit_hash[:7]
        updated = libgit.git_pull(repo_info)
        new_hash = updated.last_commit_hash[:7]
        
        if old_hash != new_hash:
            msg = safe_strtype(updated.repo.Head.Tip.Message).split('\n')[0]
            return True, "Updated {} -> {}: {}".format(old_hash, new_hash, msg)
        return True, "Already up-to-date"
        
    except Exception as ex:
        return False, str(ex)
