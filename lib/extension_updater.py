# -*- coding: utf-8 -*-
"""Extension updater module for pyByggstyrning.

Provides simple git update functionality for the extension.
"""
from __future__ import print_function
import os.path as op


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
