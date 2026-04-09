"""Update from Git and reload pyRevit."""
# -*- coding=utf-8 -*-
#pylint: disable=import-error,invalid-name,broad-except

__title__ = 'Reload'
__highlight__ = 'new'

from pyrevit import script
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo

import extension_updater as updater

print("\n🔄 pyByggstyrning - Update & Reload\n")

# Update from git
repo = updater.get_repo_info()
if repo:
    version = updater.get_version_string(repo)
    print("📌 Current: {}".format(version))
    print("🔍 Checking for updates...")
    
    success, message = updater.pull_updates(repo)
    if "Updated" in message:
        print("✅ {}".format(message))
    else:
        print("✅ Already up-to-date")
else:
    print("⚠️ No git repository found")

print("\n🚀 Reloading pyRevit...\n")

sessionmgr.reload_pyrevit()
script.get_results().newsession = sessioninfo.get_session_uuid()
