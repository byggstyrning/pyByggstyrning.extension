# -*- coding: utf-8 -*-
"""Purge unused content inside loaded families, on a new copy of the model."""

__title__ = "Family\nPurge"
__author__ = "Byggstyrning AB"
__doc__ = """Purge unused content baked into loaded families.

Saves the open model as a new file (a new central if it is workshared),
then purges unused elements inside each loadable family and its nested
families, and loads the result back. The original file is not overwritten.

Nested types that placed instances still select are kept. A family reload
that would delete elements or a type is rolled back.

Sharing duplicated nested families is optional and off by default. A host
that would lose a type or an instance is rolled back.
"""

from __future__ import print_function

import datetime
import os
import os.path as op
import sys
import traceback

import clr
clr.AddReference('RevitAPI')
from System.IO import DriveInfo, Path
from Autodesk.Revit.DB import (
    FilteredWorksetCollector,
    ModelPathUtils,
    WorksetKind,
)

from pyrevit import HOST_APP, forms, revit, script

pushbutton_dir = op.dirname(__file__)
stack_dir = op.dirname(pushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from filesize import family_purge as fp
from filesize import family_share as fs

logger = script.get_logger()
output = script.get_output()
doc = revit.doc


def _mb(n):
    return n / (1024.0 * 1024.0)


def _model_bytes(document):
    path = document.PathName or ""
    try:
        if path and op.isfile(path):
            return op.getsize(path)
    except Exception:
        pass
    return 0


def _model_stem(document):
    path = document.PathName or ""
    if document.IsWorkshared:
        try:
            central = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                document.GetWorksharingCentralModelPath())
            if central:
                path = central
        except Exception:
            pass
    if not path:
        return document.Title
    return op.splitext(op.basename(path))[0]


def _safe_filename(name):
    text = name or "model"
    bad = '<>:"/\\|?*'
    chars = []
    for char in text:
        chars.append("_" if char in bad else char)
    cleaned = "".join(chars).strip(" .")
    return (cleaned or "model")[:80]


def _free_bytes(folder):
    try:
        root = Path.GetPathRoot(folder)
        if not root:
            return None
        return DriveInfo(root).AvailableFreeSpace
    except Exception:
        return None


def _closed_user_worksets(document):
    if not document.IsWorkshared:
        return []
    names = []
    for workset in FilteredWorksetCollector(document).OfKind(WorksetKind.UserWorkset):
        if not workset.IsOpen:
            names.append(workset.Name)
    return names


class FamilyPurgeWindow(forms.WPFWindow):
    """Explains the purge and collects the output folder."""

    def __init__(self, document):
        self.folder = None
        self.purge_project = True
        self.share_nested = False
        self.model_bytes = 0
        forms.WPFWindow.__init__(
            self, op.join(pushbutton_dir, "FamilyPurgeWindow.xaml"))
        from styles import load_styles_to_window
        load_styles_to_window(self)

        self.document = document
        self.model_bytes = _model_bytes(document)
        kind = "workshared - will be saved as a new central" \
            if document.IsWorkshared else "not workshared - will be saved as a copy"
        size = "{:.0f} MB".format(_mb(self.model_bytes)) \
            if self.model_bytes else "size unknown"
        self.modelInfoText.Text = "Model: {0}  ({1}, {2})".format(
            document.Title, size, kind)
        self._update_space()

    def _needed_bytes(self):
        return self.model_bytes * 2

    def _update_space(self):
        folder = (self.folderTextBox.Text or "").strip()
        needed = self._needed_bytes()
        parts = []
        if self.model_bytes:
            parts.append("Needs about {0:.1f} GB free for the copy.".format(
                needed / 1024.0 ** 3))
        free = _free_bytes(folder) if folder else None
        if free is not None:
            parts.append("Free on {0}: {1:.1f} GB.".format(
                Path.GetPathRoot(folder), free / 1024.0 ** 3))
            if self.model_bytes and free < needed:
                parts.append("Probably NOT enough space.")
        self.spaceText.Text = "  ".join(parts)
        self.runButton.IsEnabled = bool(folder) and op.isdir(folder)
        self.statusText.Text = "Ready" if self.runButton.IsEnabled \
            else "Pick an existing output folder to start"

    def browseButton_Click(self, sender, args):
        picked = forms.pick_folder(title="Output folder for the family purge")
        if picked:
            self.folderTextBox.Text = picked

    def folderTextBox_TextChanged(self, sender, args):
        if hasattr(self, "model_bytes"):
            self._update_space()

    def runButton_Click(self, sender, args):
        needed = self._needed_bytes()
        free = _free_bytes(self.folderTextBox.Text.strip())
        if self.model_bytes and free is not None and free < needed:
            if not forms.alert(
                    "The drive may not have enough free space for the copy. "
                    "Continue anyway?",
                    title="Low disk space", yes=True, no=True):
                return
        self.folder = self.folderTextBox.Text.strip()
        self.purge_project = bool(self.purgeProjectCheckBox.IsChecked)
        self.share_nested = bool(self.shareNestedCheckBox.IsChecked)
        self.DialogResult = True
        self.Close()

    def cancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()


def _status(message):
    text = message if isinstance(message, basestring) else str(message)
    print(text)


def _print_share(share):
    if not share:
        return
    if share.get('error'):
        output.print_md("**Sharing failed.** The purged copy was still saved.")
        print(share.get('error'))
        return
    output.print_md(
        "**Sharing:** {0} duplicated families, {1} exported, "
        "{2} hosts updated, {3} hosts rolled back, "
        "{4} nested loads rolled back, instances {5}.".format(
            len(share.get('candidates') or []),
            len(share.get('exported') or []),
            share.get('hosts_updated'),
            share.get('hosts_rejected'),
            share.get('nested_rolled_back'),
            "{0:+d}".format(share.get('instances_added') or 0)))
    delta = share.get('category_delta') or {}
    if delta:
        output.print_md("**Elements added or removed by sharing:**")
        for name, count in sorted(delta.items(), key=lambda item: -item[1]):
            print(u"{0:+d}  {1}".format(count, name))
    rejected = []
    for row in share.get('hosts') or []:
        if row.get('host_rolled_back') or row.get('rolled_back'):
            rejected.append(row)
    if not rejected:
        return
    table = []
    for row in rejected[:30]:
        reasons = []
        if row.get('host_rollback_reason'):
            reasons.append(row.get('host_rollback_reason'))
        for item in row.get('rolled_back') or []:
            reasons.append(u"{0}: {1}".format(item.get('name'), item.get('reason')))
        text = u"; ".join(reasons)
        if len(text) > 120:
            text = text[:117] + u"..."
        table.append([row.get('host') or "", text])
    output.print_table(
        table,
        title="Sharing rolled back",
        columns=["Host", "Why"])


def _print_summary(result, out_dir, before_mb, after_mb):
    rows = result.get('families') or []
    changed = []
    failed = 0
    unchanged = 0
    kept_names = []
    for row in rows:
        if row.get('error'):
            failed += 1
        if row.get('unchanged') and not row.get('error'):
            unchanged += 1
        else:
            changed.append(row)
        for name in row.get('kept_original_nested') or []:
            kept_names.append(u"{0}  (inside {1})".format(name, row.get('name')))

    output.print_md("**Output folder:** `{0}`".format(out_dir))
    output.print_md(
        "**File size:** {0} MB -> {1} MB".format(
            before_mb if before_mb is not None else "?",
            after_mb if after_mb is not None else "?"))
    output.print_md(
        "**Instances:** {0} -> {1}".format(
            result.get('instances_before'), result.get('instances_after')))
    output.print_md(
        "**Families:** {0} -> {1}".format(
            result.get('families_before'), result.get('families_after')))
    output.print_md(
        "**Processed:** {0}  changed {1}  unchanged {2}  failed {3}".format(
            result.get('family_total'), len(changed), unchanged, failed))

    table = []
    changed.sort(key=lambda row: -(row.get('purged') or 0))
    for row in changed[:40]:
        error = row.get('error') or ""
        if len(error) > 80:
            error = error[:77] + "..."
        table.append([
            row.get('name') or "",
            str(row.get('purged') or 0),
            str(row.get('reloaded') or 0),
            str(len(row.get('kept_original_nested') or [])),
            str(row.get('sec') or ""),
            error,
        ])
    if table:
        output.print_table(
            table,
            title="Families that changed or failed",
            columns=["Family", "Purged", "Reloaded", "Kept", "Sec", "Error"])
        if len(changed) > 40:
            print("{0} more rows are in family_purge.jsonl".format(len(changed) - 40))
    if kept_names:
        output.print_md("**Nested families kept because a reload would have deleted elements:**")
        for name in kept_names[:30]:
            print(name)
    duplicates = result.get('duplicate_family_names') or []
    if duplicates:
        output.print_md("**Duplicate family names (only the first was purged):**")
        for name in duplicates[:20]:
            print(name)
    _print_share(result.get('share'))
    project_purge = result.get('project_purge')
    if project_purge:
        output.print_md("**Project purge:** {0}".format(
            ", ".join("{0} {1}".format(count, kind)
                      for kind, count in sorted(project_purge.items()))))
    quiet = result.get('quiet')
    if quiet:
        output.print_md("**Warnings swallowed during the run:**")
        print(quiet)


def _alert_text(result, before_mb, after_mb):
    rows = result.get('families') or []
    failed = sum(1 for row in rows if row.get('error'))
    lines = [
        "Family purge finished.",
        "",
        "{0} MB -> {1} MB.".format(
            before_mb if before_mb is not None else "?",
            after_mb if after_mb is not None else "?"),
        "Instances: {0} -> {1}.".format(
            result.get('instances_before'), result.get('instances_after')),
        "Families: {0} -> {1}.".format(
            result.get('families_before'), result.get('families_after')),
        "",
        "The open document is the purged copy, already saved.",
        "The original file was not overwritten.",
    ]
    before = result.get('instances_before')
    after = result.get('instances_after')
    if before is not None and after is not None and after < before:
        lines.append("")
        lines.append(
            "Instance count went down. Discard this copy if that was not intended.")
    kept = 0
    for row in rows:
        kept += len(row.get('kept_original_nested') or [])
    if kept:
        lines.append("")
        lines.append(
            "{0} nested families were kept because reloading them would "
            "have deleted elements.".format(kept))
    share = result.get('share') or {}
    if share.get('error'):
        lines.append("")
        lines.append("Sharing failed. The purged copy was still saved.")
    elif share:
        lines.append("")
        lines.append(
            "Sharing added {0} instances. {1} hosts were rolled back.".format(
                share.get('instances_added'), share.get('hosts_rejected')))
    if failed:
        lines.append("")
        lines.append("{0} families failed. See the output window.".format(failed))
    lines.append("")
    lines.append("Open the output folder?")
    return "\n".join(lines)


def run_purge(document, window):
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = _safe_filename(_model_stem(document))
    out_dir = op.join(window.folder, "{0}_family_purge_{1}".format(stem, stamp))
    os.makedirs(out_dir)
    copy_path = op.join(out_dir, stem + ".rvt")
    log_path = op.join(out_dir, "family_purge.jsonl")

    output.print_md("## Family purge: {0}".format(document.Title))
    print("Saving a copy to {0}".format(copy_path))
    before_mb = fp.save_project_copy(document, copy_path)
    print("Saved ({0} MB). The original file was not overwritten.".format(before_mb))

    result = None
    after_mb = None
    try:
        fp.quiet_on(HOST_APP.uiapp)
        fp.quiet_strict(True)
        fp.quiet_report(reset=True)
        result = fp.purge(
            document,
            log_path=log_path,
            on_status=_status,
            do_project_purge=window.purge_project and not window.share_nested)
        if window.share_nested:
            try:
                shared_dir = op.join(out_dir, "shared_nested")
                result['share'] = fs.share_duplicated_nested(
                    document,
                    shared_dir,
                    op.join(out_dir, "share_nested.jsonl"),
                    _status)
            except Exception:
                logger.error("Sharing failed:\n{0}".format(traceback.format_exc()))
                result['share'] = {'error': traceback.format_exc()}
                print("Sharing failed. The purged copy will still be saved.")
            if window.purge_project:
                print("Purging unused items in the project")
                result['project_purge'] = fp.purge_project(document)
            counts = fp.begin_counts(document)
            result['instances_after'] = counts['instances']
            result['families_after'] = counts['families']
        print("Saving the purged copy (compact).")
        after_mb = fp.save_compact(document)
        result['quiet'] = fp.quiet_report()
        result['before_mb'] = before_mb
        result['after_mb'] = after_mb
        result['output'] = out_dir
        fp.dump(result, op.join(out_dir, "summary.json"))
    except Exception:
        try:
            print("Saving the copy after a failure.")
            fp.save_compact(document)
        except Exception:
            logger.warning("Could not save the partial copy:\n{0}".format(
                traceback.format_exc()))
        raise
    finally:
        try:
            fp.quiet_strict(False)
            fp.quiet_off(HOST_APP.uiapp)
        except Exception:
            logger.warning("Could not turn quiet mode off:\n{0}".format(
                traceback.format_exc()))
    return result, out_dir, before_mb, after_mb


def main():
    if doc is None or doc.IsFamilyDocument:
        forms.alert("Open a project model first.", title="Family Purge")
        return
    if not hasattr(doc, "GetUnusedElements"):
        forms.alert(
            "Family Purge uses Revit's unused-element list, which needs "
            "Revit 2024 or newer.",
            title="Family Purge")
        return

    window = FamilyPurgeWindow(doc)
    if not window.ShowDialog():
        return

    closed = _closed_user_worksets(doc)
    if closed:
        shown = "\n".join(closed[:15]) + ("\n..." if len(closed) > 15 else "")
        if not forms.alert(
                "{0} workset(s) are closed. Types used only on those "
                "worksets can be purged:\n\n{1}\n\nContinue anyway?".format(
                    len(closed), shown),
                title="Closed worksets", yes=True, no=True):
            return

    try:
        result, out_dir, before_mb, after_mb = run_purge(doc, window)
    except Exception:
        logger.error("Family purge failed:\n{0}".format(traceback.format_exc()))
        forms.alert(
            "Family purge failed. The original file was not overwritten. "
            "If a copy was saved, it may only be partly purged.\n\n"
            "See the output window.",
            title="Family Purge")
        return

    _print_summary(result, out_dir, before_mb, after_mb)
    if forms.alert(
            _alert_text(result, before_mb, after_mb),
            title="Family Purge", yes=True, no=True):
        os.startfile(out_dir)


if __name__ == "__main__":
    main()
