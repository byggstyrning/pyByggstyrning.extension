# -*- coding: utf-8 -*-
"""Purge unused content inside loaded families, on a new copy of the model."""

from __future__ import print_function

__title__ = "Family\nPurge"
__author__ = "Byggstyrning AB"
__doc__ = """Purge unused content baked into loaded families.

By default, saves the open model as a new file (a new central if it is
workshared), then purges unused elements inside each loadable family and
its nested families, and loads the result back. The original file is not
overwritten. Uncheck the copy option to purge and save the open model.

Nested types that placed instances still select are kept. A family reload
that would delete elements or a type is rolled back.

Sharing duplicated nested families is optional and off by default. A host
that would lose a type or an instance is rolled back.
"""

import datetime
import os
import os.path as op
import sys
import traceback

import clr
clr.AddReference('RevitAPI')
clr.AddReference('PresentationFramework')
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Visibility
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
output = None
doc = revit.doc


def _bind_output():
    global output
    if output is None:
        output = script.get_output()
    return output


def _reload_purge_libs():
    """Pick up lib edits in an already-open Revit session.

    A failed reload must not stop the button. A new session already imported
    the current files.
    """
    reloader = None
    try:
        reloader = reload
    except NameError:
        try:
            import imp
            reloader = imp.reload
        except Exception:
            return
    try:
        reloader(fp)
        reloader(fs)
    except Exception:
        logger.warning("Could not refresh the purge library:\n{0}".format(
            traceback.format_exc()))


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


class FamilyFilterItem(forms.Reactive):
    """One loadable family in the purge list. Same shape as Family Elevations."""

    def __init__(self, name, category, instance_count=0, is_selected=True):
        super(FamilyFilterItem, self).__init__()
        self._name = name
        self._category = category or ""
        self._instance_count = instance_count or 0
        self._is_selected = is_selected

    @property
    def FamilyName(self):
        return self._name

    @property
    def CategoryName(self):
        return self._category

    @property
    def InstanceCount(self):
        return u"{:,}".format(int(self._instance_count))

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        self.OnPropertyChanged("IsSelected")


class FamilyPurgeWindow(forms.WPFWindow):
    """Explains the purge, picks families, and collects the output folder."""

    def __init__(self, document):
        self.folder = None
        self.save_copy = True
        self.family_names = []
        self.purge_project = True
        self.share_nested = False
        self.model_bytes = 0
        forms.WPFWindow.__init__(
            self, op.join(pushbutton_dir, "FamilyPurgeWindow.xaml"))
        from styles import load_styles_to_window
        load_styles_to_window(self)

        self.document = document
        self.model_bytes = _model_bytes(document)
        self._refresh_model_info()
        self._family_items = ObservableCollection[FamilyFilterItem]()
        self._family_display = ObservableCollection[FamilyFilterItem]()
        self.familyListBox.ItemsSource = self._family_display
        self._populate_families()
        self._sync_family_list()
        self._update_space()

    def _populate_families(self):
        self._family_items.Clear()
        for name, category, count in fp.loadable_families(self.document):
            self._family_items.Add(FamilyFilterItem(name, category, count, True))

    def _sync_family_list(self):
        self._family_display.Clear()
        search = ""
        if hasattr(self, "familySearchTextBox") and self.familySearchTextBox:
            search = (self.familySearchTextBox.Text or "").strip().lower()
        for item in self._family_items:
            haystack = u"{0} {1}".format(item.FamilyName, item.CategoryName).lower()
            if not search or search in haystack:
                self._family_display.Add(item)
        self._update_family_count()

    def _selected_family_names(self):
        names = []
        items = getattr(self, "_family_items", None)
        if items is None:
            return names
        for item in items:
            if item.IsSelected:
                names.append(item.FamilyName)
        return names

    def _update_family_count(self):
        selected = len(self._selected_family_names())
        total = 0
        if getattr(self, "_family_items", None) is not None:
            total = self._family_items.Count
        if hasattr(self, "familySelectionCountText") and self.familySelectionCountText:
            self.familySelectionCountText.Text = "{0} of {1} families included".format(
                selected, total)

    def familySearchTextBox_TextChanged(self, sender, args):
        if getattr(self, "_family_items", None) is not None:
            self._sync_family_list()

    def familySelectAllButton_Click(self, sender, args):
        for item in self._family_items:
            item.IsSelected = True
        self._update_family_count()
        self._update_space()

    def familyDeselectAllButton_Click(self, sender, args):
        for item in self._family_items:
            item.IsSelected = False
        self._update_family_count()
        self._update_space()

    def familyFilterCheckBox_Changed(self, sender, args):
        self._update_family_count()
        self._update_space()

    def _needed_bytes(self):
        return self.model_bytes * 2

    def _wants_copy(self):
        box = getattr(self, "saveCopyCheckBox", None)
        if box is None:
            return True
        return bool(box.IsChecked)

    def _can_save_in_place(self):
        document = self.document
        if document.IsReadOnly:
            return False
        if document.PathName:
            return True
        try:
            return bool(document.IsModelInCloud)
        except Exception:
            return False

    def _refresh_model_info(self):
        document = self.document
        if self._wants_copy():
            if document.IsWorkshared:
                kind = "workshared - will be saved as a new central"
            else:
                kind = "not workshared - will be saved as a copy"
        elif document.IsWorkshared:
            kind = "workshared - will be purged and saved locally"
        else:
            kind = "will be purged and saved in place"
        size = "{:.0f} MB".format(_mb(self.model_bytes)) \
            if self.model_bytes else "size unknown"
        self.modelInfoText.Text = "Model: {0}  ({1}, {2})".format(
            document.Title, size, kind)
        if self._wants_copy():
            self.modelStep1Text.Text = (
                "1. The open model is saved as a new file in the folder you pick. "
                "A workshared model is saved as a new central. The original file "
                "is not overwritten.")
            self.modelStep3Text.Text = (
                "3. When it finishes, the open document is the purged copy and it "
                "is already saved. Close it and reopen your normal model. This "
                "cannot be undone inside the copy.")
        else:
            self.modelStep1Text.Text = (
                "1. The model you have open is purged directly. No separate Revit "
                "file is created. It is saved when the run finishes.")
            self.modelStep3Text.Text = (
                "3. When it finishes, this same model is saved. A workshared model "
                "is saved locally and is not synced to the central. There is no "
                "separate file to fall back on.")

    def _update_space(self):
        save_copy = self._wants_copy()
        self.folderTextBox.IsEnabled = save_copy
        self.browseButton.IsEnabled = save_copy
        self._refresh_model_info()
        selected = len(self._selected_family_names())
        if save_copy:
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
            self.spaceText.Visibility = Visibility.Visible if parts else Visibility.Collapsed
            folder_ok = bool(folder) and op.isdir(folder)
            self.runButton.Content = "Save copy and purge"
            self.runButton.IsEnabled = folder_ok and selected > 0
            if self.runButton.IsEnabled:
                self.statusText.Text = "Ready to purge {0} families".format(selected)
            elif selected == 0:
                self.statusText.Text = "Select at least one family"
            else:
                self.statusText.Text = "Pick an existing output folder to start"
            return
        self.spaceText.Text = ""
        self.spaceText.Visibility = Visibility.Collapsed
        self.runButton.Content = "Purge this model"
        if selected == 0:
            self.runButton.IsEnabled = False
            self.statusText.Text = "Select at least one family"
        elif not self._can_save_in_place():
            self.runButton.IsEnabled = False
            self.statusText.Text = "This model cannot be saved in place. Save a new copy instead."
        else:
            self.runButton.IsEnabled = True
            self.statusText.Text = "Ready to purge {0} families in this model".format(selected)

    def saveCopyCheckBox_Changed(self, sender, args):
        if hasattr(self, "model_bytes"):
            self._update_space()

    def browseButton_Click(self, sender, args):
        picked = forms.pick_folder(title="Output folder for the family purge")
        if picked:
            self.folderTextBox.Text = picked

    def folderTextBox_TextChanged(self, sender, args):
        if hasattr(self, "model_bytes"):
            self._update_space()

    def runButton_Click(self, sender, args):
        self.save_copy = self._wants_copy()
        if self.save_copy:
            needed = self._needed_bytes()
            free = _free_bytes(self.folderTextBox.Text.strip())
            if self.model_bytes and free is not None and free < needed:
                if not forms.alert(
                        "The drive may not have enough free space for the copy. "
                        "Continue anyway?",
                        title="Low disk space", yes=True, no=True):
                    return
            self.folder = self.folderTextBox.Text.strip()
        else:
            if not self._can_save_in_place():
                self.statusText.Text = "This model cannot be saved in place. Save a new copy instead."
                return
            if self.document.IsWorkshared:
                question = (
                    "This purges the open model and saves it locally. "
                    "No separate file is created, and it is not synced to the central.\n\n"
                    "Continue?")
            else:
                question = (
                    "This purges the open model and saves it. "
                    "No separate file is created.\n\n"
                    "Continue?")
            if not forms.alert(question, title="Purge this model", yes=True, no=True):
                return
            self.folder = None
        self.family_names = self._selected_family_names()
        if not self.family_names:
            self.statusText.Text = "Select at least one family"
            return
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


class _BarProgress(object):
    """Drives the Revit progress bar from purge and share callbacks."""

    def __init__(self, bar):
        self.bar = bar

    def __call__(self, current, total, message):
        text = u'?'
        try:
            text = unicode(message)
        except Exception:
            text = u'?'
        safe = text.replace(u'{', u'(').replace(u'}', u')')
        if not total:
            self.bar.indeterminate = True
            self.bar._title = safe
            self.bar.update_progress(0, 1)
            return
        self.bar.indeterminate = False
        self.bar._title = u'{0}   {{value}}/{{max_value}}'.format(safe)
        current_value = int(current)
        total_value = int(total)
        if current_value < 0:
            current_value = 0
        if current_value > total_value:
            current_value = total_value
        self.bar.update_progress(current_value, total_value)


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


def _count_text(value):
    if value is None:
        return u"unknown"
    try:
        return u"{:,}".format(int(value))
    except Exception:
        return unicode(value)


def _mb_text(value):
    if value is None:
        return u"unknown"
    try:
        return u"{:,.1f} MB".format(float(value))
    except Exception:
        return unicode(value)


def _change_lines(before, after, kind):
    """Return (after line, detail, tone) for a before/after value.

    tone is good, bad, warn, or same. Callers color the result from it.
    """
    if before is None or after is None:
        return u"", u"", "same"
    if kind == "size":
        amount = abs(float(before) - float(after))
        after_line = u"to {0}".format(_mb_text(after))
        if amount < 0.05:
            return after_line, u"about the same size", "same"
        if float(before) > float(after):
            return after_line, u"{0} smaller".format(_mb_text(amount)), "good"
        return after_line, u"{0} larger".format(_mb_text(amount)), "bad"
    delta = int(before) - int(after)
    if delta == 0:
        tone = "good" if kind == "instances" else "same"
        return u"unchanged", u"", tone
    if delta > 0:
        tone = "bad" if kind == "instances" else "good"
        return (
            u"to {0}".format(_count_text(after)),
            u"{0} removed".format(_count_text(delta)),
            tone)
    return (
        u"to {0}".format(_count_text(after)),
        u"{0} added".format(_count_text(-delta)),
        "warn")


def _notes_tone(result):
    if result.get('run_error'):
        return "bad"
    before = result.get('instances_before')
    after = result.get('instances_after')
    if before is not None and after is not None and after < before:
        return "bad"
    share = result.get('share') or {}
    if share.get('error') or (share.get('hosts_rejected') or 0):
        return "warn"
    if share.get('instances_added'):
        return "warn"
    return None


def _result_notes(result):
    notes = []
    rows = result.get('families') or []
    failed = sum(1 for row in rows if row.get('error'))
    kept = 0
    for row in rows:
        kept += len(row.get('kept_original_nested') or [])
    before = result.get('instances_before')
    after = result.get('instances_after')
    if before is not None and after is not None and after < before:
        notes.append(
            u"Placed elements went down. If that was not intended, close this "
            u"copy and do not use it.")
    if kept:
        notes.append(
            u"{0} nested families were left as they were, because replacing "
            u"them would have deleted elements or a type.".format(_count_text(kept)))
    if failed:
        if result.get('in_place'):
            notes.append(
                u"{0} families could not be purged. The open model was still "
                u"saved. The log folder lists them.".format(_count_text(failed)))
        else:
            notes.append(
                u"{0} families could not be purged. The copy was still saved. "
                u"The output folder lists them.".format(_count_text(failed)))
    share = result.get('share')
    if share:
        if share.get('error'):
            if result.get('in_place'):
                notes.append(
                    u"Sharing did not finish. The open model was still saved.")
            else:
                notes.append(
                    u"Sharing did not finish. The purged copy was still saved.")
        else:
            added = share.get('instances_added') or 0
            rejected = share.get('hosts_rejected') or 0
            found = len(share.get('candidates') or [])
            if found == 0 and added == 0 and rejected == 0:
                notes.append(
                    u"Sharing was on, but no nested family sat in two or more "
                    u"of the families you checked, so nothing was shared.")
            elif added == 0 and rejected == 0:
                notes.append(
                    u"Sharing found {0} duplicated nested families and did "
                    u"not add elements.".format(_count_text(found)))
            else:
                parts = []
                if added:
                    parts.append(
                        u"Sharing added {0} elements. Each placement of a "
                        u"shared family is now its own element.".format(
                            _count_text(added)))
                else:
                    parts.append(u"Sharing did not add elements.")
                if rejected:
                    parts.append(
                        u"{0} families were left unchanged because sharing "
                        u"them would have deleted a type or an element.".format(
                            _count_text(rejected)))
                notes.append(u" ".join(parts))
    return u"\n\n".join(notes)


class FamilyPurgeResultWindow(forms.WPFWindow):
    """Plain-language summary after a purge. Does not ask a yes/no question."""

    def __init__(self, result, out_dir, before_mb, after_mb):
        self.open_folder = False
        forms.WPFWindow.__init__(
            self, op.join(pushbutton_dir, "FamilyPurgeResultWindow.xaml"))
        from styles import load_styles_to_window
        load_styles_to_window(self)

        size_after, size_detail, size_tone = _change_lines(before_mb, after_mb, "size")
        families_after, family_detail, family_tone = _change_lines(
            result.get('families_before'), result.get('families_after'), "families")
        instances_after, instance_detail, instance_tone = _change_lines(
            result.get('instances_before'), result.get('instances_after'), "instances")

        self.sizeBeforeText.Text = _mb_text(before_mb)
        self.sizeAfterText.Text = size_after
        self.sizeDetailText.Text = size_detail
        self._paint_value(self.sizeBeforeText)
        self._paint_result(self.sizeAfterText, self.sizeDetailText, size_after, size_tone)
        self.familyBeforeText.Text = _count_text(result.get('families_before'))
        self.familyAfterText.Text = families_after
        self.familyDetailText.Text = family_detail
        self._paint_value(self.familyBeforeText)
        self._paint_result(self.familyAfterText, self.familyDetailText, families_after, family_tone)
        self.instanceBeforeText.Text = _count_text(result.get('instances_before'))
        self.instanceAfterText.Text = instances_after
        self.instanceDetailText.Text = instance_detail
        self._paint_value(self.instanceBeforeText)
        self._paint_result(self.instanceAfterText, self.instanceDetailText, instances_after, instance_tone)

        notes = _result_notes(result)
        if result.get('run_error'):
            notes = result.get('run_error')
        if notes:
            self.notesText.Text = notes
            self.notesCard.Visibility = Visibility.Visible
            self._paint(self.notesText, _notes_tone(result))
        else:
            self.notesCard.Visibility = Visibility.Collapsed
        if result.get('run_error'):
            self.Title = "Family purge stopped"
            self.leadText.Text = u"Family purge stopped before it finished."
            if result.get('in_place'):
                self.originalText.Text = (
                    u"The open model was not saved after this stop. "
                    u"Close it without saving if you do not want partial changes.")
            else:
                self.originalText.Text = (
                    u"The original file was not overwritten. "
                    u"If a copy was saved, it may only be partly purged.")
        elif result.get('in_place'):
            self.leadText.Text = u"The open model was purged and saved."
            if result.get('workshared'):
                self.originalText.Text = (
                    u"This is the same file you had open. It was saved locally "
                    u"and was not synced to the central.")
            else:
                self.originalText.Text = (
                    u"This is the same file you had open. It was saved in place, "
                    u"so the file on disk now includes the purge.")
            self.folderText.Text = u"Logs: {0}".format(out_dir or u"")
        else:
            self.folderText.Text = out_dir or u""
        if not out_dir or not op.isdir(out_dir):
            self.openButton.Visibility = Visibility.Collapsed
            if not result.get('in_place'):
                self.folderText.Text = u""
        self.Topmost = True
        self.Activate()

    def _brush(self, name):
        try:
            return self.FindResource(name)
        except Exception:
            return None

    def _paint_value(self, block):
        brush = self._brush("TextBrush")
        if block is not None and brush is not None:
            block.Foreground = brush

    def _paint(self, block, tone):
        names = {
            "good": "SuccessBrush",
            "bad": "ErrorBrush",
            "warn": "WarningBrush",
        }
        name = names.get(tone)
        if not name or block is None:
            return
        brush = self._brush(name)
        if brush is not None:
            block.Foreground = brush

    def _paint_result(self, after_block, detail_block, after_text, tone):
        if after_text == u"unchanged":
            self._paint(after_block, tone)
        else:
            self._paint_value(after_block)
        self._paint(detail_block, tone)

    def openButton_Click(self, sender, args):
        self.open_folder = True
        self.Close()

    def closeButton_Click(self, sender, args):
        self.Close()


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
    if result.get('in_place'):
        output.print_md("**Model:** purged and saved in place")
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


def run_purge(document, window, progress):
    _bind_output()
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = _safe_filename(_model_stem(document))
    if window.save_copy:
        if not window.folder:
            raise Exception(u"Pick an output folder, or uncheck Save a new copy.")
        out_dir = op.join(window.folder, "{0}_family_purge_{1}".format(stem, stamp))
    else:
        temp_root = os.environ.get("TEMP") or os.environ.get("TMP") or u"."
        out_dir = op.join(temp_root, "family_purge", "{0}_{1}".format(stem, stamp))
    os.makedirs(out_dir)
    log_path = op.join(out_dir, "family_purge.jsonl")

    def _step(message):
        progress(0, 0, message)

    output.print_md("## Family purge: {0}".format(document.Title))
    if window.save_copy:
        copy_path = op.join(out_dir, stem + ".rvt")
        _step(u"Saving a new copy. The original file is not overwritten.")
        print("Saving a copy to {0}".format(copy_path))
        before_mb = fp.save_project_copy(document, copy_path)
        print("Saved ({0} MB). The original file was not overwritten.".format(before_mb))
    else:
        before_mb = fp.file_mb(document.PathName or "")
        print("Purging the open model. It will be saved in place when the purge finishes.")
        _step(u"Purging the open model")

    result = None
    after_mb = None
    try:
        fp.quiet_on(HOST_APP.uiapp)
        fp.quiet_strict(True)
        fp.quiet_report(reset=True)
        print("Purging {0} selected families.".format(len(window.family_names)))
        result = fp.purge(
            document,
            log_path=log_path,
            on_status=_status,
            do_project_purge=window.purge_project and not window.share_nested,
            family_names=window.family_names,
            on_progress=progress)
        if window.share_nested:
            try:
                shared_dir = op.join(out_dir, "shared_nested")
                result['share'] = fs.share_duplicated_nested(
                    document,
                    shared_dir,
                    op.join(out_dir, "share_nested.jsonl"),
                    _status,
                    host_names=window.family_names,
                    on_progress=progress)
            except Exception:
                logger.error("Sharing failed:\n{0}".format(traceback.format_exc()))
                result['share'] = {'error': traceback.format_exc()}
                if window.save_copy:
                    print("Sharing failed. The purged copy will still be saved.")
                else:
                    print("Sharing failed. The open model will still be saved.")
            if window.purge_project:
                print("Purging unused items in the project")
                result['project_purge'] = fp.purge_project(
                    document, on_progress=progress)
            counts = fp.begin_counts(document)
            result['instances_after'] = counts['instances']
            result['families_after'] = counts['families']
        if window.save_copy:
            _step(u"Saving the purged copy")
            print("Saving the purged copy (compact).")
        else:
            _step(u"Saving the open model")
            print("Saving the open model (compact).")
        after_mb = fp.save_compact(document)
        result['quiet'] = fp.quiet_report()
        result['before_mb'] = before_mb
        result['after_mb'] = after_mb
        result['output'] = out_dir
        result['in_place'] = not window.save_copy
        result['workshared'] = bool(document.IsWorkshared)
        fp.dump(result, op.join(out_dir, "summary.json"))
    except Exception:
        if window.save_copy:
            try:
                print("Saving the copy after a failure.")
                fp.save_compact(document)
            except Exception:
                logger.warning("Could not save the partial copy:\n{0}".format(
                    traceback.format_exc()))
        else:
            print("The open model was not saved.")
        raise
    finally:
        try:
            fp.quiet_strict(False)
            fp.quiet_off(HOST_APP.uiapp)
        except Exception:
            logger.warning("Could not turn quiet mode off:\n{0}".format(
                traceback.format_exc()))
    return result, out_dir, before_mb, after_mb


def _show_result(result, out_dir, before_mb, after_mb):
    _bind_output()
    try:
        _print_summary(result, out_dir, before_mb, after_mb)
    except Exception:
        logger.error("Could not write the purge summary:\n{0}".format(
            traceback.format_exc()))
    try:
        output.hide()
    except Exception:
        pass
    try:
        finished = FamilyPurgeResultWindow(result, out_dir, before_mb, after_mb)
        finished.ShowDialog()
    except Exception:
        logger.error("Result window could not open:\n{0}".format(
            traceback.format_exc()))
        try:
            output.show()
        except Exception:
            pass
        print("Result window could not open.")
        print(traceback.format_exc())
        return
    if finished.open_folder and out_dir and op.isdir(out_dir):
        os.startfile(out_dir)


def main():
    _reload_purge_libs()
    if doc is None or doc.IsFamilyDocument:
        forms.alert("Open a project model first.", title="Family Purge")
        return
    if not hasattr(doc, "GetUnusedElements"):
        forms.alert(
            "Family Purge uses Revit's unused-element list, which needs "
            "Revit 2024 or newer.",
            title="Family Purge")
        return

    try:
        window = FamilyPurgeWindow(doc)
        opened = window.ShowDialog()
    except Exception:
        logger.error("Family Purge could not open:\n{0}".format(
            traceback.format_exc()))
        forms.alert(
            "Family Purge could not open.\n\n{0}".format(traceback.format_exc()),
            title="Family Purge")
        return
    if not opened:
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
        with forms.ProgressBar(
                title='Family purge', indeterminate=True, height=46) as bar:
            result, out_dir, before_mb, after_mb = run_purge(
                doc, window, _BarProgress(bar))
    except Exception:
        logger.error("Family purge failed:\n{0}".format(traceback.format_exc()))
        reason = traceback.format_exc().strip().splitlines()[-1]
        _show_result({
            'run_error': reason,
            'in_place': not window.save_copy,
            'workshared': bool(doc.IsWorkshared),
            'families': [],
        }, None, None, None)
        return

    _show_result(result, out_dir, before_mb, after_mb)


if __name__ == "__main__":
    main()
