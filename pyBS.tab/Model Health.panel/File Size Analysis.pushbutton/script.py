# -*- coding: utf-8 -*-
"""Onion-peel file-size analysis of the open model.

Saves the model as a new central (or a plain copy) in a folder the user
picks, then strips it step by step with compact Save As snapshots to show
where the megabytes live.
"""

__title__ = "File Size\nAnalysis"
__author__ = "Byggstyrning AB"
__doc__ = """Find out where the megabytes in this model live.

Saves the open model as a NEW central (or a plain copy if it is not
workshared) in a folder you pick, then removes content step by step
(purge, links, CAD, views, sheets, families, phases, ...) and measures
the compacted file size after each step.

The original central is never touched. Progress and the size table go
to the pyRevit output window; report.html, CSV, Markdown and JSON land
in the output folder.
"""

# Standard library imports
import datetime
import os
import os.path as op
import re
import sys
import traceback

# .NET imports
import clr
clr.AddReference('RevitAPI')
from System.IO import DriveInfo, Path
from Autodesk.Revit.DB import (
    FilteredWorksetCollector,
    ModelPathUtils,
    SaveAsOptions,
    WorksetKind,
    WorksharingSaveAsOptions,
)

# pyRevit imports
from pyrevit import HOST_APP, forms, revit, script

# Path setup for lib imports
pushbutton_dir = op.dirname(__file__)
panel_dir = op.dirname(pushbutton_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from filesize import analyze_filesize as af

logger = script.get_logger()
output = script.get_output()

doc = revit.doc

# RBP_* settings the analysis reads from the environment. Every one is set
# explicitly per run and restored afterwards, so values never leak between
# runs in the same Revit session.
ENV_KEYS = (
    "RBP_SIZE_OUT_DIR",
    "RBP_KEEP_STEP_FILES",
    "RBP_MEASURE_FAMILY_SIZES",
    "RBP_PURIFY_FAMILIES",
    "RBP_PURIFY_FAMILY_MAX",
    "RBP_PURIFY_FAMILY_OFFSET",
    "RBP_FAMILY_PURIFY_ONLY",
    "RBP_FAMILY_PURIFY_AGGRESSIVE",
    "RBP_FAMILY_ANALYTICS_ONLY",
    "RBP_FAMILY_EXPORT_DIR",
    "RBP_FAMILY_RANKING_CSV",
    "RBP_SIDECAR_PATH",
)


def _model_bytes(document):
    """Size of the open model file on disk, or 0 when unknown (cloud)."""
    path = document.PathName or ""
    try:
        if path and op.isfile(path):
            return op.getsize(path)
    except Exception:
        pass
    return 0


def _model_stem(document):
    """File name without extension; the central's name for local files."""
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


def _mb(n):
    return n / (1024.0 * 1024.0)


def _free_bytes(folder):
    """Free bytes on the drive holding folder, or None if unknown."""
    try:
        root = Path.GetPathRoot(folder)
        if not root:
            return None
        return DriveInfo(root).AvailableFreeSpace
    except Exception:
        return None


def _closed_user_worksets(document):
    """Names of user worksets that are not open in this session."""
    if not document.IsWorkshared:
        return []
    names = []
    for ws in FilteredWorksetCollector(document).OfKind(WorksetKind.UserWorkset):
        if not ws.IsOpen:
            names.append(ws.Name)
    return names


class FileSizeAnalysisWindow(forms.WPFWindow):
    """Explains the analysis and collects the output folder and options."""

    def __init__(self, document):
        forms.WPFWindow.__init__(
            self, op.join(pushbutton_dir, "FileSizeAnalysisWindow.xaml"))
        from styles import load_styles_to_window
        load_styles_to_window(self)

        self.document = document
        self.model_bytes = _model_bytes(document)
        self.n_phases = document.Phases.Size
        self.folder = None
        self.keep_snapshots = True
        self.measure_families = True

        kind = "workshared - will be saved as a new central" \
            if document.IsWorkshared else "not workshared - will be saved as a copy"
        size = "{:.0f} MB".format(_mb(self.model_bytes)) \
            if self.model_bytes else "size unknown"
        self.modelInfoText.Text = "Model: {}  ({}, {}, {} phases)".format(
            document.Title, size, kind, self.n_phases)
        self._update_space()

    def _needed_bytes(self):
        """Upper bound of disk use: the copy plus every snapshot."""
        count = af._expected_snapshot_count(self.n_phases)
        if not self.keepSnapshotsCheckBox.IsChecked:
            count = 2
        return self.model_bytes * (count + 1), count

    def _update_space(self):
        folder = (self.folderTextBox.Text or "").strip()
        needed, count = self._needed_bytes()
        parts = []
        if self.model_bytes:
            parts.append("Needs up to {:.1f} GB ({} snapshots + the copy).".format(
                needed / 1024.0 ** 3, count))
        free = _free_bytes(folder) if folder else None
        if free is not None:
            parts.append("Free on {}: {:.1f} GB.".format(
                Path.GetPathRoot(folder), free / 1024.0 ** 3))
            if self.model_bytes and free < needed:
                parts.append("Probably NOT enough space.")
        self.spaceText.Text = "  ".join(parts)
        self.runButton.IsEnabled = bool(folder) and op.isdir(folder)
        self.statusText.Text = "Ready" if self.runButton.IsEnabled \
            else "Pick an existing output folder to start"

    def browseButton_Click(self, sender, args):
        picked = forms.pick_folder(title="Output folder for the file-size analysis")
        if picked:
            self.folderTextBox.Text = picked

    def folderTextBox_TextChanged(self, sender, args):
        if hasattr(self, "model_bytes"):
            self._update_space()

    def option_Changed(self, sender, args):
        # Checked fires while the XAML loads, before __init__ has run.
        if hasattr(self, "model_bytes"):
            self._update_space()

    def runButton_Click(self, sender, args):
        needed, _ = self._needed_bytes()
        free = _free_bytes(self.folderTextBox.Text.strip())
        if self.model_bytes and free is not None and free < needed:
            if not forms.alert(
                    "The drive may not have enough free space for all "
                    "snapshots. Continue anyway?",
                    title="Low disk space", yes=True, no=True):
                return
        self.folder = self.folderTextBox.Text.strip()
        self.keep_snapshots = bool(self.keepSnapshotsCheckBox.IsChecked)
        self.measure_families = bool(self.measureFamiliesCheckBox.IsChecked)
        self.DialogResult = True
        self.Close()

    def cancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()


def save_as_analysis_copy(document, path):
    """Save the open model to path: new central if workshared, else a copy."""
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = False
    opts.MaximumBackups = 1
    if document.IsWorkshared:
        ws_opts = WorksharingSaveAsOptions()
        ws_opts.SaveAsCentral = True
        opts.SetWorksharingOptions(ws_opts)
    document.SaveAs(path, opts)


# "[RBP-SIZE] [07] Remove views except keep 3D..." marks the start of a step.
STEP_START_RE = re.compile(r"\[(\d+)\] .*\.\.\.$")

# Total step count for the output window progress bar; set per run.
_progress = {"total": 0}


def _log_line(msg):
    """Output sink for the analysis log: pyRevit output window + logger."""
    text = msg if isinstance(msg, basestring) else str(msg)
    match = STEP_START_RE.search(text)
    if match and _progress["total"]:
        output.update_progress(int(match.group(1)), _progress["total"])
    if "FATAL" in text or text.startswith("Traceback"):
        logger.error(text)
    elif "WARN" in text:
        logger.warning(text)
    else:
        print(text)


def _set_env(settings):
    """Apply RBP_* settings; return previous values for restoring."""
    previous = {}
    for key in ENV_KEYS:
        previous[key] = os.environ.get(key)
        value = settings.get(key, "")
        if value:
            os.environ[key] = value
        elif key in os.environ:
            del os.environ[key]
    return previous


def _restore_env(previous):
    for key, value in previous.items():
        if value is None:
            if key in os.environ:
                del os.environ[key]
        else:
            os.environ[key] = value


def _chart_title(chart, text):
    chart.options.title = {"display": True, "text": text, "fontSize": 14}


def _draw_charts(steps):
    """Chart.js charts in the output window: size curve and biggest drops."""
    labels = ["{:02d} {}".format(s.get("index") or 0, s.get("title") or "")
              for s in steps]

    size_chart = output.make_line_chart()
    _chart_title(size_chart, "File size after each step (MB)")
    size_chart.data.labels = labels
    ds = size_chart.data.new_dataset("MB")
    ds.data = [round(s.get("mb") or 0, 2) for s in steps]
    ds.set_color(0, 120, 212, 0.25)
    size_chart.options.legend = {"display": False}
    size_chart.draw()

    drops = [(s.get("title") or "", -(s.get("delta_mb") or 0))
             for s in steps if (s.get("delta_mb") or 0) < 0]
    drops.sort(key=lambda d: d[1], reverse=True)
    drops = [d for d in drops if d[1] >= 0.05][:12]
    if not drops:
        return

    drop_chart = output.make_bar_chart()
    _chart_title(drop_chart, "Size removed per step (MB)")
    drop_chart.data.labels = [d[0] for d in drops]
    ds = drop_chart.data.new_dataset("MB removed")
    ds.data = [round(d[1], 2) for d in drops]
    ds.set_color(255, 187, 0, 0.8)
    drop_chart.options.legend = {"display": False}
    drop_chart.draw()


def print_results(payload, out_dir):
    """Print the step table and report links to the output window."""
    steps = (payload or {}).get("steps") or []
    rows = []
    for s in steps:
        delta = s.get("delta_mb")
        rows.append([
            "{:02d}".format(s.get("index") or 0),
            s.get("title") or "",
            "{:.1f}".format(s.get("mb") or 0),
            "" if delta is None else "{:+.1f}".format(delta),
            "{}%".format(s.get("pct_of_original") or 0),
            str(s.get("deleted") or 0),
        ])
    if rows:
        output.print_table(
            rows,
            title="File size after each step",
            columns=["#", "Step", "MB", "Delta MB", "% of original", "Deleted"])
    if steps:
        try:
            _draw_charts(steps)
        except Exception as ex:
            logger.warning("Could not draw charts: {}".format(ex))
    report = op.join(out_dir, "report.html")
    output.print_md("**Output folder:** `{}`".format(out_dir))
    if op.isfile(report):
        output.print_html('<a href="file:///{0}">Open report.html</a>'.format(
            report.replace("\\", "/")))


def run_analysis(document, window):
    """Save the analysis copy and run every peel step on it."""
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = af._safe_filename(_model_stem(document))
    out_dir = op.join(window.folder, "{}_filesize_{}".format(stem, stamp))
    os.makedirs(out_dir)
    copy_path = op.join(out_dir, stem + ".rvt")

    output.print_md("## File size analysis: {}".format(document.Title))
    print("Saving analysis copy to {}".format(copy_path))
    save_as_analysis_copy(document, copy_path)
    print("Saved. Revit now works on the copy; the original is untouched.")

    previous = _set_env({
        "RBP_SIZE_OUT_DIR": out_dir,
        "RBP_KEEP_STEP_FILES": "1" if window.keep_snapshots else "0",
        "RBP_MEASURE_FAMILY_SIZES": "1" if window.measure_families else "0",
        # Family purification reloads every family and inflates the model
        # more than it saves, which skews every later step - never run it.
        "RBP_PURIFY_FAMILIES": "0",
    })
    _progress["total"] = af._expected_snapshot_count(document.Phases.Size)
    output.update_progress(0, _progress["total"])
    af.configure_host(HOST_APP.uiapp, "pyrevit-" + stamp)
    af.set_output_sink(_log_line)
    try:
        af.log("Article: {}".format(af.SOURCE_ARTICLE))
        payload = af.run(document, copy_path)
    finally:
        output.update_progress(_progress["total"], _progress["total"])
        _progress["total"] = 0
        af.set_output_sink(None)
        _restore_env(previous)
    print_results(payload, out_dir)
    return out_dir


def main():
    if doc is None or doc.IsFamilyDocument:
        forms.alert("Open a project model first.", title="File Size Analysis")
        return

    window = FileSizeAnalysisWindow(doc)
    if not window.ShowDialog():
        return

    closed = _closed_user_worksets(doc)
    if closed:
        shown = "\n".join(closed[:15]) + ("\n..." if len(closed) > 15 else "")
        if not forms.alert(
                "{} workset(s) are closed and their elements will not be "
                "measured:\n\n{}\n\nContinue anyway?".format(len(closed), shown),
                title="Closed worksets", yes=True, no=True):
            return

    try:
        out_dir = run_analysis(doc, window)
    except Exception:
        logger.error("File size analysis failed:\n{}".format(
            traceback.format_exc()))
        forms.alert("The analysis failed - see the output window.",
                    title="File Size Analysis")
        return

    report = op.join(out_dir, "report.html")
    if forms.alert(
            "Analysis complete.\n\nThe open document is now the last, "
            "stripped snapshot - close it without saving.\n\n"
            "Open report.html now?",
            title="File Size Analysis", yes=True, no=True) \
            and op.isfile(report):
        os.startfile(report)


if __name__ == '__main__':
    main()
