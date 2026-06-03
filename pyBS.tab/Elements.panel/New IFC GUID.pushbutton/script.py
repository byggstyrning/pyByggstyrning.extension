# -*- coding: utf-8 -*-
"""Assign new IFC GUIDs to selected elements or the whole model."""

__title__ = "New IFC\nGUID"
__author__ = "Byggstyrning AB"
__doc__ = """Assign new IfcGUID values for IFC export.

Modes:
- Current selection (model instances)
- Whole model (instances and types)

Uses ExporterIFCUtils.CreateGUID and AddValueString on IFC_GUID / IFC_TYPE_GUID."""

import sys
import os.path as op
import time

import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIIFC')

from System import Action
from System.Windows import Visibility
from System.Windows.Threading import DispatcherPriority
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Transaction,
    Wall,
)
from Autodesk.Revit.DB.IFC import ExporterIFCUtils

from pyrevit import forms, revit, script

MODE_SELECTION = 'selection'
MODE_WHOLE_MODEL = 'whole_model'
MAX_UNIQUE_ATTEMPTS = 25
PROGRESS_EVERY = 25

logger = script.get_logger()

_SCRIPT_DIR = op.dirname(__file__)
_EXTENSION_DIR = op.dirname(op.dirname(op.dirname(_SCRIPT_DIR)))
_LIB_PATH = op.join(_EXTENSION_DIR, 'lib')
if _LIB_PATH not in sys.path:
    sys.path.insert(0, _LIB_PATH)


def _pump_ui(window):
    """Allow WPF progress UI to refresh during long Revit work."""
    if window is None or window.Dispatcher is None:
        return
    try:
        window.Dispatcher.Invoke(DispatcherPriority.Background, Action(lambda: None))
    except Exception:
        pass


def _element_label(element):
    try:
        cat = element.Category.Name if element.Category else 'No category'
    except Exception:
        cat = 'No category'
    kind = 'Type' if element.IsElementType else 'Instance'
    try:
        from revit.compat import get_element_id_value
        eid = get_element_id_value(element.Id)
    except Exception:
        eid = element.Id.IntegerValue
    return '{0} | Id {1} | {2}'.format(kind, eid, cat)


def can_store_ifc_guid(element):
    """Mirror exporter guard: curtain walls cannot store IfcGUID safely."""
    try:
        if isinstance(element, Wall):
            wall = element
            if wall.CurtainGrid is not None:
                return False
    except Exception:
        pass
    return True


def get_ifc_guid_param_id(element):
    if element.IsElementType:
        return ElementId(BuiltInParameter.IFC_TYPE_GUID)
    return ElementId(BuiltInParameter.IFC_GUID)


def generate_unique_ifc_guid(assigned_guids):
    """Return a new IFC GlobalId string not yet used in this run."""
    for _ in range(MAX_UNIQUE_ATTEMPTS):
        candidate = ExporterIFCUtils.CreateGUID()
        if candidate and candidate not in assigned_guids:
            return candidate
    return None


def assign_ifc_guid(element, assigned_guids):
    """Assign one element. Returns (status, message). status: updated|skipped|failed."""
    if not can_store_ifc_guid(element):
        return 'skipped', 'Curtain wall (IfcGUID not stored by exporter)'

    new_guid = generate_unique_ifc_guid(assigned_guids)
    if not new_guid:
        return 'failed', 'Could not generate a unique IFC GUID'

    param_id = get_ifc_guid_param_id(element)
    try:
        ExporterIFCUtils.AddValueString(element, param_id, new_guid)
    except Exception as ex:
        return 'failed', str(ex)

    assigned_guids.add(new_guid)
    return 'updated', new_guid


def collect_elements(doc, mode):
    """Collect elements to process for the chosen mode."""
    elements = []
    if mode == MODE_SELECTION:
        selection = revit.get_selection()
        if not selection or not selection.element_ids:
            return elements
        for element_id in selection.element_ids:
            element = doc.GetElement(element_id)
            if element is None or element.IsElementType:
                continue
            elements.append(element)
        return elements

    for element in FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
        if element is not None:
            elements.append(element)
    for element in FilteredElementCollector(doc).WhereElementIsElementType().ToElements():
        if element is not None:
            elements.append(element)
    return elements


def run_ifc_guid_assignment(doc, elements, progress_callback=None):
    """Run assignment inside an open transaction. Returns result dict."""
    started = time.time()
    assigned_guids = set()
    counts = {'updated': 0, 'skipped': 0, 'failed': 0}
    failure_lines = []
    skip_lines = []
    total = len(elements)

    with Transaction(doc, 'New IFC GUID') as txn:
        txn.Start()
        for index, element in enumerate(elements):
            status, message = assign_ifc_guid(element, assigned_guids)
            counts[status] = counts.get(status, 0) + 1
            if status == 'failed':
                line = '{0} — {1}'.format(_element_label(element), message)
                failure_lines.append(line)
                logger.warning(line)
            elif status == 'skipped' and len(skip_lines) < 200:
                skip_lines.append('{0} — {1}'.format(_element_label(element), message))

            if progress_callback and (
                index == 0 or (index + 1) % PROGRESS_EVERY == 0 or index + 1 == total
            ):
                progress_callback(index + 1, total)

        txn.Commit()

    elapsed = time.time() - started
    return {
        'total': total,
        'updated': counts.get('updated', 0),
        'skipped': counts.get('skipped', 0),
        'failed': counts.get('failed', 0),
        'failure_lines': failure_lines,
        'skip_lines': skip_lines,
        'elapsed_seconds': elapsed,
    }


class NewIfcGuidWindow(forms.WPFWindow):
    """Mode picker and progress UI."""

    def __init__(self):
        xaml_path = op.join(_SCRIPT_DIR, 'NewIfcGuidWindow.xaml')
        forms.WPFWindow.__init__(self, xaml_path)
        from styles import load_styles_to_window
        load_styles_to_window(self)
        self._result = None
        self._mode_label = ''
        self.busyOverlay.Visibility = Visibility.Collapsed

    def _selected_mode(self):
        if self.wholeModelModeRadio.IsChecked:
            return MODE_WHOLE_MODEL
        return MODE_SELECTION

    def _set_busy(self, visible):
        self.busyOverlay.Visibility = Visibility.Visible if visible else Visibility.Collapsed
        self.generateButton.IsEnabled = not visible
        self.cancelButton.IsEnabled = not visible
        self.selectionModeRadio.IsEnabled = not visible
        self.wholeModelModeRadio.IsEnabled = not visible

    def _update_progress(self, current, total):
        if total <= 0:
            self.progressBar.IsIndeterminate = True
            self.progressTextBlock.Text = 'Working...'
        else:
            self.progressBar.IsIndeterminate = False
            self.progressBar.Maximum = total
            self.progressBar.Value = current
            self.progressTextBlock.Text = 'Processing {0} of {1}...'.format(current, total)
        _pump_ui(self)

    def generateButton_Click(self, sender, args):
        doc = revit.doc
        mode = self._selected_mode()
        if mode == MODE_SELECTION:
            self._mode_label = 'Current selection'
        else:
            self._mode_label = 'Whole model'

        elements = collect_elements(doc, mode)
        if not elements:
            if mode == MODE_SELECTION:
                forms.alert(
                    'No model instances in the current selection.\n\n'
                    'Select elements or use Whole model mode.',
                    title='New IFC GUID',
                )
            else:
                forms.alert('No elements found in the model.', title='New IFC GUID')
            return

        self._set_busy(True)
        self._update_progress(0, len(elements))
        try:
            def on_progress(current, total):
                self._update_progress(current, total)

            self._result = run_ifc_guid_assignment(doc, elements, on_progress)
            self._result['mode_label'] = self._mode_label
            self.DialogResult = True
            self.Close()
        except Exception as ex:
            logger.error('IFC GUID assignment failed: {0}'.format(ex))
            forms.alert(
                'IFC GUID assignment failed:\n\n{0}'.format(ex),
                title='New IFC GUID',
            )
            self._set_busy(False)
        finally:
            self._update_progress(len(elements), len(elements))

    def cancelButton_Click(self, sender, args):
        self.DialogResult = False
        self.Close()


class NewIfcGuidResultWindow(forms.WPFWindow):
    """Styled summary after assignment."""

    def __init__(self, result):
        xaml_path = op.join(_SCRIPT_DIR, 'NewIfcGuidResultWindow.xaml')
        forms.WPFWindow.__init__(self, xaml_path)
        from styles import load_styles_to_window
        load_styles_to_window(self)
        self._bind_result(result)

    def _bind_result(self, result):
        updated = result.get('updated', 0)
        skipped = result.get('skipped', 0)
        failed = result.get('failed', 0)
        total = result.get('total', 0)
        elapsed = result.get('elapsed_seconds', 0.0)
        mode_label = result.get('mode_label', '')

        self.updatedCountText.Text = str(updated)
        self.skippedCountText.Text = str(skipped)
        self.failedCountText.Text = str(failed)

        self.summaryLineText.Text = (
            '{0} — processed {1} element(s) in {2:.1f} s. '
            'See pyRevit log for full detail.'.format(mode_label, total, elapsed)
        )

        lines = []
        lines.append('Mode: {0}'.format(mode_label))
        lines.append('Total processed: {0}'.format(total))
        lines.append('Updated: {0}'.format(updated))
        lines.append('Skipped: {0}'.format(skipped))
        lines.append('Failed: {0}'.format(failed))
        lines.append('Duration: {0:.2f} s'.format(elapsed))
        lines.append('')
        failure_lines = result.get('failure_lines') or []
        skip_lines = result.get('skip_lines') or []
        if failure_lines:
            lines.append('Failed:')
            lines.extend(failure_lines[:200])
            if len(failure_lines) > 200:
                lines.append('... and {0} more failures (see log)'.format(
                    len(failure_lines) - 200))
        if skip_lines:
            lines.append('')
            lines.append('Skipped (sample):')
            lines.extend(skip_lines[:100])
            if skipped > len(skip_lines):
                lines.append('... and {0} more skipped'.format(skipped - len(skip_lines)))
        if not failure_lines and not skip_lines:
            lines.append('All elements updated successfully.')

        self.detailsTextBox.Text = '\n'.join(lines)

    def closeButton_Click(self, sender, args):
        self.Close()


def main():
    picker = NewIfcGuidWindow()
    picker.ShowDialog()
    if not picker._result:
        return
    result_window = NewIfcGuidResultWindow(picker._result)
    result_window.ShowDialog()


if __name__ == '__main__':
    main()
