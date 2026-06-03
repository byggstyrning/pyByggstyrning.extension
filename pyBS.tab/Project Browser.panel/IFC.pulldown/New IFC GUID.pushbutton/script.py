# -*- coding: utf-8 -*-
"""Assign new IFC GUIDs to selected elements or the whole model."""

__title__ = "New IFC\nGUID"
__highlight__ = "new"
__author__ = "Byggstyrning AB"
__doc__ = """Assign new IfcGUID values for IFC export.

Modes:
- Current selection (model instances)
- Whole model (instances and types)

Uses ExporterIFCUtils.CreateGUID. Writes via parameter Set when possible,
otherwise ExporterIFCUtils.AddValueString."""

import sys
import os.path as op
import time

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIIFC')

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    ElementType,
    FilteredElementCollector,
    RevitLinkInstance,
    StorageType,
    SubTransaction,
    Transaction,
    View,
    Wall,
)
from Autodesk.Revit.DB.IFC import ExporterIFCUtils

from pyrevit import forms, revit, script

MODE_SELECTION = 'selection'
MODE_WHOLE_MODEL = 'whole_model'
MAX_UNIQUE_ATTEMPTS = 25
PROGRESS_EVERY = 25
LARGE_RUN_ELEMENT_COUNT = 5000

_SKIP_CATEGORY_BUILTIN_NAMES = (
    'OST_Views',
    'OST_Sheets',
    'OST_Schedules',
    'OST_Legends',
    'OST_LegendComponents',
    'OST_DraftingViews',
    'OST_Viewers',
    'OST_Cameras',
    'OST_ReportSchemas',
)

logger = script.get_logger()

_SCRIPT_DIR = op.dirname(__file__)
# pushbutton -> pulldown -> panel -> tab -> extension
_EXTENSION_DIR = op.dirname(op.dirname(op.dirname(op.dirname(_SCRIPT_DIR))))
_LIB_PATH = op.join(_EXTENSION_DIR, 'lib')
if _LIB_PATH not in sys.path:
    sys.path.insert(0, _LIB_PATH)


def _built_in_category_ids(names):
    """Return int category ids for BuiltInCategory members that exist in this Revit build."""
    ids = set()
    for name in names:
        try:
            bic = getattr(BuiltInCategory, name, None)
            if bic is not None:
                ids.add(int(bic))
        except Exception:
            pass
    return frozenset(ids)


_SKIP_BUILT_IN_CATEGORY_IDS = _built_in_category_ids(_SKIP_CATEGORY_BUILTIN_NAMES)


def is_element_type(element):
    """True if element is a type (Element.IsElementType not in all API versions)."""
    try:
        return isinstance(element, ElementType)
    except Exception:
        return False


def is_valid_ifc_guid(guid):
    """Match Revit IFC exporter rules (22-char GlobalId)."""
    if not guid or len(guid) != 22:
        return False
    if guid[0] < '0' or guid[0] > '3':
        return False
    for ch in guid:
        if ('0' <= ch <= '9') or ('A' <= ch <= 'Z') or ('a' <= ch <= 'z'):
            continue
        if ch in ('_', '$'):
            continue
        return False
    return True


def _element_label(element):
    try:
        cat = element.Category.Name if element.Category else 'No category'
    except Exception:
        cat = 'No category'
    kind = 'Type' if is_element_type(element) else 'Instance'
    try:
        from revit.compat import get_element_id_value
        eid = get_element_id_value(element.Id)
    except Exception:
        eid = element.Id.IntegerValue
    return '{0} | Id {1} | {2}'.format(kind, eid, cat)


def _category_built_in(element):
    try:
        if element.Category is None:
            return None
        from revit.compat import get_element_id_value
        return get_element_id_value(element.Category.Id)
    except Exception:
        return None


def belongs_to_active_document(element, doc):
    try:
        el_doc = element.Document
        if el_doc is None:
            return False
        return el_doc.Equals(doc)
    except Exception:
        try:
            return element.Document == doc
        except Exception:
            return False


def is_element_alive(element):
    try:
        return element is not None and element.IsValidObject
    except Exception:
        return element is not None


def should_process_element(element, doc, allow_types):
    """Return (ok, reason) before calling IFC GUID APIs."""
    if not is_element_alive(element):
        return False, 'Invalid element'
    if not belongs_to_active_document(element, doc):
        return False, 'Not in active document (e.g. link)'
    if isinstance(element, RevitLinkInstance):
        return False, 'Revit link instance'
    if is_element_type(element) and not allow_types:
        return False, 'Element type (use Whole model for types)'
    if isinstance(element, View):
        return False, 'View'
    if element.Category is None:
        return False, 'No category'
    cat_id = _category_built_in(element)
    if cat_id is not None and cat_id in _SKIP_BUILT_IN_CATEGORY_IDS:
        return False, 'Unsupported category'
    try:
        from revit.revit_utils import is_element_editable
        editable, reason = is_element_editable(doc, element)
        if not editable:
            return False, reason
    except Exception:
        pass
    return True, ''


def can_store_ifc_guid(element):
    """Mirror exporter guard: curtain walls cannot store IfcGUID safely."""
    try:
        if isinstance(element, Wall):
            if element.CurtainGrid is not None:
                return False
    except Exception:
        pass
    return True


def get_ifc_guid_built_in_parameter(element):
    if is_element_type(element):
        return BuiltInParameter.IFC_TYPE_GUID
    return BuiltInParameter.IFC_GUID


def get_existing_ifc_guid_parameter(element):
    """Return writable IFC GUID parameter if already on element."""
    try:
        param = element.get_Parameter(get_ifc_guid_built_in_parameter(element))
        if param is None:
            return None
        if param.IsReadOnly:
            return None
        if param.StorageType != StorageType.String:
            return None
        return param
    except Exception:
        return None


def write_ifc_guid(element, new_guid):
    """
    Prefer Parameter.Set on existing IFC_GUID; fall back to AddValueString.
    Returns (ok, message).
    """
    if not is_valid_ifc_guid(new_guid):
        return False, 'Generated GUID failed validation'

    param = get_existing_ifc_guid_parameter(element)
    if param is not None:
        try:
            param.Set(str(new_guid))
            return True, ''
        except Exception as ex:
            logger.debug(
                'Parameter.Set failed on {0}, trying AddValueString: {1}'.format(
                    _element_label(element), ex))

    param_id = ElementId(get_ifc_guid_built_in_parameter(element))
    try:
        ExporterIFCUtils.AddValueString(element, param_id, new_guid)
        return True, ''
    except Exception as ex:
        return False, str(ex)


def generate_unique_ifc_guid(assigned_guids):
    for _ in range(MAX_UNIQUE_ATTEMPTS):
        try:
            candidate = ExporterIFCUtils.CreateGUID()
        except Exception as ex:
            logger.warning('CreateGUID failed: {0}'.format(ex))
            candidate = None
        if candidate and is_valid_ifc_guid(candidate) and candidate not in assigned_guids:
            return candidate
    return None


def assign_ifc_guid(element, assigned_guids):
    """Assign one element. Returns (status, message). Never raises."""
    try:
        if not is_element_alive(element):
            return 'failed', 'Element no longer valid'
        if not can_store_ifc_guid(element):
            return 'skipped', 'Curtain wall (IfcGUID not stored by exporter)'

        new_guid = generate_unique_ifc_guid(assigned_guids)
        if not new_guid:
            return 'failed', 'Could not generate a unique IFC GUID'

        ok, message = write_ifc_guid(element, new_guid)
        if not ok:
            return 'failed', message

        assigned_guids.add(new_guid)
        return 'updated', new_guid
    except Exception as ex:
        return 'failed', str(ex)


def collect_elements(doc, mode):
    allow_types = mode == MODE_WHOLE_MODEL
    elements = []
    rejected = []

    if mode == MODE_SELECTION:
        selection = revit.get_selection()
        if not selection or not selection.element_ids:
            return elements, rejected
        for element_id in selection.element_ids:
            try:
                element = doc.GetElement(element_id)
            except Exception:
                rejected.append((None, 'Could not resolve selection id'))
                continue
            ok, reason = should_process_element(element, doc, allow_types=False)
            if ok:
                elements.append(element)
            elif element is not None:
                rejected.append((element, reason))
        return elements, rejected

    for element in FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
        ok, reason = should_process_element(element, doc, allow_types=False)
        if ok:
            elements.append(element)
    for element in FilteredElementCollector(doc).WhereElementIsElementType().ToElements():
        ok, reason = should_process_element(element, doc, allow_types=True)
        if ok:
            elements.append(element)
    return elements, rejected


def _safe_rollback_subtransaction(sub_txn):
    try:
        if sub_txn is not None and sub_txn.HasStarted() and not sub_txn.HasEnded():
            sub_txn.RollBack()
    except Exception as ex:
        logger.debug('SubTransaction rollback: {0}'.format(ex))


def _process_one_element(doc, element, assigned_guids):
    """One element inside its own SubTransaction. Returns (status, message)."""
    if not is_element_alive(element):
        return 'failed', 'Element no longer valid'

    sub_txn = SubTransaction(doc)
    try:
        sub_txn.Start()
    except Exception as ex:
        return 'failed', 'Could not start sub-transaction: {0}'.format(ex)

    try:
        status, message = assign_ifc_guid(element, assigned_guids)
        if status == 'updated':
            try:
                sub_txn.Commit()
            except Exception as ex:
                _safe_rollback_subtransaction(sub_txn)
                return 'failed', 'Commit failed: {0}'.format(ex)
        else:
            _safe_rollback_subtransaction(sub_txn)
        return status, message
    except Exception as ex:
        _safe_rollback_subtransaction(sub_txn)
        return 'failed', str(ex)


def run_ifc_guid_assignment(doc, elements, progress_callback=None):
    """Run assignment in one transaction; each element uses a SubTransaction."""
    started = time.time()
    assigned_guids = set()
    counts = {'updated': 0, 'skipped': 0, 'failed': 0}
    failure_lines = []
    skip_lines = []
    total = len(elements)

    txn = Transaction(doc, 'New IFC GUID')
    txn.Start()
    try:
        for index, element in enumerate(elements):
            status, message = _process_one_element(doc, element, assigned_guids)
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
    except Exception as ex:
        try:
            if txn.HasStarted() and not txn.HasEnded():
                txn.RollBack()
        except Exception:
            pass
        raise Exception('Transaction rolled back: {0}'.format(ex))

    elapsed = time.time() - started
    return {
        'total': total,
        'updated': counts.get('updated', 0),
        'skipped': counts.get('skipped', 0),
        'failed': counts.get('failed', 0),
        'failure_lines': failure_lines,
        'skip_lines': skip_lines,
        'elapsed_seconds': elapsed,
        'rolled_back': False,
    }


def confirm_run(mode_label, element_count, rejected_count, is_whole_model):
    """Ask user to confirm before modifying the model."""
    lines = [
        'Assign new IfcGUID values to {0} element(s)?'.format(element_count),
        '',
        'Scope: {0}.'.format(mode_label),
        'You can undo this with Revit Undo (one transaction).',
    ]
    if rejected_count:
        lines.append(
            '{0} item(s) in selection were excluded (views, links, types, etc.).'.format(
                rejected_count))
    if is_whole_model:
        lines.append('')
        lines.append('Whole model may take a while on large projects.')
        if element_count >= LARGE_RUN_ELEMENT_COUNT:
            lines.append(
                'Warning: more than {0} elements — consider testing on selection first.'.format(
                    LARGE_RUN_ELEMENT_COUNT))
    return forms.alert('\n'.join(lines), yes=True, no=True, title='New IFC GUID')


class NewIfcGuidWindow(forms.WPFWindow):
    """Mode picker — closes before Revit transaction."""

    def __init__(self):
        xaml_path = op.join(_SCRIPT_DIR, 'NewIfcGuidWindow.xaml')
        forms.WPFWindow.__init__(self, xaml_path)
        from styles import load_styles_to_window
        load_styles_to_window(self)
        self.run_requested = False
        self.pending_elements = None
        self.pending_mode_label = ''
        self.pending_rejected_count = 0
        self.pending_is_whole_model = False

    def _selected_mode(self):
        if self.wholeModelModeRadio.IsChecked:
            return MODE_WHOLE_MODEL
        return MODE_SELECTION

    def generateButton_Click(self, sender, args):
        doc = revit.doc
        if getattr(doc, 'IsFamilyDocument', False):
            forms.alert(
                'This tool runs on project documents only, not family files.',
                title='New IFC GUID',
            )
            return

        mode = self._selected_mode()
        is_whole_model = mode == MODE_WHOLE_MODEL
        if is_whole_model:
            mode_label = 'Whole model'
        else:
            mode_label = 'Current selection'

        elements, rejected = collect_elements(doc, mode)
        if not elements:
            if mode == MODE_SELECTION:
                msg = (
                    'No eligible model instances in the current selection.\n\n'
                    'Select walls, doors, furniture, etc. (not views, links, or sheets). ')
                if rejected:
                    msg += '\n{0} selected item(s) were excluded.'.format(len(rejected))
                msg += '\n\nOr use Whole model mode.'
                forms.alert(msg, title='New IFC GUID')
            else:
                forms.alert('No eligible elements found in the model.', title='New IFC GUID')
            return

        if not confirm_run(mode_label, len(elements), len(rejected), is_whole_model):
            return

        self.pending_elements = elements
        self.pending_mode_label = mode_label
        self.pending_rejected_count = len(rejected)
        self.pending_is_whole_model = is_whole_model
        self.run_requested = True
        self.Close()

    def cancelButton_Click(self, sender, args):
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
        rejected_count = result.get('rejected_count', 0)

        self.updatedCountText.Text = str(updated)
        self.skippedCountText.Text = str(skipped)
        self.failedCountText.Text = str(failed)

        summary = '{0} — processed {1} element(s) in {2:.1f} s.'.format(
            mode_label, total, elapsed)
        if rejected_count:
            summary += ' {0} item(s) excluded before run.'.format(rejected_count)
        if result.get('rolled_back'):
            summary += ' Transaction was rolled back.'
        summary += ' Use Revit Undo if you need to revert changes.'
        self.summaryLineText.Text = summary

        lines = []
        lines.append('Mode: {0}'.format(mode_label))
        lines.append('Total processed: {0}'.format(total))
        lines.append('Updated: {0}'.format(updated))
        lines.append('Skipped: {0}'.format(skipped))
        lines.append('Failed: {0}'.format(failed))
        if rejected_count:
            lines.append('Excluded before run: {0}'.format(rejected_count))
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
            lines.append('All processed elements updated successfully.')

        self.detailsTextBox.Text = '\n'.join(lines)

    def closeButton_Click(self, sender, args):
        self.Close()


def main():
    doc = revit.doc
    if doc is None:
        forms.alert('No active document.', title='New IFC GUID')
        return

    picker = NewIfcGuidWindow()
    picker.ShowDialog()
    if not picker.run_requested or not picker.pending_elements:
        return

    elements = picker.pending_elements
    mode_label = picker.pending_mode_label
    rejected_count = picker.pending_rejected_count

    result = None
    try:
        with forms.ProgressBar(title='Assigning new IFC GUIDs...') as pb:
            def on_progress(current, total):
                pb.update_progress(current, total)

            result = run_ifc_guid_assignment(doc, elements, on_progress)
    except Exception as ex:
        logger.error('IFC GUID assignment failed: {0}'.format(ex))
        forms.alert(
            'IFC GUID assignment failed. No changes were committed.\n\n{0}'.format(ex),
            title='New IFC GUID',
        )
        return

    result['mode_label'] = mode_label
    result['rejected_count'] = rejected_count
    NewIfcGuidResultWindow(result).ShowDialog()


_window_ref = None

if __name__ == '__main__':
    _window_ref = main()
