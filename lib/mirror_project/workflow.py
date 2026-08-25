# -*- coding: utf-8 -*-
"""User-facing Mirror Project workflow actions."""

from __future__ import print_function

import os

from pyrevit import forms, script

from .commands import schedule_native_command
from .constraints import classify_document_constraints, disable_eligible_constraints
from .diff import compare_snapshots
from .reporting import export_findings_csv, print_diff, print_preflight
from .review import build_review_package
from .review_views import create_review_views
from .schema import (
    MANIFEST_SCHEMA_VERSION,
    POINTER_SCHEMA_VERSION,
    REVIEW_SCHEMA_VERSION,
    SCHEMA_VERSION,
)
from .snapshot import capture_snapshot, preflight_summary, utc_now
from .storage import load_json, save_json

logger = script.get_logger()


MENU_PREFLIGHT = 'Run Preflight'
MENU_CAPTURE = 'Capture Baseline'
MENU_PREPARE = 'Prepare + Launch Mirror Project'
MENU_COMPARE = 'Compare With Baseline'
MENU_REVIEW_VIEWS = 'Create Review Views From JSON'

MENU_ITEMS = (
    MENU_PREFLIGHT,
    MENU_CAPTURE,
    MENU_PREPARE,
    MENU_COMPARE,
    MENU_REVIEW_VIEWS,
)


def run_menu(doc, uiapp):
    if doc is None:
        forms.alert('Open a project document first.', title='Mirror Project')
        return
    if doc.IsFamilyDocument:
        forms.alert('Use Mirror Project in a project document, not a family.', title='Mirror Project')
        return
    selected = forms.CommandSwitchWindow.show(
        list(MENU_ITEMS),
        message='Mirror Project workflow',
    )
    if not selected:
        return
    if selected == MENU_PREFLIGHT:
        run_preflight(doc)
    elif selected == MENU_CAPTURE:
        run_capture(doc)
    elif selected == MENU_PREPARE:
        run_prepare_and_mirror(doc, uiapp)
    elif selected == MENU_COMPARE:
        run_compare(doc)
    elif selected == MENU_REVIEW_VIEWS:
        run_create_review_views(doc)


def run_preflight(doc):
    snapshot = capture_snapshot(doc)
    summary = preflight_summary(snapshot)
    output = script.get_output()
    print_preflight(output, summary)
    return snapshot


def run_capture(doc):
    path = forms.save_file(file_ext='json', title='Save Mirror Project baseline')
    if not path:
        return None
    snapshot = capture_snapshot(doc)
    save_json(path, snapshot)
    _store_pointer(path, snapshot)
    summary = preflight_summary(snapshot)
    output = script.get_output()
    print_preflight(output, summary, snapshot_path=path)
    forms.alert('Baseline saved:\n{}'.format(path), title='Mirror Project')
    return path


def run_prepare_and_mirror(doc, uiapp):
    baseline_path = _pick_baseline(required=True)
    if not baseline_path:
        return
    try:
        baseline = load_json(baseline_path)
    except Exception as ex:
        forms.alert('Could not read baseline:\n{}'.format(ex), title='Mirror Project')
        return
    if baseline.get('schema_version') != SCHEMA_VERSION:
        forms.alert(
            'Baseline schema {} does not match {}.'.format(
                baseline.get('schema_version'), SCHEMA_VERSION),
            title='Mirror Project',
        )
        return

    buckets = classify_document_constraints(doc)
    eligible = buckets.get('eligible') or []
    baseline_constraints = (
        (((baseline.get('metrics') or {}).get('constraints') or {}).get('data') or {})
        .get('items') or [])
    baseline_eligible_locked = set(
        item.get('unique_id') for item in baseline_constraints
        if item.get('classification') == 'eligible'
        and item.get('is_locked') is True
        and item.get('unique_id'))
    current_unlocked = {}
    for item in buckets.get('unlocked') or []:
        uid = item.get('unique_id')
        if uid:
            current_unlocked[uid] = item
    preexisting_unlocked = [
        current_unlocked[uid] for uid in sorted(baseline_eligible_locked)
        if uid in current_unlocked
    ]
    residual = []
    for key in ('grouped', 'multi_segment', 'labeled', 'unsupported_type', 'missing', 'unmodifiable'):
        residual.extend(buckets.get(key) or [])
    locked_residual = [item for item in residual if item.get('is_locked')]

    identity = (baseline.get('fingerprint') or {})
    confirm = (
        'This will unlock {eligible} eligible locked constraints and then launch '
        'Revit Mirror Project.\n\n'
        'Locks stay disabled even if you cancel the native Mirror Project dialog. '
        'They are NOT restored.\n\n'
        'Baseline:\n{path}\nCaptured: {captured}\n'
        'Document: {title}\nWorkshared: {ws} | Detached: {det} | Cloud: {cloud}\n\n'
        'Eligible to unlock: {eligible}\nResidual locked (will remain): {residual}\n\n'
        'Already unlocked since baseline: {preexisting}\n'
        '(recorded for audit; this run will not modify them)\n\n'
        'Continue?'
    ).format(
        eligible=len(eligible),
        path=baseline_path,
        captured=baseline.get('captured_utc'),
        title=identity.get('title') or doc.Title,
        ws=doc.IsWorkshared,
        det=_doc_flag(doc, 'IsDetached'),
        cloud=_doc_flag(doc, 'IsModelInCloud'),
        residual=len(locked_residual),
        preexisting=len(preexisting_unlocked),
    )
    if not forms.alert(confirm, yes=True, no=True, warn_icon=True, title='Prepare + Launch Mirror Project'):
        return

    manifest_path = _sibling_path(baseline_path, '_manifest.json')
    manifest = {
        'schema_version': MANIFEST_SCHEMA_VERSION,
        'phase': 'prepared',
        'created_utc': utc_now(),
        'baseline_path': baseline_path,
        'baseline_captured_utc': baseline.get('captured_utc'),
        'baseline_fingerprint': identity,
        'action': 'disable_constraints_no_restore',
        'classified': {
            'eligible': eligible,
            'grouped': buckets.get('grouped') or [],
            'multi_segment': buckets.get('multi_segment') or [],
            'labeled': buckets.get('labeled') or [],
            'unsupported_type': buckets.get('unsupported_type') or [],
            'missing': buckets.get('missing') or [],
            'unmodifiable': buckets.get('unmodifiable') or [],
            'unlocked': buckets.get('unlocked') or [],
            'preexisting_unlocked_since_baseline': preexisting_unlocked,
        },
        'unlock_results': {
            'committed': [],
            'failed': [],
            'skipped': [],
        },
        'transaction_status': 'prepared',
    }
    save_json(manifest_path, manifest)

    results = disable_eligible_constraints(doc, eligible)
    manifest['phase'] = 'disabled'
    manifest['unlock_results'] = results
    manifest['transaction_status'] = results.get('transaction_status')
    manifest['disabled_utc'] = utc_now()
    save_json(manifest_path, manifest)
    _store_manifest_pointer(manifest_path)

    output = script.get_output()
    output.print_md('# Constraint preparation')
    output.print_md('Manifest: `{}`'.format(manifest_path))
    output.print_md('- Unlocked: {}'.format(len(results.get('committed') or [])))
    output.print_md('- Failed: {}'.format(len(results.get('failed') or [])))
    output.print_md('- Skipped: {}'.format(len(results.get('skipped') or [])))
    output.print_md('- Residual locked: {}'.format(len(locked_residual)))
    output.print_md('Transaction: {}'.format(results.get('transaction_status')))
    if results.get('failed'):
        forms.alert(
            'Some constraints failed to unlock. Mirror Project will still be posted. '
            'See the output window.',
            title='Mirror Project',
        )

    ok, message = schedule_native_command(uiapp, 'mirror_project', logger=logger)
    if not ok:
        forms.alert(message, title='Mirror Project')
        return
    output.print_md('---')
    output.print_md('**Mirror Project** starts when this script finishes.')
    output.print_md('1. Pick the mirror **axis** on the model or a level/grid line.')
    output.print_md('2. Choose the mirror **direction** when Revit asks.')
    output.print_md('If you cancel, unlocked constraints stay unlocked.')
    output.print_md('Manifest: `{}`'.format(manifest_path))


def run_compare(doc):
    baseline_path = _pick_baseline(required=True)
    if not baseline_path:
        return
    try:
        baseline = load_json(baseline_path)
    except Exception as ex:
        forms.alert('Could not read baseline:\n{}'.format(ex), title='Mirror Project')
        return
    manifest = None
    default_manifest = _sibling_path(baseline_path, '_manifest.json')
    pointer_manifest = _load_manifest_pointer()
    manifest_path = None
    if pointer_manifest and os.path.isfile(pointer_manifest):
        manifest_path = pointer_manifest
    elif os.path.isfile(default_manifest):
        manifest_path = default_manifest
    else:
        picked = forms.pick_file(file_ext='json', title='Optional mutation manifest (cancel to skip)')
        if picked:
            manifest_path = picked
    if manifest_path:
        try:
            manifest = load_json(manifest_path)
        except Exception as ex:
            logger.warning('Manifest load failed: {}'.format(ex))
            manifest = None

    current = capture_snapshot(doc)
    comparison = compare_snapshots(baseline, current, manifest=manifest)
    output = script.get_output()

    def linkify(element_token):
        return _try_linkify(output, doc, element_token)

    print_diff(output, comparison, linkify_fn=linkify)
    review_path = forms.save_file(
        file_ext='json',
        title='Export review package JSON (use in non-mirrored model)',
    )
    if review_path:
        review_package = build_review_package(
            baseline,
            current,
            comparison,
            exported_utc=utc_now(),
        )
        save_json(review_path, review_package)
        output.print_md('Review package exported: `{}`'.format(review_path))
    csv_path = forms.save_file(file_ext='csv', title='Export compare CSV (cancel to skip)')
    if csv_path:
        export_findings_csv(csv_path, comparison.get('findings') or [])
        output.print_md('CSV exported: `{}`'.format(csv_path))
    summary = comparison.get('summary') or {}
    forms.alert(
        'Compare complete.\nBlocking: {}\nReview: {}\nInfo: {}'.format(
            summary.get('blocking'), summary.get('review'), summary.get('informational')),
        title='Mirror Project',
    )


def run_create_review_views(doc):
    package_path = forms.pick_file(
        file_ext='json',
        title='Select Mirror Project review package JSON',
    )
    if not package_path:
        return
    try:
        package = load_json(package_path)
    except Exception as ex:
        forms.alert(
            'Could not read review package:\n{}'.format(ex),
            title='Mirror Project Review Views',
        )
        return
    if package.get('schema_version') != REVIEW_SCHEMA_VERSION:
        forms.alert(
            'Selected JSON is not a review package.\nFound: {}\nExpected: {}'.format(
                package.get('schema_version'), REVIEW_SCHEMA_VERSION),
            title='Mirror Project Review Views',
        )
        return

    baseline_title = (package.get('baseline_fingerprint') or {}).get('title') or ''
    if baseline_title and baseline_title != doc.Title:
        proceed = forms.alert(
            'Review package baseline is "{}" but active document is "{}".\n\n'
            'Continue only if this is the corresponding non-mirrored model.'.format(
                baseline_title, doc.Title),
            yes=True,
            no=True,
            warn_icon=True,
            title='Mirror Project Review Views',
        )
        if not proceed:
            return

    try:
        result = create_review_views(doc, package)
    except Exception as ex:
        forms.alert(
            'Could not create review views:\n{}'.format(ex),
            title='Mirror Project Review Views',
        )
        return

    output = script.get_output()
    output.print_md('# Mirror Project Review Views')
    output.print_md('Package: `{}`'.format(package_path))
    output.print_md('- Requested element targets: {}'.format(result.get('requested')))
    output.print_md('- Resolved in source model: {}'.format(result.get('resolved')))
    output.print_md('- Unresolved UniqueIds: {}'.format(len(result.get('unresolved') or [])))
    output.print_md('- Non-graphical targets: {}'.format(
        len(result.get('non_graphical') or [])))
    output.print_md('')
    output.print_md('## Created views')
    for item in result.get('created_views') or []:
        output.print_md(
            '- **{}** — {} targets / {} isolated elements. Suggested action: {}'.format(
                item.get('name'),
                item.get('target_count'),
                item.get('element_count'),
                item.get('suggested_action'),
            ))
    _print_id_list(output, 'Unresolved UniqueIds', result.get('unresolved') or [])
    _print_id_list(output, 'Non-graphical UniqueIds', result.get('non_graphical') or [])
    if result.get('isolation_failures') or result.get('section_failures'):
        output.print_md('')
        output.print_md('## View creation warnings')
        for item in result.get('isolation_failures') or []:
            output.print_md('- Isolation failed in **{}**: {}'.format(
                item.get('view'), item.get('error')))
        for item in result.get('section_failures') or []:
            output.print_md('- Section box failed in **{}**: {}'.format(
                item.get('view'), item.get('error')))

    forms.alert(
        'Created {} review views.\nResolved: {}\nUnresolved: {}\nNon-graphical: {}'.format(
            len(result.get('created_views') or []),
            result.get('resolved'),
            len(result.get('unresolved') or []),
            len(result.get('non_graphical') or []),
        ),
        title='Mirror Project Review Views',
    )


def _print_id_list(output, title, values):
    if not values:
        return
    output.print_md('')
    output.print_md('## {}'.format(title))
    for value in values[:100]:
        output.print_md('- `{}`'.format(value))
    if len(values) > 100:
        output.print_md('- ... {} more'.format(len(values) - 100))


def _doc_flag(doc, name):
    try:
        return getattr(doc, name)
    except Exception:
        return 'unknown'


def _pointer_path():
    return script.get_document_data_file('mirror_project_last_baseline', 'json')


def _manifest_pointer_path():
    return script.get_document_data_file('mirror_project_last_manifest', 'json')


def _store_pointer(path, snapshot):
    payload = {
        'schema_version': POINTER_SCHEMA_VERSION,
        'path': path,
        'captured_utc': snapshot.get('captured_utc'),
        'fingerprint': snapshot.get('fingerprint'),
    }
    try:
        save_json(_pointer_path(), payload)
    except Exception as ex:
        logger.warning('Could not store baseline pointer: {}'.format(ex))


def _store_manifest_pointer(path):
    payload = {
        'schema_version': POINTER_SCHEMA_VERSION,
        'path': path,
        'created_utc': utc_now(),
    }
    try:
        save_json(_manifest_pointer_path(), payload)
    except Exception as ex:
        logger.warning('Could not store manifest pointer: {}'.format(ex))


def _load_manifest_pointer():
    path = _manifest_pointer_path()
    if not os.path.isfile(path):
        return None
    try:
        data = load_json(path)
        return data.get('path')
    except Exception:
        return None


def _pick_baseline(required=False):
    pointer = _pointer_path()
    last_path = None
    if os.path.isfile(pointer):
        try:
            data = load_json(pointer)
            last_path = data.get('path')
        except Exception:
            last_path = None
    if last_path and os.path.isfile(last_path):
        use_last = forms.alert(
            'Use last baseline?\n{}'.format(last_path),
            yes=True,
            no=True,
            title='Mirror Project',
        )
        if use_last:
            return last_path
    picked = forms.pick_file(file_ext='json', title='Select Mirror Project baseline')
    if not picked and required:
        forms.alert('A baseline JSON is required.', title='Mirror Project')
    return picked


def _sibling_path(baseline_path, suffix):
    root, ext = os.path.splitext(baseline_path)
    return root + suffix


def _try_linkify(output, doc, token):
    if token is None:
        return ''
    text = str(token)
    element = None
    try:
        element = doc.GetElement(text)
    except Exception:
        element = None
    if element is None:
        try:
            from revit.compat import make_element_id
            element = doc.GetElement(make_element_id(int(text)))
        except Exception:
            element = None
    if element is None:
        return text
    try:
        return output.linkify(element.Id)
    except Exception:
        return text
