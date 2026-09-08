# -*- coding: utf-8 -*-
"""Markdown and CSV reporting for Mirror Project snapshots and diffs."""

from __future__ import print_function

import csv
import os

from .schema import SEVERITY_BLOCKING, SEVERITY_INFO, SEVERITY_REVIEW
from .storage import json_safe, validate_user_data_path


CSV_COLUMNS = (
    'severity',
    'check',
    'category',
    'occurrences',
    'examples',
    'element',
    'before',
    'after',
    'action',
)


def write_findings_csv(path, findings):
    safe_path = validate_user_data_path(path, ('.csv',))
    directory = os.path.dirname(safe_path)
    if directory and not os.path.isdir(directory):
        os.makedirs(directory)
    with open(safe_path, 'wb') as handle:
        writer = csv.DictWriter(handle, fieldnames=list(CSV_COLUMNS))
        writer.writeheader()
        for item in findings or []:
            row = {}
            for col in CSV_COLUMNS:
                value = item.get(col)
                if value is None:
                    row[col] = ''
                elif isinstance(value, (str,)):
                    row[col] = value.encode('utf-8') if not isinstance(value, bytes) else value
                else:
                    text = '{}'.format(json_safe(value))
                    try:
                        row[col] = text.encode('utf-8')
                    except Exception:
                        row[col] = text
            writer.writerow(row)
    return safe_path


def _write_csv_py3(path, findings):
    safe_path = validate_user_data_path(path, ('.csv',))
    directory = os.path.dirname(safe_path)
    if directory and not os.path.isdir(directory):
        os.makedirs(directory)
    with open(safe_path, 'w', newline='', encoding='utf-8') as handle:
        writer = csv.DictWriter(handle, fieldnames=list(CSV_COLUMNS))
        writer.writeheader()
        for item in findings or []:
            row = {}
            for col in CSV_COLUMNS:
                value = item.get(col)
                row[col] = '' if value is None else '{}'.format(json_safe(value))
            writer.writerow(row)
    return safe_path


def export_findings_csv(path, findings):
    """Write findings CSV using a Python 2/3 compatible writer."""
    try:
        unicode  # noqa: F821  IronPython / Python 2
        return write_findings_csv(path, findings)
    except NameError:
        return _write_csv_py3(path, findings)


def print_preflight(output, summary, snapshot_path=None):
    output.print_md('# Mirror Project Preflight')
    output.print_md('Document: **{}**'.format(summary.get('title') or ''))
    output.print_md('Revit build: {}'.format(summary.get('revit_build') or ''))
    output.print_md('Workshared: {} | Detached: {} | Cloud: {}'.format(
        summary.get('workshared'), summary.get('detached'), summary.get('cloud')))
    if snapshot_path:
        output.print_md('Baseline path: `{}`'.format(snapshot_path))
    output.print_md('')
    output.print_md('## Counts')
    output.print_md('- Instances: {}'.format(summary.get('instances')))
    output.print_md('- Doors / windows / walls: {} / {} / {}'.format(
        summary.get('doors'), summary.get('windows'), summary.get('walls')))
    output.print_md('- Views / sheets: {} / {}'.format(
        summary.get('views'), summary.get('sheets')))
    output.print_md('- Rooms / spaces: {} / {}'.format(
        summary.get('rooms'), summary.get('spaces')))
    output.print_md('- Room states: {}'.format(summary.get('room_states')))
    output.print_md('- Space states: {}'.format(summary.get('space_states')))
    output.print_md('')
    output.print_md('## Risk')
    output.print_md('- Warnings: {} ({} groups)'.format(
        summary.get('warning_total'), summary.get('warning_groups')))
    output.print_md('- Attached walls: {}'.format(summary.get('attached_walls')))
    output.print_md('- Locked constraints: {}'.format(summary.get('locked_constraints')))
    output.print_md('- Constraint classes: {}'.format(summary.get('constraint_classes')))
    output.print_md('- Groups / nested: {} / {}'.format(
        summary.get('groups'), summary.get('nested_groups')))
    output.print_md('- Unused MEP connectors: {}'.format(summary.get('mep_unused')))
    output.print_md('- Links: {}'.format(summary.get('links')))
    output.print_md('- True North (deg): {}'.format(summary.get('true_north_deg')))
    failed = summary.get('failed_collectors') or []
    if failed:
        output.print_md('')
        output.print_md('## Failed collectors')
        for name in failed:
            output.print_md('- {}'.format(name))
    output.print_md('')
    output.print_md('This report does not claim design-intent correctness.')


def print_diff(output, comparison, linkify_fn=None):
    summary = comparison.get('summary') or {}
    transform = comparison.get('transform') or {}
    findings = comparison.get('findings') or []
    output.print_md('# Mirror Project Compare')
    output.print_md('Blocking: **{}** | Review: **{}** | Info: {}'.format(
        summary.get('blocking'), summary.get('review'), summary.get('informational')))
    output.print_md(
        'Underlying occurrences: blocking {} | review {} | info {}'.format(
            summary.get('blocking_occurrences'),
            summary.get('review_occurrences'),
            summary.get('informational_occurrences'),
        ))
    output.print_md(
        'Inferred transform: `{}` (confidence {}, inlier rms {} mm, '
        'inliers {}/{}, ratio {})'.format(
        transform.get('name'),
        transform.get('confidence'),
        transform.get('rms_mm'),
        transform.get('n_inliers'),
        transform.get('n_samples'),
        transform.get('inlier_ratio'),
    ))
    if not summary.get('passed'):
        output.print_md('**Acceptance: failed** — unexplained blocking findings remain.')
    else:
        output.print_md('**Acceptance: no blocking findings.** Review items still need a model-manager check.')
    output.print_md('')
    _print_finding_section(output, findings, SEVERITY_BLOCKING, 'Blocking', linkify_fn)
    _print_finding_section(output, findings, SEVERITY_REVIEW, 'Review', linkify_fn)
    _print_finding_section(output, findings, SEVERITY_INFO, 'Informational', linkify_fn)
    output.print_md('')
    output.print_md('Handedness, egress, MEP left/right, and slope still need human review.')


def _print_finding_section(output, findings, severity, title, linkify_fn):
    rows = [item for item in findings if item.get('severity') == severity]
    output.print_md('## {} ({})'.format(title, len(rows)))
    if not rows:
        output.print_md('None.')
        return
    limit = 80
    for item in rows[:limit]:
        element = item.get('element') or ''
        linked = element
        if element and linkify_fn:
            try:
                linked = linkify_fn(element) or element
            except Exception:
                linked = element
        occurrence_text = ''
        if int(item.get('occurrences') or 1) > 1:
            occurrence_text = ' [{} occurrences]'.format(item.get('occurrences'))
        output.print_md('- **{}** / {} / {}{}: {} → {} — {}'.format(
            item.get('check'),
            item.get('category'),
            linked,
            occurrence_text,
            item.get('before'),
            item.get('after'),
            item.get('action'),
        ))
    if len(rows) > limit:
        output.print_md('- ... {} more (see CSV export).'.format(len(rows) - limit))
