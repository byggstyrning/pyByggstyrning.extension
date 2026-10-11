# -*- coding: utf-8 -*-
"""Optional sharing of nested families that are copied into several hosts.

Sharing stores one definition instead of a copy inside every host. Each nested
placement then becomes its own element. On the M0042 electrical model that cut
the file further and added thousands of fixture elements, and Revit deleted a
type in hosts that carried a different version of the nested family.

Each nested load sits in a transaction group. The group is rolled back when the
load fails, when strict quiet mode rejects a deletion, or when a type or placed
instance disappears. The host is loaded back into the project only if at least
one nested family was shared, and that reload is rolled back the same way.
"""

import os
import time
from collections import Counter

from Autodesk.Revit import DB

from filesize.family_purge import (
    _append_json_line,
    _emit,
    _find_family,
    _instance_count,
    _load_ok,
    _text,
    quiet_transaction,
)
from filesize.share_checks import category_delta, reload_loss


def _safe_name(name):
    bad = u'<>:"/\\|?*'
    cleaned = []
    for char in _text(name):
        cleaned.append(u'_' if char in bad else char)
    text = u''.join(cleaned).strip(u' .')
    return text or u'family'


def _close(document):
    if document is None:
        return
    try:
        document.Close(False)
    except Exception:
        pass


def _rollback_if_open(transaction):
    try:
        if transaction is not None and transaction.HasStarted() and not transaction.HasEnded():
            transaction.RollBack()
    except Exception:
        pass


def _group_open(group):
    try:
        return group is not None and group.HasStarted() and not group.HasEnded()
    except Exception:
        return False


def _discard_group(group):
    if not _group_open(group):
        return
    try:
        group.RollBack()
    except Exception:
        pass


def _keep_group(group):
    if not _group_open(group):
        raise Exception(u'transaction group was already closed')
    status = group.Assimilate()
    if status != DB.TransactionStatus.Committed:
        _discard_group(group)
        raise Exception(u'transaction group was not kept')


def _symbol_names(document, family_name):
    names = set()
    for family in DB.FilteredElementCollector(document).OfClass(DB.Family):
        if family.Name != family_name:
            continue
        for symbol_id in family.GetFamilySymbolIds():
            symbol = document.GetElement(symbol_id)
            if symbol is not None and symbol.Name:
                names.add(symbol.Name)
    return names


def _placed_count(document, family_name):
    count = 0
    for instance in DB.FilteredElementCollector(document).OfClass(DB.FamilyInstance):
        try:
            family = instance.Symbol.Family
        except Exception:
            family = None
        if family is not None and family.Name == family_name:
            count += 1
    return count


def _direct_nested(family_doc, name):
    for nested in DB.FilteredElementCollector(family_doc).OfClass(DB.Family):
        if nested.Name == name:
            return nested
    return None


def _is_model_family(family):
    category = family.FamilyCategory
    if category is None:
        return True
    return category.CategoryType == DB.CategoryType.Model


class _IncomingSharedLoadOptions(DB.IFamilyLoadOptions):
    """Load the shared family from the file being loaded.

    FamilySource.Project would keep the copy already in the host, so a host
    with a different version would never switch and the type-loss rollback
    would not run. Parameter values already in the host are kept.
    """

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source,
                            overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True


def _invoke_load(document, loader):
    """Call loader. Retry inside a transaction only if Revit requires one."""
    try:
        return loader()
    except Exception as ex:
        if u'transaction' not in _text(ex).lower():
            raise
    transaction = quiet_transaction(DB.Transaction(document, 'Load family'))
    transaction.Start()
    try:
        loaded = loader()
        transaction.Commit()
        if transaction.GetStatus() != DB.TransactionStatus.Committed:
            raise Exception(u'load was rolled back')
        return loaded
    except Exception:
        _rollback_if_open(transaction)
        raise


def _load_in_group(document, loader, accept=None):
    """Load inside a transaction group. accept() runs before the group is kept.

    LoadFamily opens its own transaction, so the first attempt does not wrap
    it in another one. A refusal that mentions a transaction is retried inside
    one that still belongs to this group. accept() raising rolls the load back.
    """
    group = DB.TransactionGroup(document, 'Family share load')
    group.Start()
    try:
        loaded = _invoke_load(document, loader)
        if not _load_ok(loaded):
            raise Exception(u'load was rolled back')
        if accept is not None:
            accept()
        _keep_group(group)
    except Exception:
        _discard_group(group)
        raise


def _is_shared(family):
    shared = family.get_Parameter(DB.BuiltInParameter.FAMILY_SHARED)
    return bool(shared and shared.AsInteger())


def _save_family(family_doc, path):
    options = DB.SaveAsOptions()
    options.OverwriteExistingFile = True
    options.MaximumBackups = 1
    family_doc.SaveAs(path, options)
    root, ext = os.path.splitext(path)
    for index in range(1, 10):
        backup = u'{0}.{1:04d}{2}'.format(root, index, ext)
        if os.path.exists(backup):
            try:
                os.remove(backup)
            except Exception:
                pass


def _category_counts(doc):
    counts = Counter()
    for instance in DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance):
        category = instance.Category
        name = category.Name if category is not None and category.Name else u'?'
        counts[name] += 1
    return dict(counts)


def duplicated_nested(doc, on_status=None):
    """Non-shared model families loaded in two or more hosts.

    Returns a list of (nested name, sorted host names). One host that fails to
    open is skipped.
    """
    hosts = {}
    families = []
    seen = set()
    for family in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if not family.Name or not family.IsEditable or family.IsInPlace:
            continue
        if family.Name in seen:
            continue
        seen.add(family.Name)
        families.append(family.Name)
    families.sort()
    total = len(families)
    _emit(on_status, u'Looking for nested families copied into several hosts ({0})'.format(total))
    for index, host_name in enumerate(families):
        if index % 25 == 0:
            _emit(on_status, u'  scan {0}/{1}'.format(index + 1, total))
        host = _find_family(doc, host_name)
        if host is None:
            continue
        family_doc = None
        try:
            family_doc = doc.EditFamily(host)
            for nested in DB.FilteredElementCollector(family_doc).OfClass(DB.Family):
                if not nested.Name or nested.Name == host_name:
                    continue
                if not _is_model_family(nested) or _is_shared(nested):
                    continue
                found = hosts.get(nested.Name)
                if found is None:
                    found = []
                    hosts[nested.Name] = found
                if host_name not in found:
                    found.append(host_name)
        except Exception as ex:
            _emit(on_status, u'  scan skipped {0}: {1}'.format(host_name, _text(ex)))
        finally:
            _close(family_doc)
    rows = []
    for name, host_names in hosts.items():
        if len(host_names) < 2:
            continue
        host_names.sort()
        rows.append((name, host_names))
    rows.sort(key=lambda item: (-len(item[1]), item[0]))
    return rows


def _export_shared(doc, name, host_name, out_dir):
    """Flag one nested family Shared and save it. Raises on failure."""
    host = _find_family(doc, host_name)
    if host is None:
        raise Exception(u'host was not found')
    host_doc = None
    nested_doc = None
    group = None
    try:
        host_doc = doc.EditFamily(host)
        nested = _direct_nested(host_doc, name)
        if nested is None or not nested.IsEditable:
            raise Exception(u'nested family is not editable in {0}'.format(host_name))
        nested_doc = host_doc.EditFamily(nested)
        group = DB.TransactionGroup(nested_doc, 'Share nested family')
        group.Start()
        transaction = quiet_transaction(DB.Transaction(nested_doc, 'Flag nested family shared'))
        transaction.Start()
        try:
            parameter = nested_doc.OwnerFamily.get_Parameter(DB.BuiltInParameter.FAMILY_SHARED)
            if parameter is None:
                raise Exception(u'family has no Shared parameter')
            parameter.Set(1)
            transaction.Commit()
            if transaction.GetStatus() != DB.TransactionStatus.Committed:
                raise Exception(u'flagging Shared was rolled back')
            _keep_group(group)
        except Exception:
            _rollback_if_open(transaction)
            _discard_group(group)
            raise
        base = _safe_name(name)
        path = os.path.join(out_dir, base + u'.rfa')
        suffix = 2
        while os.path.exists(path):
            path = os.path.join(out_dir, u'{0}_{1}.rfa'.format(base, suffix))
            suffix += 1
        _save_family(nested_doc, path)
        return path
    finally:
        _close(nested_doc)
        _close(host_doc)


def _load_nested(family_doc, name, path):
    """Load one shared family into a host. Roll the group back on any loss."""
    nested = _direct_nested(family_doc, name)
    if nested is None:
        raise Exception(u'not in this host')
    if _is_shared(nested):
        raise Exception(u'already shared')
    before_types = _symbol_names(family_doc, name)
    before_placed = _placed_count(family_doc, name)

    def _do_load():
        return family_doc.LoadFamily(path, _IncomingSharedLoadOptions())

    def _accept():
        after_placed = _placed_count(family_doc, name)
        missing, placed_lost = reload_loss(
            before_types, _symbol_names(family_doc, name),
            before_placed, after_placed)
        if missing or placed_lost:
            detail = u', '.join(missing) if missing else u''
            raise Exception(
                u'would drop types ({0}) or instances ({1} -> {2})'.format(
                    detail, before_placed, after_placed))

    _load_in_group(family_doc, _do_load, _accept)


def _load_host(doc, family_doc, host_name):
    """Load the host back. Roll back if the project loses instances or types."""
    before_instances = int(_instance_count(doc))
    before_host = _placed_count(doc, host_name)
    before_types = _symbol_names(doc, host_name)

    def _do_load():
        return family_doc.LoadFamily(doc, _IncomingSharedLoadOptions())

    def _accept():
        missing, host_lost = reload_loss(
            before_types, _symbol_names(doc, host_name),
            before_host, _placed_count(doc, host_name))
        after_instances = int(_instance_count(doc))
        if missing or host_lost or after_instances < before_instances:
            detail = u', '.join(missing) if missing else u''
            raise Exception(
                u'project would lose types ({0}) or instances (host {1} -> {2})'.format(
                    detail, before_host, _placed_count(doc, host_name)))

    _load_in_group(doc, _do_load, _accept)


def share_duplicated_nested(doc, out_dir, log_path=None, on_status=None):
    """Share nested families that are copied into two or more hosts.

    Writes one .rfa per shared family under out_dir. Does not save the project.
    A host that cannot accept the shared version is left as it was.
    """
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    started = time.time()
    instances_before = int(_instance_count(doc))
    categories_before = _category_counts(doc)
    found = duplicated_nested(doc, on_status)
    _emit(on_status, u'{0} nested families are copied into several hosts'.format(len(found)))

    paths = {}
    export_errors = []
    for name, host_names in found:
        _emit(on_status, u'Shared definition: {0} ({1} hosts)'.format(name, len(host_names)))
        try:
            paths[name] = _export_shared(doc, name, host_names[0], out_dir)
        except Exception as ex:
            export_errors.append({'name': name, 'error': _text(ex)})
            _emit(on_status, u'  export rolled back: {0}'.format(_text(ex)))

    by_host = {}
    for name, host_names in found:
        if name not in paths:
            continue
        for host_name in host_names:
            by_host.setdefault(host_name, []).append(name)
    host_names = sorted(by_host.keys())

    rows = []
    total = len(host_names)
    for index, host_name in enumerate(host_names):
        _emit(on_status, u'[{0}/{1}] {2}'.format(index + 1, total, host_name))
        host = _find_family(doc, host_name)
        row = {
            'host': host_name,
            'shared': [],
            'rolled_back': [],
            'host_rolled_back': False,
        }
        if host is None:
            row['error'] = u'host was not found'
            rows.append(row)
            continue
        family_doc = None
        host_started = time.time()
        try:
            family_doc = doc.EditFamily(host)
            prepared = []
            for name in by_host[host_name]:
                try:
                    _load_nested(family_doc, name, paths[name])
                    prepared.append(name)
                except Exception as ex:
                    row['rolled_back'].append({'name': name, 'reason': _text(ex)})
            if prepared:
                try:
                    _load_host(doc, family_doc, host_name)
                    row['shared'] = prepared
                except Exception as ex:
                    row['host_rolled_back'] = True
                    row['prepared'] = prepared
                    row['host_rollback_reason'] = _text(ex)
        except Exception as ex:
            row['error'] = _text(ex)
        finally:
            _close(family_doc)
        row['sec'] = round(time.time() - host_started, 1)
        if row['shared'] or row['rolled_back'] or row.get('error') or row['host_rolled_back']:
            rows.append(row)
            if log_path:
                try:
                    _append_json_line(log_path, row)
                except Exception:
                    pass
        shared_text = u', '.join(row['shared']) if row['shared'] else u'-'
        rolled_text = u', '.join(item['name'] for item in row['rolled_back']) or u'-'
        _emit(on_status, u'    shared: {0}  rolled back: {1}  {2}s'.format(
            shared_text, rolled_text, row['sec']))

    instances_after = int(_instance_count(doc))
    categories_after = _category_counts(doc)
    added = category_delta(categories_before, categories_after)
    updated = 0
    rejected = 0
    nested_rolled = 0
    for row in rows:
        if row.get('shared'):
            updated += 1
        if row.get('host_rolled_back'):
            rejected += 1
        nested_rolled += len(row.get('rolled_back') or [])
    return {
        'candidates': [{'name': name, 'hosts': len(host_names)} for name, host_names in found],
        'exported': sorted(paths.keys()),
        'export_errors': export_errors,
        'hosts': rows,
        'hosts_updated': updated,
        'hosts_rejected': rejected,
        'nested_rolled_back': nested_rolled,
        'instances_before': instances_before,
        'instances_after': instances_after,
        'instances_added': instances_after - instances_before,
        'category_delta': added,
        'seconds': round(time.time() - started, 1),
    }
