# -*- coding: utf-8 -*-
"""In-place family purge for the Model Health button.

Recursive bottom-up purge of unused elements inside loaded families, then an
optional project Purge Unused. Families are edited and loaded back in place.

Strict quiet mode rolls back a family-document failure whose resolution would
delete elements or a type, so a nested reload cannot drop geometry to get past
an error. Nested types that placed instances select through instance parameters
are kept.

This is separate from the optional Family Purify pass in analyze_filesize.py.
The button saves a new file first; nothing here writes back to the original path.
"""

import os
import time
from collections import Counter

import clr
import System

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit import DB
from System.Collections.Generic import HashSet


_QUIET_KEY = 'pybs_family_purge_quiet'

KEEP_CLASSES = (
    DB.Level,
    DB.Grid,
    DB.View,
    DB.ViewSheet,
    DB.RevitLinkType,
    DB.CADLinkType,
    DB.ImageType,
)


def _text(value):
    try:
        return unicode(value)
    except Exception:
        return u'?'


def _state():
    st = System.AppDomain.CurrentDomain.GetData(_QUIET_KEY)
    if st is None:
        st = {
            'on': False,
            'strict': False,
            'log': [],
            'counts': {},
            'handlers': None,
        }
        System.AppDomain.CurrentDomain.SetData(_QUIET_KEY, st)
    return st


def _note(kind, text):
    st = _state()
    key = u'{0}: {1}'.format(kind, text)
    st['counts'][key] = st['counts'].get(key, 0) + 1
    if len(st['log']) < 2000:
        st['log'].append(key)


def _would_delete(message):
    """True when accepting this failure makes Revit delete elements or a type."""
    try:
        if message.HasResolutions():
            resolution = message.GetDefaultResolutionType()
            if resolution in (
                    DB.FailureResolutionType.DeleteElements,
                    DB.FailureResolutionType.FixElements):
                return True
    except Exception:
        pass
    try:
        text = message.GetDescriptionText() or u''
    except Exception:
        return False
    return (
        (u'Delete Element' in text)
        or (u'Delete Type' in text)
        or (u"Can't make type" in text)
    )


def _swallow(accessor):
    """Shared failure policy. Returns the FailureProcessingResult to use."""
    messages = list(accessor.GetFailureMessages())
    if not messages:
        return DB.FailureProcessingResult.Continue
    try:
        fdoc = accessor.GetDocument()
        where = fdoc.Title
        is_family = fdoc.IsFamilyDocument
    except Exception:
        where = u'?'
        is_family = False
    if _state().get('strict') and is_family and any(_would_delete(m) for m in messages):
        for message in messages:
            try:
                _note(u'rolled-back', u'{0}  [{1}]'.format(
                    message.GetDescriptionText(), where))
            except Exception:
                pass
        return DB.FailureProcessingResult.ProceedWithRollBack
    resolved = False
    for message in messages:
        try:
            text = u'{0}  [{1}]'.format(message.GetDescriptionText(), where)
        except Exception:
            text = u'?'
        severity = message.GetSeverity()
        if severity == DB.FailureSeverity.Warning:
            accessor.DeleteWarning(message)
            _note(u'warning', text)
        elif message.HasResolutions():
            accessor.ResolveFailure(message)
            resolved = True
            _note(u'error-resolved', text)
        else:
            _note(u'error-unresolvable', text)
    if resolved:
        return DB.FailureProcessingResult.ProceedWithCommit
    return DB.FailureProcessingResult.Continue


class _SwallowFailures(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, accessor):
        if not _state().get('on'):
            return DB.FailureProcessingResult.Continue
        return _swallow(accessor)


def quiet_transaction(transaction):
    """Attach the failure preprocessor before Start()."""
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(_SwallowFailures())
    options.SetClearAfterRollback(True)
    options.SetForcedModalHandling(False)
    transaction.SetFailureHandlingOptions(options)
    return transaction


def _on_failures(sender, args):
    if not _state().get('on'):
        return
    try:
        accessor = args.GetFailuresAccessor()
        result = _swallow(accessor)
        if result != DB.FailureProcessingResult.Continue:
            args.SetProcessingResult(result)
    except Exception as ex:
        _note(u'handler-error', _text(ex))


def _on_dialog(sender, args):
    # Handlers can outlive quiet_off if IronPython fails to unsubscribe.
    # When quiet is off they must not answer dialogs.
    if not _state().get('on'):
        return
    try:
        dialog_id = getattr(args, 'DialogId', u'') or u''
        message = getattr(args, 'Message', u'') or u''
        if u'Save' in dialog_id:
            args.OverrideResult(7)  # IDNO: never save from a prompt
        else:
            try:
                args.OverrideResult(1)  # IDOK
            except Exception:
                args.OverrideResult(6)  # IDYES
        _note(u'dialog', (dialog_id + u' ' + message)[:200])
    except Exception as ex:
        _note(u'dialog-error', _text(ex))


def quiet_on(uiapp):
    """Start answering failures and dialogs.

    The Revit event handlers are registered once per session. Turning quiet off
    only clears the flag, so a leaked IronPython subscription cannot keep
    dismissing dialogs after the command finishes, and a second run does not
    subscribe again.
    """
    st = _state()
    if not st.get('subscribed'):
        uiapp.Application.FailuresProcessing += _on_failures
        try:
            uiapp.DialogBoxShowing += _on_dialog
        except Exception:
            try:
                uiapp.Application.FailuresProcessing -= _on_failures
            except Exception:
                pass
            raise
        st['subscribed'] = True
    st['on'] = True
    return u'quiet on'


def quiet_off(uiapp):
    st = _state()
    st['on'] = False
    st['strict'] = False
    return u'quiet off'


def quiet_strict(on=True):
    """Roll back family-document failures that would delete elements or types."""
    _state()['strict'] = bool(on)
    return u'strict {0}'.format(u'on' if on else u'off')


def quiet_report(top=25, reset=False):
    st = _state()
    rows = sorted(st['counts'].items(), key=lambda kv: -kv[1])[:top]
    if rows:
        lines = [u'{0:5d}  {1}'.format(count, key) for key, count in rows]
        text = u'\n'.join(lines)
    else:
        text = u'(nothing swallowed)'
    if reset:
        st['counts'] = {}
        st['log'] = []
    return text


class _KeepValuesLoadOptions(DB.IFamilyLoadOptions):
    """Overwrite the family, keep existing parameter values.

    Shared nested families keep the version already in the target.
    """

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source,
                            overwriteParameterValues):
        source.Value = DB.FamilySource.Project
        overwriteParameterValues.Value = False
        return True


def _id_value(element_id):
    """Numeric id. Revit 2026 removed ElementId.IntegerValue."""
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def _load_ok(result):
    if isinstance(result, tuple):
        return bool(result[0])
    return bool(result)


def _id_list(ids):
    values = []
    if ids is None:
        return values
    for element_id in ids:
        values.append(element_id)
    return values


def _drop_protected(fdoc, ids, protect):
    """Keep types and families that project instances still select."""
    family_names = set(pair[0] for pair in protect)
    keep = HashSet[DB.ElementId]()
    for element_id in ids:
        element = fdoc.GetElement(element_id)
        if isinstance(element, DB.FamilySymbol) and (element.FamilyName, element.Name) in protect:
            continue
        if isinstance(element, DB.Family) and element.Name in family_names:
            continue
        keep.Add(element_id)
    return keep


def purge_unused(fdoc, max_pass=12, protect=None):
    """Delete what the family document itself reports as unused."""
    total = 0
    for _ in range(max_pass):
        ids = fdoc.GetUnusedElements(HashSet[DB.ElementId]())
        if protect:
            ids = _drop_protected(fdoc, ids, protect)
        if ids is None or ids.Count == 0:
            break
        transaction = quiet_transaction(DB.Transaction(fdoc, 'Family purge unused'))
        transaction.Start()
        deleted = 0
        try:
            try:
                fdoc.Delete(ids)
                deleted = ids.Count
            except Exception:
                for element_id in _id_list(ids):
                    try:
                        fdoc.Delete(element_id)
                        deleted += 1
                    except Exception:
                        pass
            transaction.Commit()
            if transaction.GetStatus() != DB.TransactionStatus.Committed:
                raise Exception(u'purge transaction rolled back')
        except Exception:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
            raise
        total += deleted
        if deleted == 0:
            break
    return total


def _nested_to_slim(fdoc):
    """Non-shared editable model families still loaded in this family."""
    ids = []
    for nested in DB.FilteredElementCollector(fdoc).OfClass(DB.Family):
        category = nested.FamilyCategory
        if not nested.Name or not nested.IsEditable:
            continue
        if category is not None and category.CategoryType != DB.CategoryType.Model:
            continue
        shared = nested.get_Parameter(DB.BuiltInParameter.FAMILY_SHARED)
        if shared and shared.AsInteger():
            continue
        ids.append(nested.Id)
    return ids


def project_famtype_refs(doc, family):
    """Nested (family, type) pairs selected by placed instances of `family`."""
    refs = set()
    for symbol_id in family.GetFamilySymbolIds():
        instance_filter = DB.FamilyInstanceFilter(doc, symbol_id)
        instances = DB.FilteredElementCollector(doc).OfClass(
            DB.FamilyInstance).WherePasses(instance_filter)
        for instance in instances:
            for parameter in instance.Parameters:
                if parameter.StorageType != DB.StorageType.ElementId:
                    continue
                definition = parameter.Definition
                if isinstance(definition, DB.InternalDefinition) \
                        and definition.BuiltInParameter != DB.BuiltInParameter.INVALID:
                    continue
                element = doc.GetElement(parameter.AsElementId())
                if isinstance(element, DB.FamilySymbol) \
                        and _id_value(element.Family.Id) != _id_value(family.Id):
                    refs.add((element.FamilyName, element.Name))
    return refs


def slim_family_doc(fdoc, stats, depth=0, max_depth=6, protect=None, on_status=None):
    stats['docs'] = stats.get('docs', 0) + 1
    stats['purged'] = stats.get('purged', 0) + purge_unused(fdoc, protect=protect)
    if depth >= max_depth:
        return stats
    for nested_id in _nested_to_slim(fdoc):
        nested = fdoc.GetElement(nested_id)
        if nested is None:
            continue
        nested_name = nested.Name
        nested_doc = None
        try:
            _emit(on_status, u'{0} / {1}'.format(_text(fdoc.Title), _text(nested_name)))
            nested_doc = fdoc.EditFamily(nested)
            slim_family_doc(
                nested_doc, stats, depth + 1, max_depth, on_status=on_status)
            try:
                loaded = nested_doc.LoadFamily(fdoc, _KeepValuesLoadOptions())
                if _load_ok(loaded):
                    stats['reloaded'] = stats.get('reloaded', 0) + 1
                else:
                    stats.setdefault('kept_original_nested', []).append(nested_name)
            except Exception:
                stats.setdefault('kept_original_nested', []).append(nested_name)
        except Exception as ex:
            stats.setdefault('nested_errors', []).append(
                u'{0}: {1}'.format(nested_name, _text(ex)))
        finally:
            if nested_doc is not None:
                try:
                    nested_doc.Close(False)
                except Exception:
                    pass
    stats['purged'] += purge_unused(fdoc, protect=protect)
    return stats


def _find_family(doc, name):
    for family in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if family.Name == name and family.IsEditable and not family.IsInPlace:
            return family
    return None


def slim_project_family(doc, family, protect=None, on_status=None):
    """Slim one loaded family. Unchanged families are not loaded back."""
    started = time.time()
    name = family.Name
    stats = {'name': name, 'purged': 0, 'reloaded': 0, 'loaded': False}
    family_doc = None
    try:
        family_doc = doc.EditFamily(family)
        slim_family_doc(family_doc, stats, protect=protect, on_status=on_status)
        if stats.get('purged') or stats.get('reloaded'):
            loaded = family_doc.LoadFamily(doc, _KeepValuesLoadOptions())
            stats['loaded'] = _load_ok(loaded)
            if not stats['loaded']:
                stats['error'] = u'load back into the project returned false'
        else:
            stats['unchanged'] = True
    except Exception as ex:
        stats['error'] = _text(ex)
        stats['loaded'] = False
    finally:
        if family_doc is not None:
            try:
                family_doc.Close(False)
            except Exception:
                pass
    stats['sec'] = round(time.time() - started, 1)
    return stats


def _family_category_name(family):
    category = family.FamilyCategory
    if category is not None and category.Name:
        return category.Name
    return u''


def family_instance_counts(doc):
    """Placed instance count per family name in this model."""
    symbol_names = {}
    for symbol in DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol):
        try:
            family = symbol.Family
            name = family.Name if family is not None else None
            if not name:
                continue
            symbol_names[_id_value(symbol.Id)] = name
        except Exception:
            continue
    counts = {}
    for instance in DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance):
        try:
            name = symbol_names.get(_id_value(instance.GetTypeId()))
        except Exception:
            continue
        if not name:
            continue
        counts[name] = counts.get(name, 0) + 1
    return counts


def loadable_families(doc):
    """Editable non-inplace families as (name, category, instance count)."""
    counts = family_instance_counts(doc)
    rows = []
    seen = set()
    for family in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if not family.Name or not family.IsEditable or family.IsInPlace:
            continue
        if family.Name in seen:
            continue
        seen.add(family.Name)
        rows.append((
            family.Name,
            _family_category_name(family),
            counts.get(family.Name, 0),
        ))
    rows.sort(key=lambda item: (_text(item[1]).lower(), _text(item[0]).lower()))
    return rows


def _candidate_names(doc):
    names = []
    seen = set()
    duplicates = []
    for family in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if not family.Name or not family.IsEditable or family.IsInPlace:
            continue
        if family.Name in seen:
            duplicates.append(family.Name)
            continue
        seen.add(family.Name)
        names.append(family.Name)
    names.sort()
    return names, duplicates


def _named_family_count(doc):
    count = 0
    for family in DB.FilteredElementCollector(doc).OfClass(DB.Family):
        if family.Name:
            count += 1
    return count


def _instance_count(doc):
    return DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).GetElementCount()


def _emit(callback, message):
    if callback is None:
        return
    try:
        callback(message)
    except Exception:
        pass


def _progress(callback, current, total, message):
    if callback is None:
        return
    try:
        callback(current, total, message)
    except Exception:
        pass


def purge_project(doc, max_pass=15, keep_classes=KEEP_CLASSES, on_progress=None):
    """Project Purge Unused, repeated, never deleting keep_classes.

    Only meaningful in a model that still has its placed instances. In a model
    whose geometry was peeled off, every family counts as unused.
    """
    counts = Counter()
    for pass_index in range(max_pass):
        _progress(
            on_progress, 0, 0,
            u'Purging unused project content, pass {0}'.format(pass_index + 1))
        found = doc.GetUnusedElements(HashSet[DB.ElementId]())
        ids = HashSet[DB.ElementId]()
        kinds = Counter()
        for element_id in _id_list(found):
            element = doc.GetElement(element_id)
            if element is None or isinstance(element, keep_classes):
                continue
            ids.Add(element_id)
            kinds[element.GetType().Name] += 1
        if ids.Count == 0:
            break
        transaction = quiet_transaction(DB.Transaction(doc, 'Family purge project'))
        transaction.Start()
        deleted = 0
        try:
            try:
                doc.Delete(ids)
                deleted = ids.Count
                counts.update(kinds)
            except Exception:
                for element_id in _id_list(ids):
                    element = doc.GetElement(element_id)
                    kind = element.GetType().Name if element is not None else '?'
                    try:
                        doc.Delete(element_id)
                        deleted += 1
                        counts[kind] += 1
                    except Exception:
                        pass
            transaction.Commit()
            if transaction.GetStatus() != DB.TransactionStatus.Committed:
                raise Exception(u'project purge rolled back')
        except Exception:
            if transaction.HasStarted() and not transaction.HasEnded():
                transaction.RollBack()
            raise
        if deleted == 0:
            break
    return dict(counts)


def purge(doc, log_path=None, on_status=None, do_project_purge=True,
          family_names=None, on_progress=None):
    """Slim editable loaded families, then optionally purge the project.

    family_names limits the run to those families. None means every editable
    loaded family. Caller turns quiet mode on first and off afterwards,
    including around the compact save. Returns a result dict. Does not save
    the document.
    """
    before = begin_counts(doc)
    names, duplicates = _candidate_names(doc)
    if family_names is not None:
        wanted = set(family_names)
        names = [name for name in names if name in wanted]
    total = len(names)
    started = time.time()
    rows = []
    _emit(on_status, u'Slimming {0} loadable families'.format(total))
    _progress(on_progress, 0, total, u'Slimming {0} families'.format(total))

    def _note(message):
        _emit(on_status, message)
        _progress(on_progress, _note_state['index'], total or 1, message)

    _note_state = {'index': 0}
    for index, name in enumerate(names):
        _note_state['index'] = index + 1
        _note(u'Purging {0}'.format(name))
        family = _find_family(doc, name)
        if family is None:
            stats = {'name': name, 'error': u'family was not found', 'loaded': False}
        else:
            protect = project_famtype_refs(doc, family)
            stats = slim_project_family(
                doc, family, protect=protect, on_status=_note)
            if protect:
                stats['protected'] = [
                    u'{0} : {1}'.format(pair[0], pair[1]) for pair in sorted(protect)
                ]
        rows.append(stats)
        if log_path:
            _append_json_line(log_path, stats)
        if stats.get('error'):
            detail = u'    ERROR {0}'.format(stats.get('error'))
        elif stats.get('unchanged'):
            detail = u'    unchanged  {0}s'.format(stats.get('sec') or 0)
        else:
            kept = stats.get('kept_original_nested') or []
            detail = u'    purged={0} reloaded={1} kept={2}  {3}s'.format(
                stats.get('purged') or 0,
                stats.get('reloaded') or 0,
                len(kept),
                stats.get('sec') or 0)
        _note(detail.strip())
    project_counts = None
    if do_project_purge:
        _emit(on_status, u'Purging unused items in the project')
        project_counts = purge_project(doc, on_progress=on_progress)
    return {
        'families': rows,
        'family_total': total,
        'duplicate_family_names': duplicates,
        'instances_before': before['instances'],
        'instances_after': _instance_count(doc),
        'families_before': before['families'],
        'families_after': _named_family_count(doc),
        'project_purge': project_counts,
        'seconds': round(time.time() - started, 1),
    }


def begin_counts(doc):
    return {
        'instances': _instance_count(doc),
        'families': _named_family_count(doc),
    }


def file_mb(path):
    try:
        return round(os.path.getsize(path) / 1048576.0, 3)
    except Exception:
        return None


def save_project_copy(doc, path):
    """Save as a new file. Workshared models become a new central.

    Refuses to overwrite an existing file or the path of the open model.
    """
    if os.path.exists(path):
        raise Exception(u'Refusing to overwrite an existing file: {0}'.format(path))
    current = doc.PathName or u''
    if current:
        same = os.path.normcase(os.path.abspath(current)) == os.path.normcase(
            os.path.abspath(path))
        if same:
            raise Exception(u'Refusing to overwrite the open model.')
    options = DB.SaveAsOptions()
    options.OverwriteExistingFile = False
    options.MaximumBackups = 1
    if doc.IsWorkshared:
        worksharing = DB.WorksharingSaveAsOptions()
        worksharing.SaveAsCentral = True
        options.SetWorksharingOptions(worksharing)
    doc.SaveAs(path, options)
    return file_mb(path)


def _remove_numbered_backups(path):
    root, ext = os.path.splitext(path)
    for index in range(1, 10):
        backup = u'{0}.{1:04d}{2}'.format(root, index, ext)
        if os.path.exists(backup):
            try:
                os.remove(backup)
            except Exception:
                pass


def save_compact(doc):
    """Compact-save the open document in place and drop numbered backups."""
    options = DB.SaveOptions()
    options.Compact = True
    doc.Save(options)
    path = doc.PathName
    if path:
        _remove_numbered_backups(path)
    return file_mb(path)


def to_json(value, indent=0):
    """JSON writer that escapes non-ASCII. IronPython json breaks on those names."""
    pad = u' ' * indent
    if value is None:
        return u'null'
    if value is True:
        return u'true'
    if value is False:
        return u'false'
    if isinstance(value, (int, long, float)):
        return unicode(value)
    if isinstance(value, dict):
        if not value:
            return u'{}'
        items = []
        for key, item in value.items():
            items.append(u'{0} {1}: {2}'.format(
                pad, to_json(_text(key)), to_json(item, indent + 1)))
        return u'{\n' + u',\n'.join(items) + u'\n' + pad + u'}'
    if isinstance(value, (list, tuple)):
        if not value:
            return u'[]'
        inner = u',\n'.join(
            pad + u' ' + to_json(item, indent + 1) for item in value)
        return u'[\n' + inner + u'\n' + pad + u']'
    text = _text(value)
    chars = []
    for char in text:
        code = ord(char)
        if char == u'"':
            chars.append(u'\\"')
        elif char == u'\\':
            chars.append(u'\\\\')
        elif code < 32 or code > 126:
            chars.append(u'\\u{0:04x}'.format(code))
        else:
            chars.append(char)
    return u'"' + u''.join(chars) + u'"'


_OUTPUT_NAMES = (u'family_purge.jsonl', u'summary.json', u'share_nested.jsonl')


def _writable_output(path):
    """Absolute path of a purge log the command is allowed to write.

    The file name is fixed. The folder must already exist. `abspath` collapses
    any `..` segments before that check.
    """
    if not path:
        raise Exception(u'Missing output path.')
    full = os.path.abspath(path)
    name = os.path.basename(full)
    allowed = False
    for item in _OUTPUT_NAMES:
        if name == item:
            allowed = True
            break
    if not allowed:
        raise Exception(u'Refusing to write {0}'.format(name))
    parent = os.path.dirname(full)
    if not parent or not os.path.isdir(parent):
        raise Exception(u'Output folder does not exist.')
    return full


def dump(obj, path):
    data = to_json(obj)
    with open(_writable_output(path), 'wb') as handle:
        handle.write(data.encode('utf-8'))


def _append_json_line(path, obj):
    line = to_json(obj).replace(u'\n', u' ') + u'\n'
    with open(_writable_output(path), 'ab') as handle:
        handle.write(line.encode('utf-8'))
