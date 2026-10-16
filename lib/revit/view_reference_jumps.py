# -*- coding: utf-8 -*-
"""Open the sheet for the selected 3D View Reference and zoom to its viewport.

Used by the View HUD. When the selection is one or more references that all
sit on the same sheet, the bar shows that sheet number. Clicking it opens
the sheet and frames the viewport of each selected view. The viewport is
found from the view the reference already stores, not from a saved viewport
id and not from a name. Any other selection hides the chip.
"""


import time


def chip_label_for_sheet_numbers(numbers):
    """'Sheet A-101' when every number is the same real sheet, else None."""
    unique = []
    seen = set()
    for number in numbers or []:
        raw = number or u""
        text = u" ".join(raw.replace(u"\r", u" ").replace(u"\n", u" ").split())
        if not text or text == u"-":
            continue
        if text in seen:
            continue
        seen.add(text)
        unique.append(text)
    if len(unique) != 1:
        return None
    return u"Sheet {}".format(unique[0])


def union_rects(first, second):
    """Smallest rectangle containing both (min_x, min_y, max_x, max_y) rects."""
    if first is None:
        return second
    if second is None:
        return first
    return (
        min(first[0], second[0]),
        min(first[1], second[1]),
        max(first[2], second[2]),
        max(first[3], second[3]),
    )


def padded_rectangle(min_x, min_y, max_x, max_y, fraction=0.08):
    """Grow a rectangle by a fraction of its size. None when it has no area."""
    width = max_x - min_x
    height = max_y - min_y
    if width <= 0 or height <= 0:
        return None
    pad_x = width * fraction
    pad_y = height * fraction
    return (min_x - pad_x, min_y - pad_y, max_x + pad_x, max_y + pad_y)


def zoom_rect_from_outlines(outlines, fraction=0.08):
    """One padded rectangle around every viewport outline, or None."""
    rect = None
    for outline in outlines or []:
        rect = union_rects(rect, outline)
    if rect is None:
        return None
    return padded_rectangle(rect[0], rect[1], rect[2], rect[3], fraction)


_selection_cache = {}
_pending_zoom = None
_idling_delegate = None
_ZOOM_WAIT_SECONDS = 3.0


def sheet_chip_text(document, uidoc):
    """Label for the current selection, or None when the chip should hide."""
    sheet = selected_reference_sheet(document, uidoc)
    if sheet is None:
        return None
    return chip_label_for_sheet_numbers([sheet.SheetNumber])


def open_selected_reference_sheet(uiapp):
    """Open the sheet and frame the viewports of the selected references."""
    try:
        uidoc = uiapp.ActiveUIDocument
        if uidoc is None:
            return
        sheet, views = _resolve_selected_target(uidoc.Document, uidoc)
        if sheet is None or not views:
            return
        _remember_zoom(sheet, views)
        active = uidoc.ActiveView
        already_open = active is not None and active.UniqueId == sheet.UniqueId
        if already_open and _zoom_pending(uidoc):
            _clear_zoom(uiapp)
            return
        _arm_zoom(uiapp)
        if not already_open:
            uidoc.RequestViewChange(sheet)
    except Exception:
        pass


def selected_reference_sheet(document, uidoc):
    """The single sheet shared by the selected references, or None."""
    if document is None or uidoc is None:
        return None
    try:
        if not uidoc.Document.Equals(document):
            return None
        selected_ids = tuple(sorted(
            _element_id_value(element_id)
            for element_id in uidoc.Selection.GetElementIds()
        ))
    except Exception:
        return None
    cache_key = (_document_cache_key(document), selected_ids)
    if cache_key in _selection_cache:
        unique_id = _selection_cache[cache_key]
        if not unique_id:
            return None
        sheet = document.GetElement(unique_id)
        if sheet is not None:
            return sheet
    sheet, _views = _resolve_selected_target(document, uidoc)
    _selection_cache.clear()
    _selection_cache[cache_key] = None if sheet is None else sheet.UniqueId
    return sheet


def _document_cache_key(document):
    try:
        if document.PathName:
            return document.PathName
    except Exception:
        pass
    try:
        return document.Title or "document"
    except Exception:
        return "document"


def _element_id_value(element_id):
    from revit.compat import get_element_id_value
    return get_element_id_value(element_id)


def _resolve_selected_target(document, uidoc):
    from revit.compat import get_element_id_value
    from revit.view_references import build_sheet_lookup
    lookup = None
    numbers = []
    sheets = []
    views = []
    for element_id in uidoc.Selection.GetElementIds():
        element = document.GetElement(element_id)
        view_id = _linked_view_unique_id(element)
        if not view_id:
            continue
        if lookup is None:
            lookup = build_sheet_lookup(document)
        view = document.GetElement(view_id)
        if view is None:
            continue
        sheet = lookup.get(get_element_id_value(view.Id))
        if sheet is None:
            continue
        numbers.append(sheet.SheetNumber)
        sheets.append(sheet)
        views.append(view)
    if chip_label_for_sheet_numbers(numbers) is None:
        return None, []
    return sheets[0], views


def _linked_view_unique_id(element):
    from Autodesk.Revit.DB import FamilyInstance
    from revit.view_references import FAMILY_NAME, ViewReferenceSchema
    if not isinstance(element, FamilyInstance):
        return None
    try:
        family = element.Symbol.Family
        if family is None or family.Name != FAMILY_NAME:
            return None
    except Exception:
        return None
    try:
        link = ViewReferenceSchema(element, update=False)
        unique_id = link.get("view_unique_id")
    except Exception:
        return None
    if not unique_id:
        return None
    return unique_id


def _remember_zoom(sheet, views):
    global _pending_zoom
    _pending_zoom = {
        "sheet_uid": sheet.UniqueId,
        "view_uids": [view.UniqueId for view in views],
        "since": time.time(),
    }


def _clear_zoom(uiapp):
    global _pending_zoom
    _pending_zoom = None
    _unhook_zoom(uiapp)


def _arm_zoom(uiapp):
    global _idling_delegate
    if _idling_delegate is not None:
        return
    from System import EventHandler
    from Autodesk.Revit.UI import Events as UIEvents
    _idling_delegate = EventHandler[UIEvents.IdlingEventArgs](_on_zoom_idling)
    uiapp.Idling += _idling_delegate


def _unhook_zoom(uiapp):
    global _idling_delegate
    if _idling_delegate is None or uiapp is None:
        _idling_delegate = None
        return
    try:
        uiapp.Idling -= _idling_delegate
    except Exception:
        pass
    _idling_delegate = None


def _on_zoom_idling(sender, args):
    pending = _pending_zoom
    if pending is None:
        _unhook_zoom(sender)
        return
    try:
        uidoc = sender.ActiveUIDocument
        active = None if uidoc is None else uidoc.ActiveView
        ready = (
            uidoc is not None
            and active is not None
            and active.UniqueId == pending["sheet_uid"]
        )
        if ready and _zoom_pending(uidoc):
            _clear_zoom(sender)
            return
    except Exception:
        pass
    if time.time() - pending["since"] >= _ZOOM_WAIT_SECONDS:
        _clear_zoom(sender)


def _zoom_pending(uidoc):
    pending = _pending_zoom
    if pending is None:
        return False
    try:
        document = uidoc.Document
        sheet = document.GetElement(pending["sheet_uid"])
        if sheet is None:
            return False
        views = []
        for unique_id in pending["view_uids"]:
            view = document.GetElement(unique_id)
            if view is not None:
                views.append(view)
        rect = _zoom_rect_for_views(document, sheet, views)
        if rect is None:
            return False
        uiview = _open_uiview(uidoc, sheet)
        if uiview is None:
            return False
        from Autodesk.Revit.DB import XYZ
        uiview.ZoomAndCenterRectangle(
            XYZ(rect[0], rect[1], 0.0),
            XYZ(rect[2], rect[3], 0.0),
        )
        return True
    except Exception:
        return False


def _zoom_rect_for_views(document, sheet, views):
    from revit.compat import get_element_id_value
    wanted = set(get_element_id_value(view.Id) for view in views)
    outlines = []
    for viewport_id in sheet.GetAllViewports():
        viewport = document.GetElement(viewport_id)
        if viewport is None:
            continue
        if get_element_id_value(viewport.ViewId) not in wanted:
            continue
        outline = _viewport_outline(viewport)
        if outline is not None:
            outlines.append(outline)
    return zoom_rect_from_outlines(outlines)


def _viewport_outline(viewport):
    """Sheet-space rectangle of the viewport, including its title when present."""
    try:
        box = viewport.GetBoxOutline()
        rect = (
            box.MinimumPoint.X, box.MinimumPoint.Y,
            box.MaximumPoint.X, box.MaximumPoint.Y,
        )
    except Exception:
        return None
    try:
        label = viewport.GetLabelOutline()
        rect = union_rects(rect, (
            label.MinimumPoint.X, label.MinimumPoint.Y,
            label.MaximumPoint.X, label.MaximumPoint.Y,
        ))
    except Exception:
        pass
    return rect


def _open_uiview(uidoc, sheet):
    from revit.compat import get_element_id_value
    sheet_id = get_element_id_value(sheet.Id)
    try:
        open_views = uidoc.GetOpenUIViews()
    except Exception:
        return None
    for uiview in open_views:
        try:
            if get_element_id_value(uiview.ViewId) == sheet_id:
                return uiview
        except Exception:
            continue
    return None
