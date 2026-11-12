# -*- coding: utf-8 -*-
"""Write element elevation in meters, prefixed with +, into a multiline text parameter."""

__title__ = "+Height"
__author__ = "Byggstyrning AB"
__doc__ = "Write element elevation in meters (+15.25) from Project Base Point, Survey Point, or Internal Origin"

import sys
import os.path as op
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BasePoint,
    LocationPoint,
    StorageType,
    Family,
    FamilyInstance,
    IFamilyLoadOptions,
)

from pyrevit import script, forms, revit
from System.Collections.Generic import List
from System import String
from System.Windows import Visibility

script_path = __file__
extension_dir = op.dirname(op.dirname(op.dirname(op.dirname(op.dirname(script_path)))))
lib_path = op.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path)

from revit.revit_utils import (
    set_parameter_value,
    is_element_editable,
    is_parameter_writable,
)
from styles import load_styles_to_window
from revit.compat import (
    convert_from_internal_units,
    is_param_multiline_text,
    get_element_id_value,
    UnitTypeId,
)

logger = script.get_logger()

COORD_PROJECT_BASE_POINT = "Project Base Point"
COORD_SURVEY_POINT = "Survey Point"
COORD_INTERNAL_ORIGIN = "Internal Origin"
MARKER_FAMILY_NAME = u"Plush\u00f6jdsmark\u00f6r"

_PREFERRED_PARAM_NAMES = (u"plush\u00f6jd", u"plushojd", u"+height", u"plus height")


def to_meters(internal_feet):
    """Convert a length in Revit internal units (feet) to meters."""
    try:
        if UnitTypeId is not None:
            return convert_from_internal_units(internal_feet, UnitTypeId.Meters)
    except Exception:
        pass
    return internal_feet * 0.3048


def format_plus_height(meters):
    """Format meters as +15.25. Values below zero keep the minus sign."""
    rounded = round(float(meters), 2)
    if rounded == 0:
        rounded = 0.0
    if rounded < 0:
        return "{:.2f}".format(rounded)
    return "+{:.2f}".format(rounded)


def get_project_base_point(doc):
    """Return the project base point, or None."""
    for base_point in FilteredElementCollector(doc).OfClass(BasePoint):
        try:
            if not base_point.IsShared:
                return base_point
        except Exception:
            continue
    return None


def is_marker_instance(element):
    """True for a Plushöjdsmarkör instance, including one hosted on another family."""
    try:
        if not isinstance(element, FamilyInstance):
            return False
        symbol = element.Symbol
        if symbol is None or symbol.Family is None:
            return False
        return unicode(symbol.Family.Name) == MARKER_FAMILY_NAME
    except Exception:
        return False


def get_marker_base_point(instance):
    """Project position of this marker family's origin, not the host origin."""
    try:
        transform = instance.GetTransform()
        if transform is not None:
            return transform.Origin
    except Exception:
        pass
    try:
        location = instance.Location
        if isinstance(location, LocationPoint) and location.Point:
            return location.Point
    except Exception:
        pass
    return None


def elevation_internal(doc, point, coordinate_system, project_base_z):
    """Elevation in internal units for the chosen datum.

    Project Base Point: internal Z minus the base point Z.
    Survey Point: shared-coordinate elevation (same datum as spot elevation
    set to Survey Point).
    Internal Origin: raw Revit internal Z.
    """
    if coordinate_system == COORD_SURVEY_POINT:
        position = doc.ActiveProjectLocation.GetProjectPosition(point)
        return position.Elevation
    if coordinate_system == COORD_INTERNAL_ORIGIN:
        return point.Z
    return point.Z - project_base_z


def instance_chain(instance):
    """The marker, then each family it is nested inside."""
    chain = []
    current = instance
    seen = {}
    while current is not None:
        try:
            element_id = get_element_id_value(current.Id)
        except Exception:
            break
        if element_id in seen:
            break
        seen[element_id] = True
        chain.append(current)
        try:
            current = current.SuperComponent
        except Exception:
            break
    return chain


def writable_parameter_host(marker, parameter_name):
    """Instance that can store the height.

    A shared marker nested in the closet keeps plushöjd read-only and shows
    the closet parameter. The height is still the marker origin.
    """
    for element in instance_chain(marker):
        try:
            param = element.LookupParameter(parameter_name)
        except Exception:
            param = None
        if not param or param.IsReadOnly:
            continue
        if param.StorageType != StorageType.String:
            continue
        return element
    return None


def markers_in_view(doc, view):
    """Plushöjdsmarkör instances visible in the view, including shared nested ones."""
    view_ids = {}
    for element_id in FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElementIds():
        view_ids[get_element_id_value(element_id)] = True
    markers = []
    for element in FilteredElementCollector(doc).OfClass(FamilyInstance):
        if not is_marker_instance(element):
            continue
        in_view = False
        for item in instance_chain(element):
            if get_element_id_value(item.Id) in view_ids:
                in_view = True
                break
        if in_view:
            markers.append(element)
    return markers


def get_multiline_text_parameters(doc, view):
    """Writable multiline text parameters on the marker or the family it is nested in."""
    names = set()
    for marker in markers_in_view(doc, view):
        for element in instance_chain(marker):
            try:
                parameters = element.Parameters
            except Exception:
                continue
            for param in parameters:
                try:
                    if param.IsReadOnly or param.StorageType != StorageType.String:
                        continue
                    if is_param_multiline_text(param.Definition):
                        names.add(param.Definition.Name)
                except Exception:
                    continue
    return sorted(names)


def preferred_parameter(parameter_names):
    """Pick a likely plus-height parameter name when one is present."""
    for name in parameter_names:
        lowered = unicode(name).lower()
        for preferred in _PREFERRED_PARAM_NAMES:
            if preferred in lowered:
                return name
    return None


def collect_elements_with_parameter(doc, view, parameter_name):
    """Markers in the view whose height can be stored on the marker or its host."""
    elements = []
    for element in markers_in_view(doc, view):
        if writable_parameter_host(element, parameter_name) is not None:
            elements.append(element)
    return elements


def get_active_location_name(doc):
    """Name of the document's active project location (site)."""
    try:
        location = doc.ActiveProjectLocation
        if location and location.Name:
            return location.Name
    except Exception as ex:
        logger.debug("Active location name unavailable: {}".format(ex))
    return None


class _KeepMarkerLoadOptions(IFamilyLoadOptions):
    """Reload the marker family without overwriting placed parameter values."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True


def marker_family_path():
    """Bundled Plushöjdsmarkör.rfa next to this script."""
    return op.join(op.dirname(__file__), MARKER_FAMILY_NAME + u".rfa")


def marker_family_loaded(doc):
    """True when the marker family is already in the project."""
    for family in FilteredElementCollector(doc).OfClass(Family):
        try:
            if family.Name == MARKER_FAMILY_NAME:
                return True
        except Exception:
            continue
    return False


class CoordinateSystemWindow(forms.WPFWindow):
    """Z datum and multiline text parameter, on one dialog."""

    def __init__(self, location_name, parameter_names):
        xaml_file = op.join(op.dirname(__file__), 'CoordinateWindow.xaml')
        forms.WPFWindow.__init__(self, xaml_file)
        load_styles_to_window(self)
        self.selected_key = None
        self.selected_parameter = None
        if location_name:
            self.locationPill.Text = location_name
        else:
            self.locationRow.Visibility = Visibility.Collapsed
        choices = List[String]()
        choices.Add(COORD_PROJECT_BASE_POINT)
        choices.Add(COORD_SURVEY_POINT)
        choices.Add(COORD_INTERNAL_ORIGIN)
        self.coordinateCombo.ItemsSource = choices
        self.coordinateCombo.SelectedIndex = 0
        parameters = List[String]()
        preferred = preferred_parameter(parameter_names)
        preferred_index = 0
        for index, name in enumerate(parameter_names):
            parameters.Add(name)
            if preferred and name == preferred:
                preferred_index = index
        self.parameterCombo.ItemsSource = parameters
        self.parameterCombo.SelectedIndex = preferred_index

    def ok_click(self, sender, args):
        coordinate = self.coordinateCombo.SelectedItem
        parameter = self.parameterCombo.SelectedItem
        if coordinate:
            self.selected_key = str(coordinate)
        if parameter:
            self.selected_parameter = str(parameter)
        self.DialogResult = True
        self.Close()

    def load_marker_click(self, sender, args):
        doc = revit.doc
        path = marker_family_path()
        if not op.exists(path):
            forms.alert(
                u"Could not find {}.".format(MARKER_FAMILY_NAME),
                title="+Height",
            )
            return
        already_loaded = marker_family_loaded(doc)
        try:
            with revit.Transaction("Load plus height marker", doc):
                loaded = doc.LoadFamily(path, _KeepMarkerLoadOptions())
            if isinstance(loaded, tuple):
                loaded = loaded[0]
        except Exception as ex:
            logger.error("Load marker failed: {}".format(ex))
            forms.alert(
                u"Could not load {}.\n\n{}".format(MARKER_FAMILY_NAME, ex),
                title="+Height",
            )
            return
        if loaded:
            forms.show_balloon(
                header="+Height",
                text=u"{} loaded.".format(MARKER_FAMILY_NAME),
                is_new=True,
            )
        elif already_loaded or marker_family_loaded(doc):
            forms.show_balloon(
                header="+Height",
                text=u"{} is already loaded.".format(MARKER_FAMILY_NAME),
                is_new=True,
            )
        else:
            forms.alert(
                u"Could not load {}.".format(MARKER_FAMILY_NAME),
                title="+Height",
            )


def select_height_options(location_name, parameter_names):
    """Coordinate system and multiline text parameter."""
    window = CoordinateSystemWindow(location_name, parameter_names)
    window.ShowDialog()
    return window.selected_key, window.selected_parameter


def main():
    """Write +height for elements in the active view."""
    doc = revit.doc
    view = doc.ActiveView
    if not view:
        forms.alert("No active view found.", title="+Height")
        return

    location_name = get_active_location_name(doc)
    parameter_names = get_multiline_text_parameters(doc, view)
    if not parameter_names:
        forms.alert(
            "No writable multiline text parameters found on Plushöjdsmarkör in the active view.\n\n"
            "The target parameter must be Multiline Text, not single-line Text.",
            title="+Height",
        )
        return

    coordinate_system, height_parameter = select_height_options(location_name, parameter_names)
    if not coordinate_system or not height_parameter:
        return

    project_base_z = 0.0
    if coordinate_system == COORD_PROJECT_BASE_POINT:
        base_point = get_project_base_point(doc)
        if not base_point:
            forms.alert("Project Base Point was not found.", title="+Height")
            return
        try:
            project_base_z = base_point.Position.Z
        except Exception as ex:
            forms.alert(
                "Could not read Project Base Point elevation.\n\n{}".format(ex),
                title="+Height",
            )
            return

    elements = collect_elements_with_parameter(doc, view, height_parameter)
    if not elements:
        forms.alert(
            "No Plushöjdsmarkör instances with parameter \"{}\" found in the active view.".format(height_parameter),
            title="+Height",
        )
        return

    results = []
    failed_count = 0
    skipped_count = 0
    total_count = len(elements)
    batch_size = max(1, total_count // 100) if total_count > 100 else 10

    with forms.ProgressBar(title="Writing Heights ({} elements)".format(total_count)) as pb:
        for idx, element in enumerate(elements):
            if idx % batch_size == 0 or idx == total_count - 1:
                pb.update_progress(idx + 1, total_count)

            point = get_marker_base_point(element)
            if point is None:
                failed_count += 1
                continue

            target = writable_parameter_host(element, height_parameter)
            if target is None:
                skipped_count += 1
                continue

            is_editable, edit_reason = is_element_editable(doc, target)
            if not is_editable:
                skipped_count += 1
                continue

            try:
                internal_z = elevation_internal(doc, point, coordinate_system, project_base_z)
                height_text = format_plus_height(to_meters(internal_z))
            except Exception as ex:
                logger.error("Error reading elevation for element {}: {}".format(element.Id, ex))
                failed_count += 1
                continue

            results.append({
                'element': target,
                'height_text': height_text,
            })

    successful = 0
    write_failed = 0

    if results:
        with revit.Transaction("Write Plus Height", doc):
            for result in results:
                if set_parameter_value(result['element'], height_parameter, result['height_text']):
                    successful += 1
                else:
                    write_failed += 1

    lines = [
        "Updated {} elements.".format(successful),
        "Coordinate system: {}".format(coordinate_system),
    ]
    if location_name:
        lines.append("Active location: {}".format(location_name))
    lines.append("Parameter: {}".format(height_parameter))
    if skipped_count:
        lines.append("Skipped: {}".format(skipped_count))
    if failed_count:
        lines.append("Failed: {}".format(failed_count))
    if write_failed:
        lines.append("Write errors: {}".format(write_failed))
    summary = "\n".join(lines)
    if successful:
        forms.show_balloon(
            header="+Height",
            text="Updated {} elements.".format(successful),
            tooltip=summary,
            is_new=True,
        )
    else:
        forms.alert(summary, title="+Height")


if __name__ == '__main__':
    main()
