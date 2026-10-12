# -*- coding: utf-8 -*-
"""Create 3D Zone Generic Model family instances from FilledRegion boundaries.

Creates Generic Model family instances using the 3DZone.rfa template,
replacing the extrusion profile with each filled region's boundary loops.
Instances are placed in the active view respecting the view's phase.

Before creating, a checklist asks which of the regions' own instance parameters
(e.g. OP_Husdel, OP_Kalkylgrupp) should be copied onto the zones. The choice is
remembered between runs. A chosen parameter that the zone family lacks gets its
project binding extended to Generic Models so the value can land; anything that
still cannot be written is listed in a summary afterwards.

A second question sets the zones' Material: none, a material named like the
region type (e.g. "OOMB 800"), named after a region parameter's value (e.g.
OP_Kalkylgrupp -> "OOMB"), or built from a template with {Parameter} tags such
as "3Dzone({OP_Husdel})-{OP_Kalkylgrupp}", or picked explicitly per combination
of region values (e.g. OP_Husdel=300 + OP_Kalkylgrupp=KOMB -> "3Dzone(350)-KOMB")
when the project's material names follow no derivable rule. Explicit picks are
remembered and only new combinations are asked for. Names are matched ignoring
case and spaces, so an existing "3DZone(800)" or "3Dzone (800)" is reused; when a
templated name has no match, trailing "-part" segments are dropped one at a time
before a new material is created with a colour derived from its name.
"""

__title__ = "Create 3D Zones from Regions"
__author__ = "Byggstyrning AB"
__doc__ = ("Create Generic Model family instances from FilledRegion boundaries using the 3DZone.rfa "
           "template. Asks which region parameters (e.g. OP_Husdel) to copy onto the zones.")


# Import standard libraries
import sys
import json
import os.path as op

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit

# Add the extension directory to the path
logger_temp = script.get_logger()
script_dir = script.get_script_path()
# Calculate extension directory by going up from script directory
# Structure: extension_root/pyBS.tab/3D Zone.panel/col2.stack/3D Zones from Region.pushbutton/
pushbutton_dir = script_dir
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)  # Go up one more level to get extension root
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Initialize logger
logger = script.get_logger()

# Import shared modules
from zone3d.spatial_adapter import RegionAdapter
from zone3d.zone_creator import create_zones_from_spatial_elements


def show_region_filter_dialog(regions, doc):
    """Simple filter function that returns selected regions (no dialog).
    
    This function mimics the interface expected by create_zones_from_spatial_elements
    but doesn't show a dialog - just returns the regions directly.
    
    Args:
        regions: List of FilledRegion elements
        doc: Revit document
        
    Returns:
        List of FilledRegion elements (same as input)
    """
    # No dialog needed - just return the regions
    return regions


CONFIG_KEY = "region_copy_params"
MATERIAL_CONFIG_KEY = "region_material_source"


def _is_user_parameter(param):
    """True for shared/project parameters; False for Revit built-ins."""
    try:
        defn = param.Definition
        if defn is None:
            return False
        bip = getattr(defn, "BuiltInParameter", None)
        if bip is not None and bip != BuiltInParameter.INVALID:
            return False
        return True
    except Exception:
        return False


def _param_value_text(param):
    try:
        st = param.StorageType
        if st == StorageType.String:
            return param.AsString() or ""
        if st == StorageType.Integer:
            return str(param.AsInteger())
        if st == StorageType.Double:
            return param.AsValueString() or str(param.AsDouble())
        if st == StorageType.ElementId:
            return param.AsValueString() or ""
    except Exception:
        pass
    return ""


def collect_region_parameters(filled_regions):
    """Return {name: {"with_value": n, "sample": text, "storage": StorageType}} for the
    regions' own writable instance parameters (built-ins excluded)."""
    candidates = {}
    for region in filled_regions:
        for param in region.Parameters:
            try:
                if param.IsReadOnly or not _is_user_parameter(param):
                    continue
                if param.StorageType not in (StorageType.String, StorageType.Integer, StorageType.Double):
                    continue
                name = param.Definition.Name
                entry = candidates.setdefault(name, {"with_value": 0, "sample": "", "storage": param.StorageType})
                if param.HasValue:
                    text = _param_value_text(param)
                    if text:
                        entry["with_value"] += 1
                        if not entry["sample"]:
                            entry["sample"] = text
            except Exception:
                continue
    return candidates


def choose_parameters_to_copy(filled_regions, candidates):
    """Show a checklist of the regions' own instance parameters; return chosen names.

    Returns a list (possibly empty = copy nothing), or None if the user cancelled.
    The previous choice is remembered in the pyRevit user config.
    """
    if not candidates:
        forms.alert("The selected filled regions have no instance parameters of their own to copy.\n\n"
                    "Add a shared parameter (e.g. OP_Husdel) as an Instance parameter on the "
                    "Detail Items category, fill it on the regions, and run again.\n\n"
                    "The zones will be created without copied parameters.",
                    title="No parameters to copy")
        return []

    cfg = script.get_config()
    previous = cfg.get_option(CONFIG_KEY, [])
    if not isinstance(previous, list):
        previous = []

    class ParamItem(forms.TemplateListItem):
        @property
        def name(self):
            entry = candidates[self.item]
            label = "{}   ({}/{} regions have a value".format(self.item, entry["with_value"], len(filled_regions))
            if entry["sample"]:
                label += ", e.g. '{}'".format(entry["sample"])
            return label + ")"

    items = [ParamItem(n, checked=(n in previous)) for n in sorted(candidates)]
    chosen = forms.SelectFromList.show(
        items,
        title="Parameters to copy from region to 3D zone",
        button_name="Create 3D Zones",
        multiselect=True,
        width=650, height=450)
    if chosen is None:
        return None
    chosen = list(chosen)
    cfg.set_option(CONFIG_KEY, chosen)
    script.save_config()
    return chosen


MATERIAL_NONE = "No material (leave as is)"
MATERIAL_TYPE_NAME = "Region type name  (e.g. 'OOMB 800')"
MATERIAL_TEMPLATE = "Name template with {Parameter} tags  (e.g. '3Dzone({OP_Husdel})-{OP_Kalkylgrupp}')"
MATERIAL_MAP = "Pick per husdel/kalkylgrupp combination  (asks only for new combinations, remembered)"
MATERIAL_MAP_REVIEW = "Pick per combination, review all remembered choices"
MATERIAL_MAP_MARKER = "__map__"
MAP_CONFIG_KEY = "region_material_map"       # JSON: {"keys": [...], "map": {"300|KOMB": "..."}}
DEFAULT_TEMPLATE = "3Dzone({OP_Husdel})-{OP_Kalkylgrupp}"
PREFERRED_KEY_PARAMS = ["OP_Husdel", "OP_Kalkylgrupp"]


def _load_material_map(cfg):
    raw = cfg.get_option(MAP_CONFIG_KEY, "")
    try:
        data = json.loads(raw) if raw else {}
    except Exception:
        data = {}
    if not isinstance(data, dict):
        data = {}
    data.setdefault("keys", [])
    data.setdefault("map", {})
    return data


def _project_material_names(doc):
    names = []
    for mat in FilteredElementCollector(doc).OfClass(Material):
        try:
            names.append(mat.Name)
        except Exception:
            continue
    # zone materials first, then the rest, both alphabetical
    zone = sorted(n for n in names if "zone" in n.lower())
    rest = sorted(n for n in names if "zone" not in n.lower())
    return zone + rest


def _suggest_material(material_names, combo_values):
    """Best guess for a combination, only used to pre-select in the picker.

    First a name containing every value (shortest wins), then one containing just the
    first value (e.g. the husdel). With names like "3Dzone(350)-KOMB" for husdel 300
    there is often no guess at all, which is exactly why the user picks explicitly.
    """
    lookup = lambda s: "".join(s.split()).lower()
    for wanted in (combo_values, combo_values[:1]):
        best = None
        for name in material_names:
            key = lookup(name)
            if all(lookup(v) in key for v in wanted):
                if best is None or len(name) < len(best):
                    best = name
        if best:
            return best
    return None


def build_material_map(filled_regions, candidates, doc, review_all):
    """Explicit combination -> material mapping, chosen by the user and remembered.

    Returns {"keys": [...], "map": {...}} or None if cancelled.
    """
    cfg = script.get_config()
    data = _load_material_map(cfg)

    usable = [n for n in sorted(candidates) if candidates[n]["storage"] in (StorageType.String, StorageType.Integer)]
    keys = [k for k in data["keys"] if k in usable]
    if not keys:
        keys = [k for k in PREFERRED_KEY_PARAMS if k in usable]
    if not keys or review_all:
        picked = forms.SelectFromList.show(
            [forms.TemplateListItem(n, checked=(n in keys)) for n in usable],
            title="Which region parameters identify the material?",
            button_name="Continue", multiselect=True, width=500, height=350)
        if not picked:
            return None
        keys = list(picked)
    if keys != data["keys"]:
        data["keys"] = keys
        data["map"] = {}   # a different key set invalidates old combinations

    combos = {}
    for region in filled_regions:
        values = []
        for k in keys:
            p = region.LookupParameter(k)
            values.append(_param_value_text(p) if p is not None and p.HasValue else "")
        if all(values):
            combos["|".join(values)] = values

    material_names = _project_material_names(doc)
    if not material_names:
        forms.alert("The project has no materials to choose from.", title="Material per combination")
        return None

    to_ask = sorted(c for c in combos if review_all or c not in data["map"])
    for i, combo in enumerate(to_ask):
        values = combos[combo]
        current = data["map"].get(combo) or _suggest_material(material_names, values)
        ordered = list(material_names)
        if current in ordered:
            ordered.remove(current)
            ordered.insert(0, current)
        label = ", ".join("{}={}".format(k, v) for k, v in zip(keys, values))
        items = [forms.TemplateListItem(n, checked=(n == current)) for n in ordered]
        chosen = forms.SelectFromList.show(
            items,
            title="Material for {}   ({} of {})".format(label, i + 1, len(to_ask)),
            button_name="Use this material", multiselect=False, width=600, height=550)
        if chosen is None:
            return None
        data["map"][combo] = chosen

    cfg.set_option(MAP_CONFIG_KEY, json.dumps(data))
    script.save_config()
    return data


def choose_material_source(candidates, filled_regions, doc):
    """Ask where the zone material name should come from.

    Returns None if cancelled, "" for no material, RegionAdapter.MATERIAL_FROM_TYPE_NAME
    for the region type name, a region parameter name, a name template containing
    {Parameter} tags, or a {"keys", "map"} dict for an explicit mapping. Remembered.
    """
    string_params = sorted(n for n, e in candidates.items() if e["storage"] == StorageType.String)
    options = [MATERIAL_NONE, MATERIAL_MAP, MATERIAL_MAP_REVIEW, MATERIAL_TYPE_NAME, MATERIAL_TEMPLATE] + \
              ["Parameter {}".format(n) for n in string_params]

    cfg = script.get_config()
    previous = cfg.get_option(MATERIAL_CONFIG_KEY, "")
    default = MATERIAL_NONE
    if previous == MATERIAL_MAP_MARKER:
        default = MATERIAL_MAP
    elif previous == RegionAdapter.MATERIAL_FROM_TYPE_NAME:
        default = MATERIAL_TYPE_NAME
    elif previous and "{" in previous:
        default = MATERIAL_TEMPLATE
    elif previous and previous in string_params:
        default = "Parameter {}".format(previous)

    # Put the remembered choice first so Enter repeats the last run
    if default in options:
        options.remove(default)
        options.insert(0, default)

    picked = forms.CommandSwitchWindow.show(
        options,
        message="Material on the 3D zones: take the material name from ... "
                "(matched ignoring case and spaces; created if the project has none)")
    if picked is None:
        return None
    if picked == MATERIAL_NONE:
        source = ""
    elif picked in (MATERIAL_MAP, MATERIAL_MAP_REVIEW):
        mapping = build_material_map(filled_regions, candidates, doc, review_all=(picked == MATERIAL_MAP_REVIEW))
        if mapping is None:
            return None
        cfg.set_option(MATERIAL_CONFIG_KEY, MATERIAL_MAP_MARKER)
        script.save_config()
        return mapping
    elif picked == MATERIAL_TYPE_NAME:
        source = RegionAdapter.MATERIAL_FROM_TYPE_NAME
    elif picked == MATERIAL_TEMPLATE:
        template_default = previous if (previous and "{" in previous) else DEFAULT_TEMPLATE
        source = forms.ask_for_string(
            default=template_default,
            prompt="Material name template. Tags in braces are region parameters.\n"
                   "If no material matches the full name, trailing '-part' segments are\n"
                   "dropped one at a time: '3Dzone(800)-OOMB' falls back to '3Dzone(800)'.\n\n"
                   "Available: " + ", ".join("{" + n + "}" for n in string_params),
            title="Material name template")
        if source is None:
            return None
        source = source.strip()
        if "{" not in source:
            forms.alert("The template has no {Parameter} tag; no material will be set.",
                        title="Material name template")
            source = ""
    else:
        source = picked[len("Parameter "):]
    cfg.set_option(MATERIAL_CONFIG_KEY, source)
    script.save_config()
    return source


def report_parameter_copy(adapter, parameters_to_copy):
    """Summarise what landed on the zones, and warn about what did not."""
    rep = getattr(adapter, "copy_report", None)
    if rep is None:
        return
    if not parameters_to_copy and not adapter.material_source:
        return
    lines = []
    for name in parameters_to_copy:
        n = rep["copied"].get(name, 0)
        lines.append("{}: written on {} zone(s)".format(name, n))
    if rep["bound"]:
        lines.append("")
        lines.append("Binding extended to Generic Models (project parameter changed): " + ", ".join(sorted(rep["bound"])))
    if rep["materials"]:
        lines.append("")
        lines.append("Material set: " + ", ".join(
            "{} on {} zone(s)".format(k, v) for k, v in sorted(rep["materials"].items())))
    if rep["materials_created"]:
        lines.append("Materials created (adjust colour under Manage > Materials): " + ", ".join(sorted(rep["materials_created"])))
    problems = []
    if rep["material_param_missing"]:
        problems.append("The zone family has no writable '{}' parameter; material not set".format(adapter.material_param))
    if rep["materials_missing"]:
        problems.append("No material found or created for: " + ", ".join(sorted(rep["materials_missing"])))
    if rep.get("unmapped_combos"):
        problems.append("No material chosen for combination(s): " + ", ".join(sorted(rep["unmapped_combos"])) +
                        "  (run again and pick 'review all')")
    if rep["missing"]:
        problems.append("Not found on the 3D zone and could not be bound: " + ", ".join(sorted(rep["missing"])))
    if rep["type_mismatch"]:
        problems.append("Different data type on the 3D zone, not copied: " + ", ".join(sorted(rep["type_mismatch"])))
    if rep["no_value"]:
        problems.append("Empty on some regions, nothing written there: " + ", ".join(sorted(rep["no_value"])))
    if problems:
        lines.append("")
        lines.extend(problems)
    message = "\n".join(lines)
    logger.info(message)
    if problems:
        forms.alert(message, title="Parameters copied to 3D zones")
    else:
        try:
            script.get_output().print_md("**Parameters copied to 3D zones**\n\n" + "\n\n".join(lines))
        except Exception:
            pass


# --- Main Execution ---

if __name__ == '__main__':
    doc = revit.doc
    
    # Get selected elements
    selection = revit.get_selection()
    selected_ids = selection.element_ids
    
    if not selected_ids or len(selected_ids) == 0:
        forms.alert("Please select at least one FilledRegion element before running this tool.",
                   title="No Selection", exitscript=True)
    
    # Get selected elements and filter for FilledRegion
    selected_elements = [doc.GetElement(eid) for eid in selected_ids]
    filled_regions = [elem for elem in selected_elements if isinstance(elem, FilledRegion)]
    
    if not filled_regions:
        forms.alert("No FilledRegion elements found in selection.\n\nPlease select FilledRegion elements and try again.",
                   title="Invalid Selection", exitscript=True)
    
    logger.debug("Found {} FilledRegion element(s) in selection".format(len(filled_regions)))
    
    # Get view from first region (all regions should be in the same view)
    # Use this view for phase handling
    active_view = None
    if filled_regions:
        try:
            first_region = filled_regions[0]
            view_id = first_region.OwnerViewId
            if view_id and view_id != ElementId.InvalidElementId:
                active_view = doc.GetElement(view_id)
        except Exception as e:
            logger.debug("Error getting view from region: {}".format(e))
    
    # Fallback to active view if region view not found
    if not active_view:
        try:
            active_view = revit.active_view
            if not active_view:
                active_view = doc.ActiveView
        except:
            pass
    
    if not active_view:
        logger.warning("Could not determine view - phase may not be set correctly")
    else:
        logger.debug("Using view '{}' (ID: {}) for phase handling".format(
            active_view.Name if hasattr(active_view, 'Name') else 'Unknown',
            active_view.Id if active_view else 'None'))
    
    # Let the user pick which region parameters travel to the zones, and where the
    # zone material name comes from
    candidates = collect_region_parameters(filled_regions)
    parameters_to_copy = choose_parameters_to_copy(filled_regions, candidates)
    if parameters_to_copy is None:
        script.exit()
    material_source = choose_material_source(candidates, filled_regions, doc)
    if material_source is None:
        script.exit()

    # Create adapter with view for phase handling and the chosen parameter list
    adapter = RegionAdapter(active_view=active_view, parameters_to_copy=parameters_to_copy,
                            material_source=material_source or None)

    # Create zones using shared orchestration function
    # Note: We bypass the dialog by providing a simple filter function
    success_count, fail_count, failed_elements, created_instance_ids = create_zones_from_spatial_elements(
        filled_regions,
        doc,
        adapter,
        extension_dir,
        pushbutton_dir,
        show_region_filter_dialog,
        "Regions"
    )

    report_parameter_copy(adapter, parameters_to_copy)

    # Select newly created instances
    if created_instance_ids:
        try:
            from System.Collections.Generic import List
            uidoc = revit.uidoc
            element_ids = List[ElementId]()
            for instance_id in created_instance_ids:
                element_ids.Add(instance_id)
            
            uidoc.Selection.SetElementIds(element_ids)
            logger.debug("Selected {} newly created zone instance(s)".format(len(created_instance_ids)))
        except Exception as select_error:
            logger.debug("Error selecting created instances: {}".format(select_error))
            # Don't fail the script if selection fails

