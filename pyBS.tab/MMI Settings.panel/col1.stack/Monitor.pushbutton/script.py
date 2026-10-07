# -*- coding: utf-8 -*-
"""Toggles the MMI Monitor on and off.

When active, the monitor watches for relevant changes based on configuration.
When inactive, it does nothing.
"""

__title__ = "Monitor"
__author__ = "Byggstyrning AB"
__doc__ = "Toggle MMI Monitor on/off for the current session"
__highlight__ = 'updated'
# Keep the IronPython engine alive after this script exits so that
# our DocumentChanged / Sync event handlers retain their delegates and
# module-level globals (baseline set, caches) survive between edits.
__persistentengine__ = True

# Import standard libraries
import sys
import os
import re
import datetime

# Import Revit API
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System import EventHandler
from System.Collections.Generic import List
from Autodesk.Revit.DB.Events import DocumentChangedEventArgs, DocumentSynchronizingWithCentralEventArgs, DocumentSynchronizedWithCentralEventArgs

# Import pyRevit modules
from pyrevit import script
from pyrevit import forms
from pyrevit import revit
from pyrevit.userconfig import user_config
from pyrevit.coreutils.ribbon import ICON_MEDIUM
from pyrevit.revit import ui
import pyrevit.extensions as exts

# Add the extension directory to the path - FIXED PATH RESOLUTION
import os.path as op
script_path = __file__
pushbutton_dir = op.dirname(script_path)
splitpushbutton_dir = op.dirname(pushbutton_dir)
stack_dir = op.dirname(splitpushbutton_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Try direct import from current directory's parent path
sys.path.append(op.dirname(op.dirname(panel_dir)))

# Initialize logger
logger = script.get_logger()

# Import MMI libraries
from mmi.config import CONFIG_SECTION, CONFIG_KEY_ACTIVE
from mmi.config import (
    is_monitor_active,
    set_monitor_active,
    DEFAULT_PIN_MMI_LIMIT,
    DEFAULT_MOVE_MMI_LIMIT,
    DEFAULT_TYPE_INSTANCE_LIMIT,
    DEFAULT_INSTANCE_PARAM_LIMIT,
)
from mmi.core import get_mmi_parameter_name, load_monitor_config, get_default_mmi
from mmi.threshold import (
    at_or_above,
    include_in_instance_param_count,
    normalize_limit,
)
from mmi.utils import (
    get_element_location,
    get_element_mmi_value,
    validate_mmi_value,
    is_mmi_value_blank_for_default,
)
from revit.compat import get_element_id_value, make_element_id

# Import MMI Schema
try:
    from mmi.schema import MMIParameterSchema
except Exception as ex:
    logger.error("Failed to import MMI Schema: {}".format(ex))

# Event Handler for external events
class MMIEventHandler(IExternalEventHandler):
    def __init__(self):
        self.elements_to_pin = []
        self.notify_message = None
        self.mmi_threshold = DEFAULT_PIN_MMI_LIMIT
        self.elements_to_validate = []
        self.validate_corrections = {}
        self.undo_element_ids = None
        self.undo_already_posted = False
        
    def queue_undo_select(self, element_ids):
        """Remember elements to select after the next undo."""
        self.undo_element_ids = list(element_ids)

    def Execute(self, uiapp):
        try:
            logger.debug("MMI Event Handler Executing")
            undo_ids = self.undo_element_ids
            self.undo_element_ids = None
            if undo_ids:
                already_posted = bool(self.undo_already_posted)
                self.undo_already_posted = False
                self.elements_to_pin = []
                self.elements_to_validate = []
                self.validate_corrections = {}
                self.notify_message = None
                state = _UndoThenSelect(uiapp, undo_ids)
                if already_posted:
                    state.phase = "select"
                state.start()
                return
            
            # Check if monitor is still active
            if not is_monitor_active():
                logger.debug("Monitor is not active, clearing queued operations")
                # Clear any queued operations
                self.elements_to_pin = []
                self.elements_to_validate = []
                self.validate_corrections = {}
                self.notify_message = None
                return
            
            doc = uiapp.ActiveUIDocument.Document
            
            # Process validation corrections if any
            if self.elements_to_validate and self.validate_corrections:
                logger.debug("Processing MMI validation for {} elements".format(len(self.elements_to_validate)))
                with Transaction(doc, "Correct MMI Values") as t:
                    t.Start()
                    
                    correction_details = []
                    
                    for element_id, correction in self.validate_corrections.items():
                        element = doc.GetElement(element_id)
                        if not element:
                            continue
                            
                        orig_value = correction["original"]
                        fixed_value = correction["fixed"]
                        param_name = correction["param"]
                        
                        # Get the parameter
                        param = element.LookupParameter(param_name)
                        if not param:
                            # Try element type parameter
                            try:
                                type_id = element.GetTypeId()
                                if type_id and type_id != ElementId.InvalidElementId:
                                    element_type = doc.GetElement(type_id)
                                    if element_type:
                                        param = element_type.LookupParameter(param_name)
                            except Exception as e:
                                logger.debug("Error getting type parameter: {}".format(e))
                        
                        apply_skip = None
                        if not param:
                            apply_skip = "no_param"
                        elif param.IsReadOnly:
                            apply_skip = "readonly"
                        elif param.StorageType != StorageType.String:
                            apply_skip = "not_string_storage"
                        # Blank MMI: HasValue is often False; Set is still valid for writable string params.
                        if apply_skip is None:
                            try:
                                param.Set(str(fixed_value))
                                if correction.get("reason") != "default":
                                    correction_details.append("'{}' → '{}'".format(orig_value, fixed_value))
                                logger.debug("Corrected MMI value from '{}' to '{}' for element {}".format(
                                    orig_value, fixed_value, element_id))
                            except Exception as set_ex:
                                logger.debug(
                                    "MMI correction Set failed for {}: {}".format(element_id, set_ex))
                    
                    t.Commit()
                    
                    # Default MMI fills are applied quietly. Other corrections still announce.
                    if correction_details:
                        forms.show_balloon(
                            header="MMI Value Correction",
                            text="{} MMI values automatically corrected".format(len(correction_details)),
                            tooltip="Details:\n" + "\n".join(correction_details[:5]) +
                                   ("\n..." if len(correction_details) > 5 else ""),
                            is_new=True
                        )
                
                # Clear validation data
                self.elements_to_validate = []
                self.validate_corrections = {}
            
            # Process element pinning
            if self.elements_to_pin:
                logger.debug("Processing pin operation for {} elements".format(len(self.elements_to_pin)))
                
                # Pin the elements in a transaction
                with Transaction(doc, "Pin High MMI Elements") as t:
                    t.Start()
                    
                    pin_count = 0
                    for element in self.elements_to_pin:
                        element_id = element
                        
                        # Get the element from its ID
                        try:
                            element = doc.GetElement(element_id)
                            if element and hasattr(element, "Pinned") and not element.Pinned:
                                element.Pinned = True
                                pin_count += 1
                                logger.debug("Pinned element {}".format(element_id))
                        except Exception as elem_ex:
                            logger.error("Error pinning element {}: {}".format(element_id, elem_ex))
                        
                    t.Commit()
                
                # Show notification if requested
                if pin_count > 0 and self.notify_message:
                    # Use show_balloon instead of forms.alert
                    message = self.notify_message.format(pin_count, self.mmi_threshold)
                    forms.show_balloon(
                        header="MMI Monitor",
                        text=message,
                        tooltip="Elements with MMI value >= {} were automatically pinned".format(self.mmi_threshold),
                        is_new=True
                    )
                
                # Clear the queue
                self.elements_to_pin = []
                self.notify_message = None
            
        except Exception as ex:
            logger.error("Error in MMI Event Handler: {}".format(ex))
            
    def GetName(self):
        return "MMI Monitor Event Handler"
        
    def pin_elements_deferred(self, element_ids, threshold, notify=True):
        """Queue elements for pinning in a deferred execution"""
        self.elements_to_pin = element_ids
        self.mmi_threshold = threshold
        
        if notify:
            self.notify_message = "Pinned {} elements with MMI value >= {}"
        else:
            self.notify_message = None
            
    def validate_mmi_values_deferred(self, elements_to_validate, corrections):
        """Add elements to the pending MMI correction queue.

        Chain commands raise DocumentChanged once per segment, and ExternalEvent
        runs only after the command yields. Replacing the queue kept the last segment.
        """
        if not self.validate_corrections:
            self.validate_corrections = {}
        if not self.elements_to_validate:
            self.elements_to_validate = []
        seen = {}
        for existing_id in self.validate_corrections.keys():
            try:
                seen[get_element_id_value(existing_id)] = True
            except Exception:
                pass
        for element_id, correction in (corrections or {}).items():
            try:
                key = get_element_id_value(element_id)
            except Exception:
                key = None
            if key is not None and key in seen:
                continue
            self.validate_corrections[element_id] = correction
            self.elements_to_validate.append(element_id)
            if key is not None:
                seen[key] = True

# Global handlers and events. A second run of this persistent script must not
# drop the live delegates, or the previous DocumentChanged handler stays attached.
if not globals().get("_monitor_handlers_ready"):
    mmi_event_handler = None
    external_event = None
    doc_changed_handler = None
    doc_synchronizing_handler = None
    doc_synchronized_handler = None
    _monitor_handlers_ready = True
element_location_cache = {}  # Cache to store element locations for move detection
element_mmi_cache = {}  # Cache to store element MMI values to detect changes
element_mmi_blank = set()  # Element ids whose MMI was blank the last time we saw them
# Integer ElementId values: snapshot at monitor ON; ids not in this set are new post-activation
baseline_element_ids_for_default = set()
element_type_cache = {}  # element id int -> type id int
type_warn_last = {}  # type id int -> datetime of last type-change balloon
instance_param_warn_last = None
TYPE_WARN_DEBOUNCE_SECONDS = 5
INSTANCE_PARAM_WARN_DEBOUNCE_SECONDS = 5
_MONITOR_SUB_KEY = "_pybs_mmi_monitor_subs"
OWN_MONITOR_TRANSACTIONS = set([
    "Correct MMI Values",
    "Pin High MMI Elements",
])
CLICK_TO_UNDO = "Click to Undo."
_balloon_click_handlers = []
_pending_undo_select = None


def _selectable_ids(doc, element_ids, instances_only):
    chosen = List[ElementId]()
    seen = {}
    for element_id in element_ids:
        try:
            value = get_element_id_value(element_id)
        except Exception:
            continue
        if value in seen:
            continue
        element = doc.GetElement(element_id)
        if element is None:
            continue
        if instances_only and isinstance(element, ElementType):
            continue
        seen[value] = True
        chosen.Add(element.Id)
    return chosen


def _select_existing(uiapp, element_ids):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return 0
    doc = uidoc.Document
    chosen = _selectable_ids(doc, element_ids, False)
    if chosen.Count < 1:
        return 0
    try:
        uidoc.Selection.SetElementIds(chosen)
        return chosen.Count
    except Exception:
        instances = _selectable_ids(doc, element_ids, True)
        if instances.Count < 1:
            return 0
        uidoc.Selection.SetElementIds(instances)
        return instances.Count


class _UndoThenSelect(object):
    """Undo the latest change, then select the elements that change touched."""

    def __init__(self, uiapp, element_ids):
        self.uiapp = uiapp
        self.element_ids = list(element_ids)
        self.phase = "undo"
        self.attempts = 0
        self.handler = None

    def start(self):
        global _pending_undo_select
        self.handler = self._on_idling
        _pending_undo_select = self
        try:
            self.uiapp.Idling += self.handler
        except Exception as ex:
            logger.debug("Could not subscribe Idling for undo select: {}".format(ex))
            self.stop()

    def stop(self):
        global _pending_undo_select
        if self.handler is not None:
            try:
                self.uiapp.Idling -= self.handler
            except Exception:
                pass
            self.handler = None
        if _pending_undo_select is self:
            _pending_undo_select = None

    def _on_idling(self, sender, args):
        self.attempts += 1
        if self.phase == "undo":
            try:
                command_id = RevitCommandId.LookupPostableCommandId(PostableCommand.Undo)
                if self.uiapp.CanPostCommand(command_id):
                    self.uiapp.PostCommand(command_id)
                    self.phase = "select"
                    self.attempts = 0
                    return
            except Exception as ex:
                logger.debug("Undo command failed: {}".format(ex))
                self.stop()
                return
            if self.attempts > 30:
                logger.debug("Undo command was not postable")
                self.stop()
            return
        try:
            _select_existing(self.uiapp, self.element_ids)
        except Exception as ex:
            logger.debug("Selection after undo failed: {}".format(ex))
        self.stop()


def _post_undo(uiapp):
    """Post Revit's Undo command. Returns (posted, error)."""
    try:
        command_id = RevitCommandId.LookupPostableCommandId(PostableCommand.Undo)
    except Exception as ex:
        return False, "lookup: {}".format(ex)
    try:
        can_post = bool(uiapp.CanPostCommand(command_id))
    except Exception as ex:
        return False, "can: {}".format(ex)
    if not can_post:
        return False, "not_postable"
    try:
        uiapp.PostCommand(command_id)
    except Exception as ex:
        return False, "post: {}".format(ex)
    return True, ""


def _schedule_undo_and_select(element_ids):
    posted = False
    try:
        posted, post_error = _post_undo(__revit__)
        if post_error:
            logger.debug("Undo was not posted: {}".format(post_error))
    except Exception as ex:
        logger.debug("Undo post failed: {}".format(ex))
    try:
        if mmi_event_handler is None or external_event is None:
            if posted:
                state = _UndoThenSelect(__revit__, element_ids)
                state.phase = "select"
                state.start()
            return
        mmi_event_handler.undo_already_posted = posted
        mmi_event_handler.queue_undo_select(element_ids)
        external_event.Raise()
    except Exception as ex:
        logger.debug("Could not queue undo select: {}".format(ex))


def _show_undo_balloon(header, text, element_ids):
    import time
    now = time.time()
    last = getattr(sys, "_pybs_mmi_last_balloon", 0)
    if isinstance(last, float) and now - last < 1.0:
        return
    setattr(sys, "_pybs_mmi_last_balloon", now)
    ids = list(element_ids)
    def _on_click(sender, event_args):
        _schedule_undo_and_select(ids)
    _balloon_click_handlers.append(_on_click)
    forms.show_balloon(
        header=header,
        text="{}\n{}".format(text, CLICK_TO_UNDO),
        tooltip=CLICK_TO_UNDO,
        is_new=True,
        click_result=_on_click,
    )


def _remember_changed_element(changed_types, type_int, element_id):
    item = changed_types.get(type_int)
    if item is None:
        return
    item["ids"].append(element_id)


def _monitor_subscriptions():
    """Handlers stored on sys so a script rerun can detach the previous ones."""
    bag = getattr(sys, _MONITOR_SUB_KEY, None)
    if not isinstance(bag, dict):
        bag = {
            "doc_changed": [],
            "synchronizing": [],
            "synchronized": [],
        }
        setattr(sys, _MONITOR_SUB_KEY, bag)
    return bag


def _remember_subscription(key, handler):
    bag = _monitor_subscriptions()
    handlers = bag.get(key)
    if handlers is None:
        handlers = []
        bag[key] = handlers
    handlers.append(handler)


def _detach_monitor_subscriptions():
    """Remove every monitor handler this engine has subscribed."""
    bag = getattr(sys, _MONITOR_SUB_KEY, None)
    removed = 0
    if not bag:
        return removed
    try:
        app = __revit__.Application
    except Exception as ex:
        logger.debug("Could not detach monitor handlers: {}".format(ex))
        return removed
    pairs = (
        ("doc_changed", "DocumentChanged"),
        ("synchronizing", "DocumentSynchronizingWithCentral"),
        ("synchronized", "DocumentSynchronizedWithCentral"),
    )
    for key, event_name in pairs:
        event = getattr(app, event_name, None)
        for handler in list(bag.get(key) or []):
            try:
                event -= handler
                removed += 1
            except Exception:
                pass
        bag[key] = []
    return removed


def update_element_location_cache(element_id, location):
    """Update the element location cache."""
    global element_location_cache
    element_location_cache[get_element_id_value(element_id)] = {
        "location": location,
        "timestamp": datetime.datetime.now()
    }

def clean_element_location_cache():
    """Clean old entries from the element location cache."""
    global element_location_cache
    now = datetime.datetime.now()
    # Keep entries not older than 5 minutes
    element_location_cache = {
        id: data for id, data in element_location_cache.items()
        if (now - data["timestamp"]).total_seconds() < 300
    }

def populate_initial_location_cache(doc, min_mmi=None, cache_all=False):
    """Populate the location cache on monitor activation.

    min_mmi caches elements at or above that MMI value.
    cache_all caches every model element that has a location, so instance-parameter
    edits can be told apart from moves.
    """
    global element_location_cache
    try:
        mmi_param_name = None
        if not cache_all:
            mmi_param_name = get_mmi_parameter_name(doc)
            if not mmi_param_name:
                return
        
        logger.debug("Populating initial location cache...")
        
        # Get all elements in the model
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        
        cache_count = 0
        for element in all_elements:
            # Skip elements that can't be pinned
            if not hasattr(element, "Pinned"):
                continue
            if cache_all:
                if not element.Category or element.Category.CategoryType != CategoryType.Model:
                    continue
                current_location = get_element_location(element)
                if current_location:
                    update_element_location_cache(element.Id, current_location)
                    cache_count += 1
                continue
            
            # Get the MMI value for the element
            mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
            
            # Only cache elements at or above the move limit
            if mmi_value is not None and at_or_above(mmi_value, min_mmi):
                current_location = get_element_location(element)
                if current_location:
                    update_element_location_cache(element.Id, current_location)
                    cache_count += 1
        
        logger.debug("Cached locations for {} elements".format(cache_count))
        
    except Exception as ex:
        logger.error("Error populating initial location cache: {}".format(ex))


def populate_type_id_cache(doc):
    """Snapshot element id -> type id so later type-selector changes can be seen."""
    global element_type_cache
    try:
        element_type_cache = {}
        cache_count = 0
        for element in FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
            if not element.Category or element.Category.CategoryType != CategoryType.Model:
                continue
            try:
                type_id = element.GetTypeId()
            except Exception:
                continue
            if type_id is None or type_id == ElementId.InvalidElementId:
                continue
            element_type_cache[get_element_id_value(element.Id)] = get_element_id_value(type_id)
            cache_count += 1
        logger.debug("Cached type ids for {} elements".format(cache_count))
    except Exception as ex:
        logger.error("Error populating type id cache: {}".format(ex))

def populate_initial_mmi_cache(doc):
    """Populate the MMI cache with all elements on monitor activation.
    This allows us to detect MMI value changes."""
    global element_mmi_cache, element_mmi_blank
    try:
        mmi_param_name = get_mmi_parameter_name(doc)
        if not mmi_param_name:
            return
        
        logger.debug("Populating initial MMI cache...")
        
        # Get all elements in the model
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        
        element_mmi_cache = {}
        element_mmi_blank = set()
        cache_count = 0
        for element in all_elements:
            # Skip elements that can't be pinned
            if not hasattr(element, "Pinned"):
                continue
            
            # Get the MMI value for the element
            mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
            element_key = get_element_id_value(element.Id)
            # Cache all MMI values (both high and low). Blank is remembered separately.
            if mmi_value is not None:
                element_mmi_cache[element_key] = mmi_value
                cache_count += 1
            elif param is not None:
                element_mmi_blank.add(element_key)
        
        logger.debug("Cached MMI values for {} elements".format(cache_count))
        
    except Exception as ex:
        logger.error("Error populating initial MMI cache: {}".format(ex))


def populate_baseline_element_ids_for_default(doc):
    """Snapshot all current model element ids. New ids after monitor ON get default MMI (if enabled)."""
    global baseline_element_ids_for_default
    try:
        baseline_element_ids_for_default = set()
        for element in FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
            if not element.Category or element.Category.CategoryType != CategoryType.Model:
                continue
            if not hasattr(element, "Pinned"):
                continue
            baseline_element_ids_for_default.add(get_element_id_value(element.Id))
        logger.debug(
            "Baseline element ids for default MMI: {}".format(
                len(baseline_element_ids_for_default)))
    except Exception as ex:
        logger.error("Error populating baseline for default MMI: {}".format(ex))


def pin_all_high_mmi_elements(doc, pin_limit):
    """Proactively pin all high MMI elements when monitor activates.
    This prevents movement before it happens."""
    try:
        mmi_param_name = get_mmi_parameter_name(doc)
        if not mmi_param_name:
            return 0
        
        logger.debug("Scanning for high MMI elements to pin...")
        
        # Get all elements in the model
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        
        elements_to_pin = []
        for element in all_elements:
            # Skip elements that can't be pinned
            if not hasattr(element, "Pinned"):
                continue
            
            # Skip already pinned elements
            if element.Pinned:
                continue
            
            # Get the MMI value for the element
            mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
            
            # Only pin high MMI elements
            if mmi_value is not None and at_or_above(mmi_value, pin_limit):
                elements_to_pin.append(element)
        
        if not elements_to_pin:
            logger.debug("No unpinned high MMI elements found")
            return 0
        
        # Pin all elements in a single transaction
        with Transaction(doc, "Pin High MMI Elements") as t:
            t.Start()
            
            pinned_count = 0
            for element in elements_to_pin:
                try:
                    element.Pinned = True
                    pinned_count += 1
                except Exception as e:
                    logger.debug("Could not pin element {}: {}".format(element.Id, e))
            
            t.Commit()
        
        logger.debug("Proactively pinned {} high MMI elements".format(pinned_count))
        return pinned_count
        
    except Exception as ex:
        logger.error("Error in proactive pinning: {}".format(ex))
        return 0

def _transaction_names(args):
    try:
        names = args.GetTransactionNames()
    except Exception:
        return []
    if not names:
        return []
    result = []
    for name in names:
        result.append(str(name))
    return result


def _is_own_monitor_transaction(args):
    for name in _transaction_names(args):
        if name in OWN_MONITOR_TRANSACTIONS:
            return True
    return False


def _default_mmi_number(doc):
    raw = get_default_mmi(doc)
    if raw is None:
        return None
    text = str(raw).strip()
    if not text:
        return None
    try:
        return int(text)
    except Exception:
        try:
            return int(float(text))
        except Exception:
            return None


def _transaction_sets_default_mmi(args, default_number):
    if default_number is None:
        return False
    prefix = "Set MMI Parameter to "
    for name in _transaction_names(args):
        if not name.startswith(prefix):
            continue
        suffix = name[len(prefix):].strip()
        try:
            if int(suffix) == int(default_number):
                return True
        except Exception:
            if suffix == str(default_number):
                return True
    return False


def _mmi_changed_to_default(cache_key, mmi_value, default_number):
    """True when this edit wrote the configured default MMI."""
    if default_number is None or mmi_value is None:
        return False
    try:
        current = int(mmi_value)
    except Exception:
        return False
    if current != int(default_number):
        return False
    if cache_key in element_mmi_cache:
        return element_mmi_cache.get(cache_key) != current
    if cache_key in element_mmi_blank:
        return True
    return False


def _remember_mmi(cache_key, mmi_value):
    global element_mmi_cache, element_mmi_blank
    if mmi_value is None:
        element_mmi_cache.pop(cache_key, None)
        element_mmi_blank.add(cache_key)
        return
    element_mmi_blank.discard(cache_key)
    element_mmi_cache[cache_key] = mmi_value


def _recently_warned(stamp_map, key, now, seconds):
    previous = stamp_map.get(key)
    if previous is None:
        return False
    try:
        return (now - previous).total_seconds() < seconds
    except Exception:
        return False


def _parameter_text(element, built_in_param):
    try:
        param = element.get_Parameter(built_in_param)
    except Exception:
        return None
    if param is None:
        return None
    try:
        text = param.AsString()
    except Exception:
        return None
    if text:
        return text
    return None


def _element_type_name(element):
    """Return Family: Type. IronPython cannot read Element.Name directly."""
    if element is None:
        return "Type"
    name = None
    try:
        name = Element.Name.GetValue(element)
    except Exception:
        name = None
    if not name:
        name = _parameter_text(element, BuiltInParameter.SYMBOL_NAME_PARAM)
    if not name:
        name = _parameter_text(element, BuiltInParameter.ALL_MODEL_TYPE_NAME)
    family_name = None
    try:
        family = element.Family
        if family is not None:
            family_name = Element.Name.GetValue(family)
    except Exception:
        family_name = None
    if family_name and name and family_name != name:
        return "{}: {}".format(family_name, name)
    if name:
        return name
    try:
        return str(get_element_id_value(element.Id))
    except Exception:
        return "Type"


def count_instances_at_or_above(doc, type_id, mmi_param_name, limit):
    """Return ids of instances of type_id whose MMI is at or above limit."""
    if not mmi_param_name:
        return []
    try:
        type_elem = doc.GetElement(type_id)
        if type_elem is None:
            return []
        dependents = type_elem.GetDependentElements(None)
    except Exception as ex:
        logger.debug("Could not count instances: {}".format(ex))
        return []
    target = get_element_id_value(type_id)
    found = []
    for dep_id in dependents:
        element = doc.GetElement(dep_id)
        if element is None or isinstance(element, ElementType):
            continue
        try:
            element_type_id = element.GetTypeId()
        except Exception:
            continue
        if element_type_id is None or element_type_id == ElementId.InvalidElementId:
            continue
        if get_element_id_value(element_type_id) != target:
            continue
        mmi_value, _value_str, _param = get_element_mmi_value(element, mmi_param_name, doc)
        if mmi_value is not None and at_or_above(mmi_value, limit):
            found.append(element.Id)
    return found


def _queue_type_warning(doc, type_id, type_limit, changed_types, now, mmi_param_name):
    """Record a type once when it has instances at or above the MMI limit."""
    if type_id is None or type_id == ElementId.InvalidElementId:
        return
    type_int = get_element_id_value(type_id)
    if type_int in changed_types:
        return
    if _recently_warned(type_warn_last, type_int, now, TYPE_WARN_DEBOUNCE_SECONDS):
        return
    if not mmi_param_name:
        return
    instance_ids = count_instances_at_or_above(doc, type_id, mmi_param_name, type_limit)
    if not instance_ids:
        return
    type_elem = doc.GetElement(type_id)
    name = _element_type_name(type_elem) if type_elem is not None else str(type_int)
    changed_types[type_int] = {
        "name": name,
        "count": len(instance_ids),
        "ids": list(instance_ids),
    }


def _show_type_change_warning(changed_types, type_limit, now):
    if not changed_types:
        return
    items = list(changed_types.values())
    items.sort(key=lambda item: item["count"], reverse=True)
    if len(items) == 1:
        text = "Type '{}' changed ({} instances at or above MMI {})".format(
            items[0]["name"], items[0]["count"], type_limit)
    else:
        text = "{} types changed with instances at or above MMI {}".format(len(items), type_limit)
    select_ids = []
    for item in items:
        select_ids.extend(item.get("ids", []))
    _show_undo_balloon("Type Change", text, select_ids)
    for type_int in changed_types:
        type_warn_last[type_int] = now


def _show_instance_param_warning(edited_ids, param_limit, now):
    global instance_param_warn_last
    count = len(edited_ids)
    if count < 1:
        return
    if instance_param_warn_last is not None:
        try:
            elapsed = (now - instance_param_warn_last).total_seconds()
            if elapsed < INSTANCE_PARAM_WARN_DEBOUNCE_SECONDS:
                return
        except Exception:
            pass
    _show_undo_balloon(
        "Instance Parameter Edit",
        "{} elements at or above MMI {} had instance parameters edited".format(count, param_limit),
        edited_ids,
    )
    instance_param_warn_last = now


def warn_on_heavy_edits(doc, args, modified_element_ids, monitor_settings):
    """Balloon when a heavy type change or instance-parameter edit happens."""
    global element_type_cache
    type_enabled = bool(monitor_settings.get("warn_on_type_change", False))
    param_enabled = bool(monitor_settings.get("warn_on_instance_params", False))
    if not type_enabled and not param_enabled:
        return
    if _is_own_monitor_transaction(args):
        logger.debug("Skipping heavy-edit warnings for a monitor transaction")
        return

    default_number = _default_mmi_number(doc)
    setting_default_txn = _transaction_sets_default_mmi(args, default_number)
    type_limit = normalize_limit(
        monitor_settings.get("type_instance_limit"), DEFAULT_TYPE_INSTANCE_LIMIT)
    param_limit = normalize_limit(
        monitor_settings.get("instance_param_limit"), DEFAULT_INSTANCE_PARAM_LIMIT)
    mmi_param_name = get_mmi_parameter_name(doc)
    now = datetime.datetime.now()

    modified_types = []
    instances = []
    for element_id in modified_element_ids:
        element = doc.GetElement(element_id)
        if element is None:
            continue
        if isinstance(element, ElementType):
            modified_types.append(element)
        else:
            instances.append(element)

    modified_type_ids = set()
    changed_types = {}
    for element in modified_types:
        type_int = get_element_id_value(element.Id)
        modified_type_ids.add(type_int)
        if type_enabled:
            _queue_type_warning(doc, element.Id, type_limit, changed_types, now, mmi_param_name)
            _remember_changed_element(changed_types, type_int, element.Id)

    edited_ids = []
    for element in instances:
        if not element.Category or element.Category.CategoryType != CategoryType.Model:
            continue
        try:
            current_type_id = element.GetTypeId()
        except Exception:
            continue
        if current_type_id is None or current_type_id == ElementId.InvalidElementId:
            continue
        cache_key = get_element_id_value(element.Id)
        current_type_int = get_element_id_value(current_type_id)
        type_known = cache_key in element_type_cache
        type_changed = type_known and element_type_cache.get(cache_key) != current_type_int
        current_location = get_element_location(element)
        location_changed = False
        if current_location is not None and cache_key in element_location_cache:
            prev_location = element_location_cache[cache_key]["location"]
            try:
                location_changed = current_location.DistanceTo(prev_location) > 0.1
            except Exception:
                location_changed = False
        regenerated = current_type_int in modified_type_ids
        if type_enabled and type_changed:
            previous_type_int = element_type_cache.get(cache_key)
            _queue_type_warning(
                doc, current_type_id, type_limit, changed_types, now, mmi_param_name)
            if previous_type_int is not None:
                _queue_type_warning(
                    doc, make_element_id(previous_type_int), type_limit, changed_types, now,
                    mmi_param_name)
        if current_type_int in changed_types:
            _remember_changed_element(changed_types, current_type_int, element.Id)
        if type_changed:
            previous_type_int = element_type_cache.get(cache_key)
            if previous_type_int in changed_types:
                _remember_changed_element(changed_types, previous_type_int, element.Id)
        if param_enabled and type_known and include_in_instance_param_count(
                type_changed, location_changed, regenerated):
            # No stored location means this element cannot be a move we missed.
            if current_location is None or cache_key in element_location_cache:
                mmi_value = None
                mmi_param = None
                if mmi_param_name:
                    mmi_value, _value_str, mmi_param = get_element_mmi_value(
                        element, mmi_param_name, doc)
                if mmi_value is not None and at_or_above(mmi_value, param_limit):
                    wrote_default = setting_default_txn or _mmi_changed_to_default(
                        cache_key, mmi_value, default_number)
                    if not wrote_default:
                        edited_ids.append(element.Id)
                if mmi_param is not None:
                    _remember_mmi(cache_key, mmi_value)
        element_type_cache[cache_key] = current_type_int

    if type_enabled:
        _show_type_change_warning(changed_types, type_limit, now)
    if param_enabled:
        _show_instance_param_warning(edited_ids, param_limit, now)


def _refresh_tracked_locations(doc, modified_element_ids):
    """Store current locations after this change so the next event can see moves."""
    for element_id in modified_element_ids:
        element = doc.GetElement(element_id)
        if element is None or isinstance(element, ElementType):
            continue
        current_location = get_element_location(element)
        if current_location is not None:
            update_element_location_cache(element.Id, current_location)


def document_changed_handler(sender, args):
    """Handler for document changed event."""
    try:
        # Check if we should monitor
        if not is_monitor_active():
            return
        
        doc = args.GetDocument()
        modified_element_ids = args.GetModifiedElementIds()
        added_element_ids = args.GetAddedElementIds()
        mod_count = modified_element_ids.Count if modified_element_ids else 0
        add_count = added_element_ids.Count if added_element_ids else 0
        
        if mod_count == 0 and add_count == 0:
            return

        monitor_settings = load_monitor_config(doc, use_display_names=False)
        if mod_count > 0:
            warn_on_heavy_edits(doc, args, modified_element_ids, monitor_settings)
        
        # Get MMI parameter name
        mmi_param_name = get_mmi_parameter_name(doc)
        if not mmi_param_name:
            logger.warning("No MMI parameter name configured. Use MMI Config tool first.")
            if monitor_settings.get("warn_on_instance_params", False) and mod_count > 0:
                _refresh_tracked_locations(doc, modified_element_ids)
            return
        
        default_mmi = get_default_mmi(doc)
        
        validate_enabled = monitor_settings["validate_mmi"]
        warn_on_move_enabled = monitor_settings["warn_on_move"]
        pin_elements_enabled = monitor_settings["pin_elements"]
        
        has_modified_features = validate_enabled or warn_on_move_enabled or pin_elements_enabled
        toggle_default_on_new = bool(monitor_settings.get("default_on_new_instances", False))
        has_default_on_new = toggle_default_on_new and bool(default_mmi and str(default_mmi).strip())
        
        if not has_modified_features and not has_default_on_new:
            logger.debug("No MMI monitor features enabled and no default on new instances. Skipping.")
            if monitor_settings.get("warn_on_instance_params", False) and mod_count > 0:
                _refresh_tracked_locations(doc, modified_element_ids)
            return

        pin_limit = normalize_limit(
            monitor_settings.get("pin_mmi_limit"), DEFAULT_PIN_MMI_LIMIT)
        move_limit = normalize_limit(
            monitor_settings.get("move_mmi_limit"), DEFAULT_MOVE_MMI_LIMIT)
        
        global baseline_element_ids_for_default
        elements_to_validate = []
        validation_corrections = {}
        elements_to_pin = []
        moved_high_mmi_elements = []
        
        # ----- Added elements: Default on new instances (GetAddedElementIds) -----
        if has_default_on_new and add_count > 0:
            logger.debug(
                "Processing {} added elements for default MMI (param: {})".format(
                    add_count, mmi_param_name))
            for element_id in added_element_ids:
                element = doc.GetElement(element_id)
                if element is None or not hasattr(element, "Pinned"):
                    continue
                if not element.Category or element.Category.CategoryType != CategoryType.Model:
                    continue
                mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
                if param is None:
                    continue
                if mmi_value is not None:
                    continue
                if not is_mmi_value_blank_for_default(mmi_value, value_str):
                    continue
                if element_id in validation_corrections:
                    continue
                elements_to_validate.append(element_id)
                validation_corrections[element_id] = {
                    "original": "(empty)",
                    "fixed": str(default_mmi).strip(),
                    "param": mmi_param_name,
                    "reason": "default",
                }
                logger.debug(
                    "New instance (added) {} queued for default MMI {}".format(element_id, default_mmi))
            for element_id in added_element_ids:
                baseline_element_ids_for_default.add(get_element_id_value(element_id))
        
        # ----- Modified elements: new ids not in baseline (e.g. some walls only in modified set) -----
        if has_default_on_new and mod_count > 0:
            for element_id in modified_element_ids:
                eid_i = get_element_id_value(element_id)
                if eid_i in baseline_element_ids_for_default:
                    continue
                element = doc.GetElement(element_id)
                if element and hasattr(element, "Pinned") and element.Category and element.Category.CategoryType == CategoryType.Model:
                    mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
                    if (param
                            and mmi_value is None
                            and is_mmi_value_blank_for_default(mmi_value, value_str)
                            and element_id not in validation_corrections):
                        elements_to_validate.append(element_id)
                        validation_corrections[element_id] = {
                            "original": "(empty)",
                            "fixed": str(default_mmi).strip(),
                            "param": mmi_param_name,
                            "reason": "default",
                        }
                        logger.debug(
                            "New instance (modified) {} queued for default MMI {}".format(
                                element_id, default_mmi))
                baseline_element_ids_for_default.add(eid_i)
        
        # ----- Modified elements: validate / warn / pin (unchanged) -----
        if has_modified_features and mod_count > 0:
            logger.debug("Processing {} modified elements with MMI parameter: {}".format(
                mod_count, mmi_param_name))
            clean_element_location_cache()
            for element_id in modified_element_ids:
                element = doc.GetElement(element_id)
                
                if element is None or not hasattr(element, "Pinned"):
                    continue
                    
                mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
                
                if mmi_value is not None:
                    if validate_enabled and param:
                        orig_value, fixed_value = validate_mmi_value(value_str)
                        if orig_value and fixed_value:
                            elements_to_validate.append(element_id)
                            validation_corrections[element_id] = {
                                "original": orig_value,
                                "fixed": fixed_value,
                                "param": mmi_param_name
                            }
                            logger.debug("Element {} needs MMI value correction: '{}' to '{}'".format(
                                element_id, orig_value, fixed_value))
                    
                    if warn_on_move_enabled and at_or_above(mmi_value, move_limit):
                        current_location = get_element_location(element)
                        if current_location:
                            cache_key = get_element_id_value(element_id)
                            if cache_key in element_location_cache:
                                prev_location = element_location_cache[cache_key]["location"]
                                distance = current_location.DistanceTo(prev_location)
                                if distance > 0.1:
                                    moved_high_mmi_elements.append({
                                        "id": element_id,
                                        "mmi": mmi_value,
                                        "distance": distance
                                    })
                                    logger.debug("High MMI Element {} moved {:.2f} meters".format(
                                        element_id, distance))
                            update_element_location_cache(element_id, current_location)
                    
                    if pin_elements_enabled and at_or_above(mmi_value, pin_limit):
                        element_id_int = get_element_id_value(element_id)
                        prev_mmi = element_mmi_cache.get(element_id_int)
                        
                        should_pin = False
                        
                        if not element.Pinned:
                            if prev_mmi is None:
                                should_pin = True
                                logger.debug("Element {} newly detected with MMI {} - queuing for pin".format(
                                    element_id, mmi_value))
                            elif (not at_or_above(prev_mmi, pin_limit)) and at_or_above(mmi_value, pin_limit):
                                should_pin = True
                                logger.debug("Element {} MMI changed from {} to {} - queuing for pin".format(
                                    element_id, prev_mmi, mmi_value))
                        
                        if should_pin:
                            elements_to_pin.append(element_id)
                        
                        element_mmi_cache[element_id_int] = mmi_value
                    elif mmi_value is not None:
                        element_mmi_cache[get_element_id_value(element_id)] = mmi_value
        
        # Process validation if needed
        if elements_to_validate and validation_corrections and mmi_event_handler and external_event:
            mmi_event_handler.validate_mmi_values_deferred(elements_to_validate, validation_corrections)
            external_event.Raise()
            logger.debug("Queued {} elements for MMI validation correction".format(len(elements_to_validate)))
        
        # Process move warnings if needed
        if moved_high_mmi_elements:
            # Group and limit to top 5 highest MMI elements
            moved_high_mmi_elements.sort(key=lambda x: x["mmi"], reverse=True)
            count = len(moved_high_mmi_elements)
            top_elements = moved_high_mmi_elements[:5]
            
            details = []
            for item in top_elements:
                details.append("Element ID: {} (MMI: {}, Distance: {:.2f}m)".format(
                    get_element_id_value(item["id"]), item["mmi"], item["distance"]))
            
            tooltip = "High MMI elements should be carefully managed:\n" + "\n".join(details)
            if count > 5:
                tooltip += "\n... and {} more".format(count - 5)
            
            forms.show_balloon(
                header="High MMI Element Move",
                text="{} elements with MMI >= {} were moved".format(count, move_limit),
                tooltip=tooltip,
                is_new=True
            )
            logger.debug("Warned about {} moved high MMI elements".format(count))
        
        # Process pinning if needed
        if elements_to_pin and mmi_event_handler and external_event:
            # If already doing validation, avoid raising another external event immediately
            # The pin operation will be scheduled after validation completes
            if not elements_to_validate:
                mmi_event_handler.pin_elements_deferred(elements_to_pin, pin_limit)
                external_event.Raise()
                logger.debug("Queued {} elements for pinning using external event".format(len(elements_to_pin)))
            else:
                # Store the pinning request - it will be processed after validation
                mmi_event_handler.pin_elements_deferred(elements_to_pin, pin_limit)
                logger.debug("Pinning of {} elements will occur after validation".format(len(elements_to_pin)))

        if monitor_settings.get("warn_on_instance_params", False) and mod_count > 0:
            _refresh_tracked_locations(doc, modified_element_ids)
    
    except Exception as ex:
        logger.error("Error in document changed handler: {}".format(ex))

def document_synchronizing_handler(sender, args):
    """Handler for document synchronizing event - capture pre-sync state."""
    try:
        # Check if sync checking is enabled
        if not is_monitor_active():
            return
            
        # For sync events, sender is Application, get document from args or active document
        try:
            # Try to get document from event args first
            doc = args.Document
        except:
            # Fall back to active document from application
            app = sender  # sender is Application for sync events
            doc = app.ActiveUIDocument.Document if app.ActiveUIDocument else None
            
        if not doc:
            logger.warning("Could not get document from sync event")
            return
            
        monitor_settings = load_monitor_config(doc, use_display_names=False)
        
        if not monitor_settings.get("check_mmi_after_sync", False):
            logger.debug("Post-sync MMI checking is disabled")
            return
            
        logger.debug("Document synchronizing - tracking user elements for post-sync check")
        
        # Import sync checker and track elements
        from mmi.sync_checker import track_modified_elements_before_sync
        track_modified_elements_before_sync(doc)
        
    except Exception as ex:
        logger.error("Error in document synchronizing handler: {}".format(ex))

def document_synchronized_handler(sender, args):
    """Handler for document synchronized event - check MMI post-sync."""
    try:
        # Check if sync checking is enabled
        if not is_monitor_active():
            return
            
        # For sync events, sender is Application, get document from args or active document
        try:
            # Try to get document from event args first
            doc = args.Document
        except:
            # Fall back to active document from application
            app = sender  # sender is Application for sync events
            doc = app.ActiveUIDocument.Document if app.ActiveUIDocument else None
            
        if not doc:
            logger.warning("Could not get document from sync event")
            return
            
        monitor_settings = load_monitor_config(doc, use_display_names=False)
        
        if not monitor_settings.get("check_mmi_after_sync", False):
            logger.debug("Post-sync MMI checking is disabled")
            return
            
        logger.debug("Document synchronized - processing post-sync MMI check")
        
        # Import sync checker and process check
        from mmi.sync_checker import process_post_sync_check
        process_post_sync_check(doc)
        
    except Exception as ex:
        logger.error("Error in document synchronized handler: {}".format(ex))

def register_event_handlers():
    """Register the necessary event handlers for monitoring."""
    global mmi_event_handler, external_event, doc_changed_handler, doc_synchronizing_handler, doc_synchronized_handler
    try:
        _detach_monitor_subscriptions()
        doc_changed_handler = None
        doc_synchronizing_handler = None
        doc_synchronized_handler = None
        mmi_event_handler = MMIEventHandler()
        external_event = ExternalEvent.Create(mmi_event_handler)
        logger.debug("MMI Event Handler Created.")

        app = revit.doc.Application
        doc_changed_handler = EventHandler[DocumentChangedEventArgs](document_changed_handler)
        app.DocumentChanged += doc_changed_handler
        _remember_subscription("doc_changed", doc_changed_handler)
        logger.debug("Document Changed Handler registered.")

        doc_synchronizing_handler = EventHandler[DocumentSynchronizingWithCentralEventArgs](document_synchronizing_handler)
        app.DocumentSynchronizingWithCentral += doc_synchronizing_handler
        _remember_subscription("synchronizing", doc_synchronizing_handler)
        logger.debug("Document Synchronizing Handler registered.")

        doc_synchronized_handler = EventHandler[DocumentSynchronizedWithCentralEventArgs](document_synchronized_handler)
        app.DocumentSynchronizedWithCentral += doc_synchronized_handler
        _remember_subscription("synchronized", doc_synchronized_handler)
        logger.debug("Document Synchronized Handler registered.")
        logger.debug("MMI Monitor event registration completed.")
        return True
    except Exception as e:
        logger.error("Failed to register MMI Monitor events: {}".format(e))
        return False

def deregister_event_handlers():
    """Deregister event handlers."""
    global mmi_event_handler, external_event, doc_changed_handler, doc_synchronizing_handler, doc_synchronized_handler
    try:
        _detach_monitor_subscriptions()
        app = revit.doc.Application
        if doc_changed_handler is not None:
            try:
                app.DocumentChanged -= doc_changed_handler
            except Exception:
                pass
            doc_changed_handler = None
            logger.debug("Document Changed Handler unregistered.")

        if doc_synchronizing_handler is not None:
            try:
                app.DocumentSynchronizingWithCentral -= doc_synchronizing_handler
            except Exception:
                pass
            doc_synchronizing_handler = None
            logger.debug("Document Synchronizing Handler unregistered.")

        if doc_synchronized_handler is not None:
            try:
                app.DocumentSynchronizedWithCentral -= doc_synchronized_handler
            except Exception:
                pass
            doc_synchronized_handler = None
            logger.debug("Document Synchronized Handler unregistered.")

        if mmi_event_handler is not None:
            mmi_event_handler.elements_to_pin = []
            mmi_event_handler.elements_to_validate = []
            mmi_event_handler.validate_corrections = {}
            mmi_event_handler.notify_message = None
            logger.debug("Cleared all queued MMI operations")

        external_event = None
        mmi_event_handler = None
        logger.debug("MMI Monitor event deregistration completed.")
        return True
    except Exception as e:
        logger.error("Failed to deregister MMI Monitor events: {}".format(e))
        return False

# --- Button Initialization --- 

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    """Initialize the button icon based on the current active state."""
    try:
        # Use the same approach as Tab Coloring script
        on_icon = ui.resolve_icon_file(script_cmp.directory, exts.DEFAULT_ON_ICON_FILE)
        off_icon = ui.resolve_icon_file(script_cmp.directory, exts.DEFAULT_OFF_ICON_FILE)

        button_icon = script_cmp.get_bundle_file(
            on_icon if is_monitor_active() else off_icon
        )
        ui_button_cmp.set_icon(button_icon, icon_size=ICON_MEDIUM)
    except Exception as e:
        logger.error("Error initializing MMI Monitor button: {}".format(e))

# --- Main Execution --- 

if __name__ == '__main__':
    was_active = is_monitor_active()
    new_active_state = not was_active

    success = False
    if new_active_state:
        # Activate: Register handlers
        logger.debug("Activating MMI Monitor...")
        if register_event_handlers():
            set_monitor_active(True)
            script.toggle_icon(new_active_state)  # Toggle icon to active state
           
            # Get the current settings
            monitor_settings = load_monitor_config(revit.doc, use_display_names=False)
            mmi_param_name = get_mmi_parameter_name(revit.doc) or "Not set"
            
            pin_limit = normalize_limit(
                monitor_settings.get("pin_mmi_limit"), DEFAULT_PIN_MMI_LIMIT)
            move_limit = normalize_limit(
                monitor_settings.get("move_mmi_limit"), DEFAULT_MOVE_MMI_LIMIT)
            type_limit = normalize_limit(
                monitor_settings.get("type_instance_limit"), DEFAULT_TYPE_INSTANCE_LIMIT)
            param_limit = normalize_limit(
                monitor_settings.get("instance_param_limit"), DEFAULT_INSTANCE_PARAM_LIMIT)
            
            # Populate initial location cache if warn_on_move is enabled.
            # Instance-parameter warnings need locations for every model element.
            if monitor_settings.get("warn_on_instance_params", False):
                populate_initial_location_cache(revit.doc, cache_all=True)
            elif monitor_settings["warn_on_move"]:
                populate_initial_location_cache(revit.doc, min_mmi=move_limit)

            if (monitor_settings.get("warn_on_type_change", False)
                    or monitor_settings.get("warn_on_instance_params", False)):
                populate_type_id_cache(revit.doc)
            
            # Populate initial MMI cache if pin_elements is enabled (to detect changes)
            if monitor_settings["pin_elements"] or monitor_settings.get("warn_on_instance_params", False):
                populate_initial_mmi_cache(revit.doc)
            
            # Baseline model element ids: post-activation ids are treated as new for default MMI
            populate_baseline_element_ids_for_default(revit.doc)
            
            # Proactively pin all high MMI elements if pin_elements is enabled
            pinned_count = 0
            if monitor_settings["pin_elements"]:
                pinned_count = pin_all_high_mmi_elements(revit.doc, pin_limit)
            
            # Create a readable list of enabled features
            enabled_features = []
            if monitor_settings["pin_elements"]:
                pin_feature_text = "Pin elements >={}".format(pin_limit)
                if pinned_count > 0:
                    pin_feature_text += " ({} pinned)".format(pinned_count)
                enabled_features.append(pin_feature_text)
            if monitor_settings["warn_on_move"]:
                enabled_features.append("Warn when moving elements >={}".format(move_limit))
            if monitor_settings.get("warn_on_type_change", False):
                enabled_features.append(
                    "Warn on type changes with instances >= MMI {}".format(type_limit))
            if monitor_settings.get("warn_on_instance_params", False):
                enabled_features.append(
                    "Warn on instance parameter edits >= MMI {}".format(param_limit))
            if monitor_settings["validate_mmi"]:
                enabled_features.append("Attempt to fix MMI values")
            if monitor_settings["check_mmi_after_sync"]:
                enabled_features.append("Check MMI after sync")
            
            default_mmi_saved = get_default_mmi(revit.doc)
            if monitor_settings.get("default_on_new_instances", False) and default_mmi_saved:
                enabled_features.append(
                    "Default on new instances: {}".format(default_mmi_saved))
            elif monitor_settings.get("default_on_new_instances", False):
                enabled_features.append("Default on new instances (pick a value in Settings)")
                
            if not enabled_features:
                enabled_features.append("No features enabled (configure in Settings)")
            
            # Show the activation balloon with all enabled features
            forms.show_balloon(
                header="MMI Monitor", 
                text="Monitor activated \n\nActive features:\n• {}\n\nParameter: {}".format(
                    "\n• ".join(enabled_features),
                    mmi_param_name
                ),
                is_new=True
            )
            success = True
        else:
            forms.show_balloon(
                header="Error", 
                text="Failed to activate MMI Monitor",
                tooltip="Check logs for details",
                is_new=True
            )
    else:
        # Deactivate: Deregister handlers
        logger.debug("Deactivating MMI Monitor...")
        
        if deregister_event_handlers():
            set_monitor_active(False)
            script.toggle_icon(new_active_state)  # Toggle icon to inactive state
            
            # Clear caches
            element_location_cache = {}
            element_mmi_cache = {}
            element_mmi_blank = set()
            element_type_cache = {}
            type_warn_last = {}
            instance_param_warn_last = None
            baseline_element_ids_for_default = set()
            logger.debug("Cleared element location, MMI, type, and baseline caches")
            
            success = True
        else:
            forms.show_balloon(
                header="Error", 
                text="Failed to deactivate MMI Monitor",
                tooltip="Check logs for details",
                is_new=True
            )

    if success:
        logger.debug("MMI Monitor state toggled to: {}".format("ON" if new_active_state else "OFF"))
    else:
        logger.error("Failed to toggle MMI Monitor state.")

# --------------------------------------------------
# 💡 pyRevit with VSCode: Use pyrvt or pyrvtmin snippet
# 📄 Template has been developed by Baptiste LECHAT and inspired by Erik FRITS.