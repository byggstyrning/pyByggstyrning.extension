# -*- coding: utf-8 -*-
"""Toggles the MMI Monitor on and off.

When active, the monitor watches for relevant changes based on configuration.
When inactive, it does nothing.
"""

__title__ = "Monitor"
__author__ = "Byggstyrning AB"
__doc__ = "Toggle MMI Monitor on/off for the current session"
__highlight__ = 'new'

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
from mmi.config import CONFIG_SECTION, CONFIG_KEY_ACTIVE, MMI_THRESHOLD
from mmi.config import is_monitor_active, set_monitor_active
from mmi.core import get_mmi_parameter_name, load_monitor_config
from mmi.utils import get_element_location, get_element_mmi_value, validate_mmi_value

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
        self.mmi_threshold = MMI_THRESHOLD
        self.elements_to_validate = []
        self.validate_corrections = {}
        
    def Execute(self, uiapp):
        try:
            logger.debug("MMI Event Handler Executing")
            
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
                    
                    corrected_count = 0
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
                                
                        # Update the value if parameter exists
                        if param and param.HasValue and param.StorageType == StorageType.String:
                            param.Set(fixed_value)
                            corrected_count += 1
                            correction_details.append("'{}' → '{}'".format(orig_value, fixed_value))
                            logger.debug("Corrected MMI value from '{}' to '{}' for element {}".format(
                                orig_value, fixed_value, element_id))
                    
                    t.Commit()
                    
                    # Notify user of corrections
                    if corrected_count > 0:
                        forms.show_balloon(
                            header="MMI Value Correction",
                            text="{} MMI values automatically corrected".format(corrected_count),
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
                        tooltip="Elements with MMI value > {} were automatically pinned".format(self.mmi_threshold),
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
            self.notify_message = "Pinned {} elements with MMI value > {}"
        else:
            self.notify_message = None
            
    def validate_mmi_values_deferred(self, elements_to_validate, corrections):
        """Queue elements for MMI value validation and correction"""
        self.elements_to_validate = elements_to_validate
        self.validate_corrections = corrections

# Global handlers and events
mmi_event_handler = None
external_event = None
doc_changed_handler = None
doc_synchronizing_handler = None
doc_synchronized_handler = None
element_location_cache = {}  # Cache to store element locations for move detection
element_mmi_cache = {}  # Cache to store element MMI values to detect changes


def update_element_location_cache(element_id, location):
    """Update the element location cache."""
    global element_location_cache
    element_location_cache[element_id.IntegerValue] = {
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

def populate_initial_location_cache(doc):
    """Populate the location cache with all high MMI elements on monitor activation.
    This ensures we can detect movement even on the first move."""
    global element_location_cache
    try:
        mmi_param_name = get_mmi_parameter_name(doc)
        if not mmi_param_name:
            return
        
        logger.debug("Populating initial location cache for high MMI elements...")
        
        # Get all elements in the model
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        
        cache_count = 0
        for element in all_elements:
            # Skip elements that can't be pinned
            if not hasattr(element, "Pinned"):
                continue
            
            # Get the MMI value for the element
            mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
            
            # Only cache high MMI elements
            if mmi_value is not None and mmi_value > MMI_THRESHOLD:
                current_location = get_element_location(element)
                if current_location:
                    update_element_location_cache(element.Id, current_location)
                    cache_count += 1
        
        logger.debug("Cached locations for {} high MMI elements".format(cache_count))
        
    except Exception as ex:
        logger.error("Error populating initial location cache: {}".format(ex))

def populate_initial_mmi_cache(doc):
    """Populate the MMI cache with all elements on monitor activation.
    This allows us to detect MMI value changes."""
    global element_mmi_cache
    try:
        mmi_param_name = get_mmi_parameter_name(doc)
        if not mmi_param_name:
            return
        
        logger.debug("Populating initial MMI cache...")
        
        # Get all elements in the model
        all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        
        cache_count = 0
        for element in all_elements:
            # Skip elements that can't be pinned
            if not hasattr(element, "Pinned"):
                continue
            
            # Get the MMI value for the element
            mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
            
            # Cache all MMI values (both high and low)
            if mmi_value is not None:
                element_mmi_cache[element.Id.IntegerValue] = mmi_value
                cache_count += 1
        
        logger.debug("Cached MMI values for {} elements".format(cache_count))
        
    except Exception as ex:
        logger.error("Error populating initial MMI cache: {}".format(ex))

def pin_all_high_mmi_elements(doc):
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
            if mmi_value is not None and mmi_value >= MMI_THRESHOLD:
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

def document_changed_handler(sender, args):
    """Handler for document changed event."""
    try:
        # Check if we should monitor
        if not is_monitor_active():
            return
            
        # Get the modified elements
        modified_element_ids = args.GetModifiedElementIds()
        
        if modified_element_ids.Count == 0:
            return
            
        # Get the document
        doc = args.GetDocument()
        
        # Get MMI parameter name
        mmi_param_name = get_mmi_parameter_name(doc)
        if not mmi_param_name:
            logger.warning("No MMI parameter name configured. Use MMI Config tool first.")
            return
         
        # Get monitor settings
        monitor_settings = load_monitor_config(doc, use_display_names=False)
        
        # Check which features are enabled
        validate_enabled = monitor_settings["validate_mmi"]
        warn_on_move_enabled = monitor_settings["warn_on_move"]
        pin_elements_enabled = monitor_settings["pin_elements"]
        
        if not (validate_enabled or warn_on_move_enabled or pin_elements_enabled):
            logger.debug("No MMI monitor features are enabled. Skipping processing.")
            return
        
        logger.debug("Processing {} modified elements with MMI parameter: {}".format(
            modified_element_ids.Count, mmi_param_name))
        
        # Track elements for different operations
        elements_to_validate = []
        validation_corrections = {}
        elements_to_pin = []
        moved_high_mmi_elements = []
        
        # Clean old entries from element location cache
        clean_element_location_cache()
        
        # First pass: Process all elements for validation and move detection
        for element_id in modified_element_ids:
            element = doc.GetElement(element_id)
            
            # Skip null elements or elements that can't be pinned
            if element is None or not hasattr(element, "Pinned"):
                continue
                
            # Get the MMI value for the element
            mmi_value, value_str, param = get_element_mmi_value(element, mmi_param_name, doc)
            
            if mmi_value is not None:
                # ===== STEP 1: VALIDATE =====
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
                
                # ===== STEP 2: WARN ON MOVE =====
                if warn_on_move_enabled and mmi_value > MMI_THRESHOLD:
                    # Get current location
                    current_location = get_element_location(element)
                    if current_location:
                        # Check if we have a previous location
                        if element_id.IntegerValue in element_location_cache:
                            prev_location = element_location_cache[element_id.IntegerValue]["location"]
                            # Calculate distance moved
                            distance = current_location.DistanceTo(prev_location)
                            # If moved more than a small threshold (0.1 meters)
                            if distance > 0.1:
                                moved_high_mmi_elements.append({
                                    "id": element_id,
                                    "mmi": mmi_value,
                                    "distance": distance
                                })
                                logger.debug("High MMI Element {} moved {:.2f} meters".format(
                                    element_id, distance))
                        
                        # Update location cache for future checks
                        update_element_location_cache(element_id, current_location)
                
                # ===== STEP 3: PIN ELEMENTS (only when MMI changes) =====
                if pin_elements_enabled and mmi_value >= MMI_THRESHOLD:
                    # Check if MMI value changed to/above threshold
                    element_id_int = element_id.IntegerValue
                    prev_mmi = element_mmi_cache.get(element_id_int)
                    
                    # Pin if: 
                    # 1. Element is not already pinned, AND
                    # 2. Either we don't have a cached MMI (new element) OR the MMI value changed
                    should_pin = False
                    
                    if not element.Pinned:
                        if prev_mmi is None:
                            # First time seeing this element with high MMI
                            should_pin = True
                            logger.debug("Element {} newly detected with MMI {} - queuing for pin".format(
                                element_id, mmi_value))
                        elif prev_mmi < MMI_THRESHOLD and mmi_value >= MMI_THRESHOLD:
                            # MMI value changed from below to above threshold
                            should_pin = True
                            logger.debug("Element {} MMI changed from {} to {} - queuing for pin".format(
                                element_id, prev_mmi, mmi_value))
                    
                    if should_pin:
                        elements_to_pin.append(element_id)
                    
                    # Update MMI cache
                    element_mmi_cache[element_id_int] = mmi_value
                elif mmi_value is not None:
                    # Update cache for low MMI elements too (to detect future changes)
                    element_mmi_cache[element_id.IntegerValue] = mmi_value
        
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
                    item["id"].IntegerValue, item["mmi"], item["distance"]))
            
            tooltip = "High MMI elements should be carefully managed:\n" + "\n".join(details)
            if count > 5:
                tooltip += "\n... and {} more".format(count - 5)
            
            forms.show_balloon(
                header="High MMI Element Move",
                text="{} elements with MMI >= 425 were moved".format(count),
                tooltip=tooltip,
                is_new=True
            )
            logger.debug("Warned about {} moved high MMI elements".format(count))
        
        # Process pinning if needed
        if elements_to_pin and mmi_event_handler and external_event:
            # If already doing validation, avoid raising another external event immediately
            # The pin operation will be scheduled after validation completes
            if not elements_to_validate:
                mmi_event_handler.pin_elements_deferred(elements_to_pin, MMI_THRESHOLD)
                external_event.Raise()
                logger.debug("Queued {} elements for pinning using external event".format(len(elements_to_pin)))
            else:
                # Store the pinning request - it will be processed after validation
                mmi_event_handler.pin_elements_deferred(elements_to_pin, MMI_THRESHOLD)
                logger.debug("Pinning of {} elements will occur after validation".format(len(elements_to_pin)))
    
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
        # Create external event handler (for manual operations)
        if mmi_event_handler is None:
            mmi_event_handler = MMIEventHandler()
            external_event = ExternalEvent.Create(mmi_event_handler)
            logger.debug("MMI Event Handler Created.")

        # Register for document changed events
        if doc_changed_handler is None:
            doc_changed_handler = EventHandler[DocumentChangedEventArgs](document_changed_handler)
            # Note: Application is better than Document for app-level monitoring
            revit.doc.Application.DocumentChanged += doc_changed_handler
            logger.debug("Document Changed Handler registered.")
        
        # Register for document synchronizing events
        if doc_synchronizing_handler is None:
            doc_synchronizing_handler = EventHandler[DocumentSynchronizingWithCentralEventArgs](document_synchronizing_handler)
            revit.doc.Application.DocumentSynchronizingWithCentral += doc_synchronizing_handler
            logger.debug("Document Synchronizing Handler registered.")
            
        # Register for document synchronized events  
        if doc_synchronized_handler is None:
            doc_synchronized_handler = EventHandler[DocumentSynchronizedWithCentralEventArgs](document_synchronized_handler)
            revit.doc.Application.DocumentSynchronizedWithCentral += doc_synchronized_handler
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
        # Unregister document changed event handler
        if doc_changed_handler is not None:
            revit.doc.Application.DocumentChanged -= doc_changed_handler
            doc_changed_handler = None
            logger.debug("Document Changed Handler unregistered.")
        
        # Unregister document synchronizing event handler
        if doc_synchronizing_handler is not None:
            revit.doc.Application.DocumentSynchronizingWithCentral -= doc_synchronizing_handler
            doc_synchronizing_handler = None
            logger.debug("Document Synchronizing Handler unregistered.")
            
        # Unregister document synchronized event handler
        if doc_synchronized_handler is not None:
            revit.doc.Application.DocumentSynchronizedWithCentral -= doc_synchronized_handler
            doc_synchronized_handler = None
            logger.debug("Document Synchronized Handler unregistered.")
        
        # Clear any pending operations and mark handlers as inactive
        if mmi_event_handler is not None:
            # Clear all queued operations
            mmi_event_handler.elements_to_pin = []
            mmi_event_handler.elements_to_validate = []
            mmi_event_handler.validate_corrections = {}
            mmi_event_handler.notify_message = None
            logger.debug("Cleared all queued MMI operations")
        
        # Dispose the external event if created
        if external_event is not None:
            external_event = None # Mark as inactive
            mmi_event_handler = None
            logger.debug("MMI ExternalEvent marked as inactive.")

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
            
            # Populate initial location cache if warn_on_move is enabled
            if monitor_settings["warn_on_move"]:
                populate_initial_location_cache(revit.doc)
            
            # Populate initial MMI cache if pin_elements is enabled (to detect changes)
            if monitor_settings["pin_elements"]:
                populate_initial_mmi_cache(revit.doc)
            
            # Proactively pin all high MMI elements if pin_elements is enabled
            pinned_count = 0
            if monitor_settings["pin_elements"]:
                pinned_count = pin_all_high_mmi_elements(revit.doc)
            
            # Create a readable list of enabled features
            enabled_features = []
            if monitor_settings["pin_elements"]:
                pin_feature_text = "Pin elements >={}".format(MMI_THRESHOLD)
                if pinned_count > 0:
                    pin_feature_text += " ({} pinned)".format(pinned_count)
                enabled_features.append(pin_feature_text)
            if monitor_settings["warn_on_move"]:
                enabled_features.append("Warn when moving elements >{}".format(MMI_THRESHOLD))
            if monitor_settings["validate_mmi"]:
                enabled_features.append("Validate MMI format")
            if monitor_settings["check_mmi_after_sync"]:
                enabled_features.append("Check MMI after sync")
                
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
            logger.debug("Cleared element location and MMI caches")
            
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