# -*- coding: utf-8 -*-
"""Monitor event handlers for 3D Zone parameter mapping."""

import datetime
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.Events import DocumentChangedEventArgs
from Autodesk.Revit.UI import IExternalEventHandler
from pyrevit import revit, script, forms
try:
    from zone3d import config, core, containment
except ImportError:
    # Handle relative imports
    import config
    import core
    import containment

# Initialize logger
logger = script.get_logger()

# Global state
_monitor_active = False
_element_location_cache = {}

def is_monitor_active():
    """Check if monitor is currently active."""
    return _monitor_active

def set_monitor_active(active):
    """Set monitor active state."""
    global _monitor_active
    _monitor_active = active
    if not active:
        # Clear cache when deactivating
        clear_location_cache()

def update_element_location_cache(element_id, location):
    """Update the element location cache."""
    global _element_location_cache
    _element_location_cache[element_id.IntegerValue] = {
        "location": location,
        "timestamp": datetime.datetime.now()
    }

def clear_location_cache():
    """Clear the element location cache."""
    global _element_location_cache
    _element_location_cache = {}

def clean_location_cache():
    """Clean old entries from the location cache."""
    global _element_location_cache
    now = datetime.datetime.now()
    # Keep entries not older than 5 minutes
    _element_location_cache = {
        el_id: data for el_id, data in _element_location_cache.items()
        if (now - data["timestamp"]).total_seconds() < 300
    }

class Zone3DEventHandler(IExternalEventHandler):
    """Event handler for processing zone parameter updates."""
    
    def __init__(self):
        self.elements_to_update = []
        self.configs_to_process = []
        
    def Execute(self, uiapp):
        """Execute parameter updates for moved elements."""
        try:
            logger.info("Zone3D Event Handler Executing")
            
            # Check if monitor is still active
            if not is_monitor_active():
                logger.info("Monitor is not active, clearing queued operations")
                self.elements_to_update = []
                self.configs_to_process = []
                return
            
            doc = uiapp.ActiveUIDocument.Document
            
            if not self.elements_to_update or not self.configs_to_process:
                return
            
            # Process each configuration in order
            for zone_config in self.configs_to_process:
                try:
                    config_name = zone_config.get("name", "Unknown")
                    logger.info("Processing configuration: {}".format(config_name))
                    
                    # Process moved elements for this configuration
                    updated_count = 0
                    
                    with Transaction(doc, "3D Zone Monitor: {}".format(config_name)) as t:
                        t.Start()
                        
                        for element_id in self.elements_to_update:
                            try:
                                element = doc.GetElement(element_id)
                                if not element:
                                    continue
                                
                                # Detect containment strategy
                                source_categories = zone_config.get("source_categories", [])
                                strategy = containment.detect_containment_strategy(source_categories)
                                
                                if not strategy:
                                    continue
                                
                                # Find containing element
                                containing_el = containment.get_containing_element_by_strategy(
                                    element, doc, strategy, source_categories
                                )
                                
                                if not containing_el:
                                    continue
                                
                                # Copy parameters
                                source_params = zone_config.get("source_params", [])
                                target_params = zone_config.get("target_params", [])
                                
                                if len(source_params) != len(target_params):
                                    continue
                                
                                params_copied = core.copy_parameters(
                                    containing_el, element,
                                    source_params, target_params
                                )
                                
                                if params_copied > 0:
                                    updated_count += 1
                                    logger.info("Updated element {} with {} parameters".format(
                                        element_id, params_copied))
                            
                            except Exception as e:
                                logger.info("Error processing element {}: {}".format(element_id, str(e)))
                                continue
                        
                        t.Commit()
                    
                    if updated_count > 0:
                        logger.info("Configuration '{}': Updated {} elements".format(
                            config_name, updated_count))
                
                except Exception as e:
                    logger.error("Error processing configuration: {}".format(str(e)))
                    continue
            
            # Clear the queue
            self.elements_to_update = []
            self.configs_to_process = []
            
        except Exception as ex:
            logger.error("Error in Zone3D Event Handler: {}".format(ex))
    
    def GetName(self):
        return "3D Zone Monitor Event Handler"
    
    def queue_elements_for_update(self, element_ids, configs):
        """Queue elements for parameter update in deferred execution."""
        self.elements_to_update = element_ids
        self.configs_to_process = configs

# Global handler instance
_zone3d_event_handler = None
_external_event = None
_doc_changed_handler = None

def get_event_handler():
    """Get or create the event handler instance."""
    global _zone3d_event_handler
    if _zone3d_event_handler is None:
        _zone3d_event_handler = Zone3DEventHandler()
    return _zone3d_event_handler

def get_external_event():
    """Get or create the external event instance."""
    global _external_event
    if _external_event is None:
        from Autodesk.Revit.UI import ExternalEvent
        _external_event = ExternalEvent.Create(get_event_handler())
    return _external_event

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
        
        # Load enabled configurations
        configs = config.get_enabled_configs(doc)
        if not configs:
            return
        
        logger.info("Processing {} modified elements with {} configurations".format(
            modified_element_ids.Count, len(configs)))
        
        # Clean old entries from cache
        clean_location_cache()
        
        # Track moved elements
        moved_elements = []
        
        # Check each modified element for movement
        for element_id in modified_element_ids:
            element = doc.GetElement(element_id)
            if not element:
                continue
            
            # Get current location
            current_location = containment.get_element_representative_point(element)
            if not current_location:
                continue
            
            # Check if we have a previous location
            element_id_int = element_id.IntegerValue
            if element_id_int in _element_location_cache:
                prev_location = _element_location_cache[element_id_int]["location"]
                # Calculate distance moved
                distance = current_location.DistanceTo(prev_location)
                # If moved more than a small threshold (0.1 meters)
                if distance > 0.1:
                    moved_elements.append(element_id)
                    logger.info("Element {} moved {:.2f} meters".format(element_id, distance))
            
            # Update location cache for future checks
            update_element_location_cache(element_id, current_location)
        
        # If elements moved, queue for update
        if moved_elements:
            logger.info("Queuing {} moved elements for parameter update".format(len(moved_elements)))
            handler = get_event_handler()
            external_evt = get_external_event()
            
            handler.queue_elements_for_update(moved_elements, configs)
            external_evt.Raise()
    
    except Exception as ex:
        logger.error("Error in document changed handler: {}".format(ex))

def register_event_handlers(doc):
    """Register event handlers for monitoring."""
    global _doc_changed_handler
    
    try:
        from System import EventHandler
        
        # Register for document changed events
        if _doc_changed_handler is None:
            _doc_changed_handler = EventHandler[DocumentChangedEventArgs](document_changed_handler)
            doc.Application.DocumentChanged += _doc_changed_handler
            logger.info("Document Changed Handler registered")
        
        return True
    except Exception as e:
        logger.error("Failed to register event handlers: {}".format(e))
        return False

def deregister_event_handlers(doc):
    """Deregister event handlers."""
    global _doc_changed_handler
    
    try:
        if _doc_changed_handler is not None:
            doc.Application.DocumentChanged -= _doc_changed_handler
            _doc_changed_handler = None
            logger.info("Document Changed Handler unregistered")
        
        return True
    except Exception as e:
        logger.error("Failed to deregister event handlers: {}".format(e))
        return False

def populate_initial_location_cache(doc, configs):
    """Populate the location cache with all elements that might be affected.
    
    Args:
        doc: Revit document
        configs: List of configurations
    """
    try:
        logger.info("Populating initial location cache...")
        
        # Collect all target elements from configurations
        target_categories = set()
        for cfg in configs:
            target_filter = cfg.get("target_filter_categories", [])
            target_categories.update(target_filter)
        
        # Collect target elements
        from Autodesk.Revit.DB import FilteredElementCollector
        collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
        
        if target_categories:
            for category in target_categories:
                collector = collector.OfCategory(category)
        
        all_elements = collector.ToElements()
        
        cache_count = 0
        for element in all_elements:
            location = containment.get_element_representative_point(element)
            if location:
                update_element_location_cache(element.Id, location)
                cache_count += 1
        
        logger.info("Cached locations for {} elements".format(cache_count))
    
    except Exception as ex:
        logger.error("Error populating initial location cache: {}".format(ex))

