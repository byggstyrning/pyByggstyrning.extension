# -*- coding: utf-8 -*-
"""IFC export event handlers for NestedParentId parameter mapping.

This module implements automatic parameter writing during IFC export:

1. FileExporting Event: Writes NestedParentId parameters via SubTransaction
   (temporary write - helps with export but may not persist)

2. FileExported Event: Uses ExternalEvent to write parameters persistently
   after the export transaction completes, then saves the document.
"""

from pyrevit import script, HOST_APP

# Import ExternalEvent classes
try:
    from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
except ImportError:
    IExternalEventHandler = None
    ExternalEvent = None

# Initialize logger
try:
    logger = script.get_logger()
except Exception:
    import logging
    logger = logging.getLogger('nested_family.ifc_export')

# Global handler references
_ifc_export_handler = None
_ifc_exported_handler = None

# ExternalEvent for post-export parameter writing
_post_export_write_handler = None
_post_export_write_event = None

# Cache for parameter writes
_parameter_write_cache = {}  # Format: {element_id: parent_id}

# Configuration
PARAM_NAME = "NestedParentId"


def get_revit_version():
    """Get Revit version as integer."""
    return int(HOST_APP.version)


def get_elementid_value(element_id):
    """Get ElementId value based on Revit version (API changed in 2024)."""
    version = get_revit_version()
    if version > 2023:
        return element_id.Value
    else:
        return element_id.IntegerValue


def get_nested_family_instances(doc):
    """Get all nested shared family instances in the document.
    
    Args:
        doc: Revit document
        
    Returns:
        list: List of tuples (nested_instance, parent_instance)
    """
    from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance
    
    nested_pairs = []
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance)
    
    for fi in collector:
        try:
            super_component = fi.SuperComponent
            if super_component and isinstance(super_component, FamilyInstance):
                nested_pairs.append((fi, super_component))
        except Exception:
            continue
    
    return nested_pairs


def check_parameter_exists(doc, param_name=PARAM_NAME):
    """Check if the parameter is bound in the document."""
    bindings = doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    while iterator.MoveNext():
        if iterator.Key.Name == param_name:
            return True
    return False


def set_nested_parent_ids(doc, dry_run=False, param_name=PARAM_NAME, use_cache=False, only_empty=True):
    """Set NestedParentId on all nested shared family instances.
    
    Args:
        doc: Revit document
        dry_run: If True, only report what would be done
        param_name: Name of the parameter to set
        use_cache: If True, populate cache for later use
        only_empty: If True, only set values on elements where parameter is empty
        
    Returns:
        dict: Results with counts and details
    """
    global _parameter_write_cache
    
    results = {
        "total_nested": 0,
        "successfully_set": 0,
        "already_set": 0,
        "param_not_found": 0,
        "param_readonly": 0,
        "errors": 0,
    }
    
    nested_pairs = get_nested_family_instances(doc)
    results["total_nested"] = len(nested_pairs)
    
    if not nested_pairs:
        return results
    
    if use_cache:
        _parameter_write_cache = {}
    
    for nested_instance, parent_instance in nested_pairs:
        try:
            nested_id = get_elementid_value(nested_instance.Id)
            parent_id = get_elementid_value(parent_instance.Id)
            parent_id_str = str(parent_id)
            
            param = nested_instance.LookupParameter(param_name)
            
            if not param:
                results["param_not_found"] += 1
                continue
            
            if param.IsReadOnly:
                results["param_readonly"] += 1
                continue
            
            # Check current value (stored as string)
            current_value = param.AsString() if param.HasValue else ""
            
            # Skip if already has a value (when only_empty=True)
            if only_empty and current_value:
                results["already_set"] += 1
                # Still cache the current value
                if use_cache:
                    _parameter_write_cache[nested_id] = current_value
                continue
            
            # Skip if already correct
            if current_value == parent_id_str:
                results["already_set"] += 1
                if use_cache:
                    _parameter_write_cache[nested_id] = parent_id_str
                continue
            
            # Cache the value for later use
            if use_cache:
                _parameter_write_cache[nested_id] = parent_id_str
            
            # Set the value (unless dry run)
            if not dry_run:
                try:
                    param.Set(parent_id_str)
                    results["successfully_set"] += 1
                except Exception:
                    results["errors"] += 1
            else:
                results["successfully_set"] += 1
                
        except Exception:
            results["errors"] += 1
    
    return results


def write_cached_parameters(doc, param_name=PARAM_NAME, only_empty=True):
    """Write cached parameter values to elements.
    
    Args:
        doc: Revit document
        param_name: Name of the parameter to set
        only_empty: If True, only set values on elements where parameter is empty
        
    Returns:
        dict: Results with counts
    """
    global _parameter_write_cache
    
    from Autodesk.Revit.DB import ElementId
    
    results = {
        "successfully_set": 0,
        "already_set": 0,
        "param_not_found": 0,
        "errors": 0,
    }
    
    if not _parameter_write_cache:
        return results
    
    for nested_id, parent_id_str in _parameter_write_cache.items():
        try:
            element = doc.GetElement(ElementId(nested_id))
            if not element:
                results["errors"] += 1
                continue
            
            param = element.LookupParameter(param_name)
            if not param:
                results["param_not_found"] += 1
                continue
            
            if param.IsReadOnly:
                results["errors"] += 1
                continue
            
            current_value = param.AsString() if param.HasValue else ""
            
            # Skip if already has a value (when only_empty=True)
            if only_empty and current_value:
                results["already_set"] += 1
                continue
            
            if current_value == parent_id_str:
                results["already_set"] += 1
                continue
            
            param.Set(parent_id_str)
            results["successfully_set"] += 1
            
        except Exception:
            results["errors"] += 1
    
    return results


class PostExportWriteHandler(IExternalEventHandler):
    """Handler for ExternalEvent to write parameters after IFC export transaction completes."""
    
    def __init__(self):
        self.doc = None
        self.should_save = True
    
    def Execute(self, uiapp):
        """Execute parameter writes in a normal transaction after export completes."""
        if not self.doc:
            return
        
        try:
            # Check if parameter exists
            if not check_parameter_exists(self.doc, PARAM_NAME):
                return
            
            # Write parameters in a transaction
            from Autodesk.Revit.DB import Transaction
            
            transaction = Transaction(self.doc, "Write NestedParentId (Post-IFC Export)")
            try:
                transaction.Start()
                
                # Use cached values if available, otherwise recalculate
                # Only update empty parameters (once set, they don't change)
                global _parameter_write_cache
                if _parameter_write_cache:
                    results = write_cached_parameters(self.doc, PARAM_NAME, only_empty=True)
                else:
                    results = set_nested_parent_ids(self.doc, dry_run=False, param_name=PARAM_NAME, only_empty=True)
                
                transaction.Commit()
                
                # Only log if we actually set something
                if results.get("successfully_set", 0) > 0:
                    logger.debug("NestedParentId: Set {} new values".format(results["successfully_set"]))
                
            except Exception as tx_err:
                if transaction.HasStarted():
                    transaction.RollBack()
                logger.error("NestedParentId transaction error: {}".format(str(tx_err)))
                return
            finally:
                transaction.Dispose()
            
            # Save the document if changes were made
            if self.should_save and results.get("successfully_set", 0) > 0:
                try:
                    self.doc.Save()
                    logger.debug("NestedParentId: Document saved")
                except Exception as save_err:
                    logger.warning("Could not save document: {}".format(str(save_err)))
            
        except Exception as ex:
            logger.error("Error in NestedParentId post-export handler: {}".format(str(ex)))
        finally:
            self.doc = None
            _parameter_write_cache.clear()
    
    def GetName(self):
        return "NestedParentId Post-Export Write"


def file_exporting_handler(sender, args):
    """Handler for file exporting event - writes NestedParentId before IFC export.
    
    Args:
        sender: Application object
        args: FileExportingEventArgs
    """
    global _parameter_write_cache
    
    try:
        doc = args.Document
        if not doc:
            return
        
        # Check if this is an IFC export
        export_format = getattr(args, 'Format', None)
        if export_format:
            format_str = str(export_format).upper()
            if "IFC" not in format_str:
                return
        
        # Check if parameter exists
        if not check_parameter_exists(doc, PARAM_NAME):
            return
        
        # Try to write using SubTransaction (may or may not persist)
        # Only update empty parameters (once set, they don't change)
        try:
            from Autodesk.Revit.DB import SubTransaction
            
            sub_trans = SubTransaction(doc)
            sub_trans.Start()
            
            try:
                # Write parameters and cache values for post-export
                results = set_nested_parent_ids(doc, dry_run=False, param_name=PARAM_NAME, use_cache=True, only_empty=True)
                sub_trans.Commit()
                
                # Only log if we set new values
                if results.get("successfully_set", 0) > 0:
                    logger.debug("NestedParentId pre-export: Set {} new values".format(results["successfully_set"]))
                
            except Exception as sub_err:
                if sub_trans.HasStarted():
                    sub_trans.RollBack()
                # Still cache values for post-export write
                _parameter_write_cache = {}
                nested_pairs = get_nested_family_instances(doc)
                for nested, parent in nested_pairs:
                    nested_id = get_elementid_value(nested.Id)
                    parent_id = get_elementid_value(parent.Id)
                    _parameter_write_cache[nested_id] = str(parent_id)
            finally:
                sub_trans.Dispose()
                
        except Exception:
            # Cache values anyway for post-export
            _parameter_write_cache = {}
            nested_pairs = get_nested_family_instances(doc)
            for nested, parent in nested_pairs:
                nested_id = get_elementid_value(nested.Id)
                parent_id = get_elementid_value(parent.Id)
                _parameter_write_cache[nested_id] = str(parent_id)
        
    except Exception as ex:
        logger.error("Error in NestedParentId IFC export handler: {}".format(str(ex)))


def file_exported_handler(sender, args):
    """Handler for FileExported event - writes parameters after IFC export completes.
    
    Args:
        sender: Application object
        args: FileExportedEventArgs
    """
    try:
        doc = args.Document
        if not doc:
            return
        
        # Check if this is an IFC export
        export_format = getattr(args, 'Format', None)
        if not export_format or "IFC" not in str(export_format).upper():
            return
        
        # Check if parameter exists
        if not check_parameter_exists(doc, PARAM_NAME):
            return
        
        # Check if we have cached values to write
        global _parameter_write_cache
        if not _parameter_write_cache:
            return
        
        # Initialize ExternalEvent if not already created
        global _post_export_write_handler, _post_export_write_event
        if _post_export_write_handler is None:
            if IExternalEventHandler is None or ExternalEvent is None:
                return
            _post_export_write_handler = PostExportWriteHandler()
            _post_export_write_event = ExternalEvent.Create(_post_export_write_handler)
        
        # Store doc reference and trigger ExternalEvent
        _post_export_write_handler.doc = doc
        _post_export_write_handler.should_save = True
        
        try:
            _post_export_write_event.Raise()
        except Exception as event_err:
            logger.error("Failed to raise NestedParentId ExternalEvent: {}".format(str(event_err)))
        
    except Exception as ex:
        logger.error("Error in NestedParentId FileExported handler: {}".format(str(ex)))


def register_ifc_export_handler():
    """Register the IFC export event handler at application level.
    
    Returns:
        bool: True if successful, False otherwise
    """
    global _ifc_export_handler, _ifc_exported_handler
    
    try:
        from System import EventHandler
        from Autodesk.Revit.DB.Events import FileExportingEventArgs, FileExportedEventArgs
        
        app = HOST_APP.app
        if not app:
            logger.error("Could not get Revit application")
            return False
        
        # Register FileExporting handler
        if _ifc_export_handler is None:
            _ifc_export_handler = EventHandler[FileExportingEventArgs](file_exporting_handler)
            app.FileExporting += _ifc_export_handler
            logger.debug("NestedParentId IFC export handler registered")
        
        # Register FileExported handler
        if _ifc_exported_handler is None:
            _ifc_exported_handler = EventHandler[FileExportedEventArgs](file_exported_handler)
            app.FileExported += _ifc_exported_handler
            logger.debug("NestedParentId IFC exported handler registered")
        
        return True
        
    except Exception as e:
        logger.error("Failed to register NestedParentId IFC export handler: {}".format(str(e)))
        return False


def deregister_ifc_export_handler():
    """Deregister the IFC export event handler.
    
    Returns:
        bool: True if successful, False otherwise
    """
    global _ifc_export_handler, _ifc_exported_handler
    
    try:
        app = HOST_APP.app
        if not app:
            return False
        
        if _ifc_export_handler is not None:
            app.FileExporting -= _ifc_export_handler
            _ifc_export_handler = None
            logger.debug("NestedParentId IFC export handler unregistered")
        
        if _ifc_exported_handler is not None:
            app.FileExported -= _ifc_exported_handler
            _ifc_exported_handler = None
            logger.debug("NestedParentId IFC exported handler unregistered")
        
        return True
        
    except Exception as e:
        logger.error("Failed to deregister NestedParentId IFC export handler: {}".format(str(e)))
        return False
