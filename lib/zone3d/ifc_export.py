# -*- coding: utf-8 -*-
"""IFC export event handlers for 3D Zone parameter mapping.

This module implements automatic parameter writing during IFC export using a caching pattern:

1. FileExporting Event: Runs containment calculations and writes parameters via SubTransaction
   (temporary write - helps with export but doesn't persist due to Revit's transaction rollback)

2. FileExported Event: Uses cached parameter values from FileExporting and writes them
   persistently via ExternalEvent (executes after export transaction completes)

The caching mechanism avoids duplicate containment calculations - values are calculated once
in FileExporting and reused in FileExported for direct parameter writes.
"""

from pyrevit import script, revit
try:
    from zone3d import config, core
except ImportError:
    # Handle relative imports
    import config
    import core

# Import ExternalEvent classes
try:
    from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
except ImportError:
    IExternalEventHandler = None
    ExternalEvent = None

# Initialize logger (with error handling for import-time issues)
try:
    logger = script.get_logger()
except Exception:
    # Fallback logger if script.get_logger() fails during import
    import logging
    logger = logging.getLogger('zone3d.ifc_export')

# Global handler references
_ifc_export_handler = None
_ifc_exported_handler = None

# Store export info for FileExported verification
_pending_export_settings = None

# ExternalEvent for post-export parameter writing
_post_export_write_handler = None
_post_export_write_event = None

# Cache for parameter writes (element_id -> {param_name: param_value})
# This cache is populated in FileExporting and used in FileExported to avoid recalculation
_parameter_write_cache = {}  # Format: {config_name: {element_id: {param_name: param_value}}}

class PostExportWriteHandler(IExternalEventHandler):
    """Handler for ExternalEvent to write parameters after IFC export transaction completes."""
    
    def __init__(self):
        self.configs_to_execute = []
        self.doc = None
    
    def Execute(self, uiapp):
        """Execute parameter writes in a normal transaction after export completes."""
        global _pending_export_settings

        if not self.doc or not self.configs_to_execute:
            return
        
        try:
            summary = {
                "total_configs": len(self.configs_to_execute),
                "config_results": [],
                "total_elements_updated": 0,
                "total_elements_already_correct": 0,
                "total_parameters_copied": 0,
                "total_parameters_already_correct": 0
            }
            
            for config_idx, zone_config in enumerate(self.configs_to_execute):
                config_name = zone_config.get("name", "Unknown")

                try:
                    # Use cached parameter values from FileExporting (no recalculation)
                    global _parameter_write_cache
                    config_cache = _parameter_write_cache.get(config_name, {})
                    
                    if config_cache:
                        # Write cached values directly (no containment calculation)
                        target_params = zone_config.get("target_params", [])
                        only_empty = zone_config.get("ifc_export_only_empty", False)

                        # Write cached values in a normal transaction
                        from Autodesk.Revit.DB import Transaction
                        transaction = Transaction(self.doc, "3D Zone: {} (Cached)".format(config_name))
                        try:
                            transaction.Start()
                            result = core.write_cached_parameters(self.doc, config_cache, target_params, only_empty=only_empty)
                            transaction.Commit()
                            result["config_name"] = config_name
                            result["config_order"] = config_idx
                        except Exception as tx_err:
                            transaction.RollBack()
                            raise
                        finally:
                            transaction.Dispose()
                    else:
                        # Fallback: full recalculation if cache not available

                        result = core.execute_configuration(self.doc, zone_config, progress_bar=None, view_id=None, force_transaction=True, use_subtransaction=False)
                        result["config_name"] = config_name
                        result["config_order"] = config_idx
                    
                    summary["config_results"].append(result)
                    summary["total_elements_updated"] += result.get("elements_updated", 0)
                    summary["total_elements_already_correct"] += result.get("elements_already_correct", 0)
                    summary["total_parameters_copied"] += result.get("parameters_copied", 0)
                    summary["total_parameters_already_correct"] += result.get("parameters_already_correct", 0)
                    
                except Exception as e:
                    error_msg = str(e)
                    logger.error("Error executing configuration '{}': {}".format(config_name, error_msg))
                    summary["config_results"].append({
                        "config_name": config_name,
                        "config_order": config_idx,
                        "elements_updated": 0,
                        "parameters_copied": 0,
                        "errors": [str(e)]
                    })

        except Exception as ex:

            logger.error("Error in post-export write handler: {}".format(str(ex)))
        finally:
            # Clear configs and doc reference
            self.configs_to_execute = []
            self.doc = None
    
    def GetName(self):
        return "Post-Export Parameter Write"

class ExportReexecuteHandler:
    """Handler for ExternalEvent to re-trigger IFC export after parameter write."""
    
    def Execute(self, app):
        """Execute the re-triggered export."""
        global _pending_export_settings
        
        try:
            import json
            import os.path as op
            log_path = op.join(op.dirname(op.dirname(op.dirname(op.abspath(__file__)))), ".cursor", "debug.log")
            with open(log_path, "a") as f:
                f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"L","location":"ifc_export.py:ExportReexecuteHandler.Execute","message":"Re-executing export","data":{"has_settings":_pending_export_settings is not None},"timestamp":__import__("time").time()}) + "\n")
        except: pass
        
        if not _pending_export_settings:
            logger.warning("No pending export settings - cannot re-trigger export")
            return
        
        try:
            doc = _pending_export_settings.get("doc")
            export_path = _pending_export_settings.get("export_path")
            view_id = _pending_export_settings.get("view_id")
            stored_ifc_options = _pending_export_settings.get("ifc_options")
            
            if not doc or not export_path:
                logger.warning("Missing document or export path - cannot re-trigger export")
                return
            
            # Import IFC export classes
            from Autodesk.Revit.DB import ElementId
            from Autodesk.Revit.DB.IFC import IFCExportOptions
            
            # Parse export path to get folder and filename
            import os.path as op
            folder_path = op.dirname(export_path)
            file_name = op.basename(export_path)
            
            # Use stored IFC export options if available, otherwise create new ones
            if stored_ifc_options:
                # Try to clone or reuse the stored options
                # Note: IFCExportOptions might not be directly cloneable, so we'll create new and copy key settings
                ifc_options = IFCExportOptions()
                # Copy key properties from stored options if possible
                try:
                    if hasattr(stored_ifc_options, "ViewId") and stored_ifc_options.ViewId:
                        ifc_options.ViewId = stored_ifc_options.ViewId
                except:
                    pass
            else:
                # Create new IFC export options
                ifc_options = IFCExportOptions()
                
                # Set view if available
                if view_id:
                    ifc_options.ViewId = view_id
            
            # Re-trigger the export

            doc.Export(folder_path, file_name, ifc_options)

        except Exception as export_err:
            logger.error("Error re-triggering IFC export: {}".format(str(export_err)))
        finally:
            # Clear pending settings
            _pending_export_settings = None
    
    def GetName(self):
        return "IFC Export Re-execute"

def file_exporting_handler(sender, args):
    """Handler for file exporting event - writes zone parameters before IFC export.
    
    Args:
        sender: Application object
        args: FileExportingEventArgs
    """
    try:
        # Get the document being exported
        doc = args.Document
        if not doc:

            return

        # Check if this is an IFC export
        # The export format is available in the event args
        export_format = None
        try:
            # Check if export format is IFC
            # FileExportingEventArgs has a Format property
            export_format = getattr(args, 'Format', None)

            if export_format:
                # Check if format string contains "IFC" (case-insensitive)
                format_str = str(export_format).upper()
                if "IFC" not in format_str:

                    logger.debug("Export format is not IFC: {}".format(export_format))
                    return
        except Exception as e:

            logger.debug("Could not determine export format: {}".format(str(e)))
            # If we can't determine format, assume it might be IFC and proceed
            # This is safer than missing IFC exports
        
        logger.debug("IFC export detected, checking for configurations with write_before_ifc_export enabled")
        
        # Load all configurations
        all_configs = config.load_configs(doc)

        if not all_configs:

            logger.debug("No configurations found")
            return
        
        # Filter configurations that have write_before_ifc_export enabled
        ifc_configs = [
            cfg for cfg in all_configs 
            if cfg.get("write_before_ifc_export", False) and cfg.get("enabled", False)
        ]

        if not ifc_configs:

            logger.debug("No configurations with write_before_ifc_export enabled")
            return
        
        logger.debug("Found {} configuration(s) to execute before IFC export".format(len(ifc_configs)))
        
        # Get the view being exported (if available)
        # FileExportingEventArgs may have ViewId property or it may be in export context
        view_id = None

        try:
            # Try to get ViewId directly from event args
            view_id = getattr(args, 'ViewId', None)

            # If not found, try to get from export context or options
            if not view_id:
                try:
                    # Check if there's a Context property with ViewId
                    context = getattr(args, 'Context', None)

                    if context:
                        view_id = getattr(context, 'ViewId', None)

                except Exception as ctx_e:

                    pass
            
            # Try to inspect export options for view information
            if not view_id:
                try:
                    # Check for ExportOptions or similar
                    export_options = getattr(args, 'ExportOptions', None)

                    if export_options:
                        view_id = getattr(export_options, 'ViewId', None)

                except: pass

            if view_id:
                logger.debug("IFC export from view: {}".format(view_id.IntegerValue))
            else:
                logger.debug("No ViewId found in export event - processing all elements (not filtered by view)")
        except Exception as e:

            logger.debug("Could not get ViewId from export event: {}".format(str(e)))
        
        # CRITICAL: Capture export settings before canceling
        # We'll need these to re-trigger the export after writing parameters
        export_path = None
        export_view_id = view_id  # Already captured above
        
        # Try to get export path from args
        try:
            export_path = getattr(args, "Path", None)
            if export_path:
                # Path might be a full path string
                export_path = str(export_path)
        except:
            pass
        
        # Try to get IFCExportOptions if available (for more complete settings)
        ifc_export_options = None
        try:
            # Check if args has IFCExportOptions property
            ifc_export_options = getattr(args, "IFCExportOptions", None)
        except:
            pass

        # NOTE: FileExportingEventArgs.Cancel is read-only - we cannot cancel the export.
        # We'll attempt to write parameters during the export, but changes may not persist
        # because we're inside Revit's export pipeline. The command hook approach (writing
        # when dialog opens via ExternalEvent) is preferred, but this serves as a fallback.

        # Execute configurations
        # Create a summary results dict
        summary = {
            "total_configs": len(ifc_configs),
            "config_results": [],
            "total_elements_updated": 0,
            "total_elements_already_correct": 0,
            "total_parameters_copied": 0,
            "total_parameters_already_correct": 0
        }

        # Execute each configuration with caching enabled
        # Note: Progress bars don't display reliably in event handlers, so we use logger output instead
        global _parameter_write_cache
        
        for config_idx, zone_config in enumerate(ifc_configs):
            config_name = zone_config.get("name", "Unknown")
            config_order = zone_config.get("order", 0)

            # Create cache dict for this configuration
            config_cache = {}
            if config_name not in _parameter_write_cache:
                _parameter_write_cache[config_name] = {}
            
            try:
                # Execute configuration within a transaction WITH CACHING
                # Pass view_id to filter elements by view visibility
                # Note: progress_bar=None because progress bars don't work in event handlers
                # User can see progress via logger output in pyRevit output window
                # In FileExporting, we're inside Revit's internal "Export IFC" transaction.
                # Try SubTransaction to nest inside the parent transaction.
                # Cache parameter values for later use in FileExported (no recalculation needed)

                result = core.execute_configuration(doc, zone_config, progress_bar=None, view_id=view_id, force_transaction=True, use_subtransaction=True, cache_dict=config_cache)
                
                # Store cache in global cache dict
                if config_cache:
                    _parameter_write_cache[config_name] = config_cache

                result["config_name"] = config_name
                result["config_order"] = config_order

                # Check for errors in result and show formatted message
                errors = result.get("errors", [])
                if errors:
                    error_msg = errors[0] if errors else ""
                    # Check if this is a transaction error (SubTransaction also failed)
                    if "not permitted" in error_msg.lower() or "read-only" in error_msg.lower() or "cannot start" in error_msg.lower() or "subtransaction" in error_msg.lower():
                        logger.error("=" * 60)
                        logger.error("ERROR: Cannot write parameters during IFC export!")
                        logger.error("Both IdleAction and SubTransaction approaches failed.")
                        logger.error("")
                        logger.error("SOLUTION: Run the 'Write' command BEFORE starting the IFC export.")
                        logger.error("This will ensure parameters are written to the model before export.")
                        logger.error("=" * 60)
                    else:
                        logger.error("Error executing configuration '{}': {}".format(config_name, error_msg))
                
                summary["config_results"].append(result)
                summary["total_elements_updated"] += result.get("elements_updated", 0)
                summary["total_elements_already_correct"] += result.get("elements_already_correct", 0)
                summary["total_parameters_copied"] += result.get("parameters_copied", 0)
                summary["total_parameters_already_correct"] += result.get("parameters_already_correct", 0)
                
            except Exception as e:

                # Check if this is a transaction error (document read-only during export)
                error_msg = str(e)
                if "not permitted" in error_msg.lower() or "read-only" in error_msg.lower() or "cannot start" in error_msg.lower():
                    logger.error("=" * 60)
                    logger.error("ERROR: Cannot write parameters during IFC export!")
                    logger.error("Revit does not allow transactions during FileExporting event.")
                    logger.error("")
                    logger.error("SOLUTION: Run the 'Write' command BEFORE starting the IFC export.")
                    logger.error("This will ensure parameters are written to the model before export.")
                    logger.error("=" * 60)
                else:
                    logger.error("Error executing configuration '{}': {}".format(config_name, error_msg))
                
                summary["config_results"].append({
                    "config_name": config_name,
                    "config_order": config_order,
                    "elements_updated": 0,
                    "parameters_copied": 0,
                    "errors": [error_msg]
                })

        # Store export info for FileExported event verification
        # We'll verify if parameters persisted after export completes
        global _pending_export_settings
        _pending_export_settings = {
            "doc": doc,
            "export_path": export_path,
            "view_id": export_view_id,
            "summary": summary  # Store summary for verification
        }

        logger.debug("IFC pre-export write complete: {} configs, {} elements updated ({} params), {} elements already correct ({} params)".format(
            summary["total_configs"],
            summary["total_elements_updated"],
            summary["total_parameters_copied"],
            summary["total_elements_already_correct"],
            summary["total_parameters_already_correct"]
        ))
        
    except Exception as ex:

        logger.error("Error in IFC export handler: {}".format(str(ex)))

def file_exported_handler(sender, args):
    """Handler for FileExported event - writes zone parameters after IFC export completes.
    
    NOTE: Parameters written here will NOT affect the current export (already completed),
    but will be available for the next export or other uses.
    
    Args:
        sender: Application object
        args: FileExportedEventArgs
    """
    global _pending_export_settings

    try:
        doc = args.Document
        if not doc:

            return

        # Check if this is an IFC export
        export_format = getattr(args, 'Format', None)

        if not export_format or "IFC" not in str(export_format).upper():

            return

        # Load all configurations
        all_configs = config.load_configs(doc)

        if not all_configs:
            logger.debug("No configurations found")
            return
        
        # Filter configurations that have write_before_ifc_export enabled
        ifc_configs = [
            cfg for cfg in all_configs 
            if cfg.get("write_before_ifc_export", False) and cfg.get("enabled", False)
        ]

        if not ifc_configs:
            logger.debug("No IFC pre-export configurations enabled.")
            return

        # Execute configurations with normal transactions (FileExported should allow this)
        summary = {
            "total_configs": len(ifc_configs),
            "config_results": [],
            "total_elements_updated": 0,
            "total_elements_already_correct": 0,
            "total_parameters_copied": 0,
            "total_parameters_already_correct": 0
        }

        for config_idx, zone_config in enumerate(ifc_configs):
            config_name = zone_config.get("name", "Unknown")

            # FileExported fires INSIDE Revit's export transaction (same as FileExporting)
            # Schedule parameter write via ExternalEvent to execute AFTER transaction completes

            # Initialize ExternalEvent if not already created
            global _post_export_write_handler, _post_export_write_event
            if _post_export_write_handler is None:
                if IExternalEventHandler is None or ExternalEvent is None:
                    logger.error("ExternalEvent classes not available - cannot schedule post-export write")
                    return
                _post_export_write_handler = PostExportWriteHandler()
                _post_export_write_event = ExternalEvent.Create(_post_export_write_handler)
            
            # Store configs and doc reference for ExternalEvent
            _post_export_write_handler.configs_to_execute = ifc_configs[:]  # Copy list
            _post_export_write_handler.doc = doc
            
            # Schedule the write via ExternalEvent (executes after transaction completes)
            try:
                _post_export_write_event.Raise()
            except Exception as event_err:

                logger.error("Failed to schedule post-export write: {}".format(str(event_err)))
        
        # Post-write verification: Check if parameters persisted
        if summary.get("total_elements_updated", 0) > 0:
            updated_element_ids = []
            for config_result in summary.get("config_results", []):
                config_updated_ids = config_result.get("updated_element_ids", [])
                if config_updated_ids:
                    updated_element_ids.extend(config_updated_ids)
            
            if updated_element_ids:
                sample_element_id = updated_element_ids[0]
                try:
                    from Autodesk.Revit.DB import ElementId, StorageType
                    sample_element = doc.GetElement(ElementId(sample_element_id))
                    if sample_element:
                        # Get config to check parameter names
                        if ifc_configs:
                            zone_config = ifc_configs[0]
                            target_params = zone_config.get("target_params", [])
                            if target_params:
                                verify_param_name = target_params[0]
                                verify_param = sample_element.LookupParameter(verify_param_name)
                except Exception as verify_err:
                    pass
        
    except Exception as ex:
        logger.error("Error in FileExported handler: {}".format(str(ex)))
    finally:
        # Clear pending settings after execution
        _pending_export_settings = None

def register_ifc_export_handler():
    """Register the IFC export event handler at application level.
    
    Returns:
        bool: True if successful, False otherwise
    """
    global _ifc_export_handler
    
    try:
        from System import EventHandler
        from Autodesk.Revit.DB.Events import FileExportingEventArgs
        
        # Get the application
        from pyrevit import HOST_APP
        app = HOST_APP.app
        
        if not app:
            logger.error("Could not get Revit application")
            return False
        
        # Register FileExporting handler if not already registered
        if _ifc_export_handler is None:
            _ifc_export_handler = EventHandler[FileExportingEventArgs](file_exporting_handler)
            app.FileExporting += _ifc_export_handler
            logger.debug("IFC export handler registered")
        
        # Register FileExported handler for verification
        global _ifc_exported_handler
        if _ifc_exported_handler is None:
            from Autodesk.Revit.DB.Events import FileExportedEventArgs
            _ifc_exported_handler = EventHandler[FileExportedEventArgs](file_exported_handler)
            app.FileExported += _ifc_exported_handler
            logger.debug("IFC exported handler registered")
        
        return True
    except Exception as e:
        logger.error("Failed to register IFC export handler: {}".format(str(e)))
        return False

def deregister_ifc_export_handler():
    """Deregister the IFC export event handler.
    
    Returns:
        bool: True if successful, False otherwise
    """
    global _ifc_export_handler
    
    try:
        from pyrevit import HOST_APP
        app = HOST_APP.app
        
        if not app:
            return False
        
        if _ifc_export_handler is not None:
            app.FileExporting -= _ifc_export_handler
            _ifc_export_handler = None
            logger.debug("IFC export handler unregistered")
        
        # Deregister FileExported handler
        global _ifc_exported_handler
        if _ifc_exported_handler is not None:
            from Autodesk.Revit.DB.Events import FileExportedEventArgs
            app.FileExported -= _ifc_exported_handler
            _ifc_exported_handler = None
            logger.debug("IFC exported handler unregistered")
        
        return True
    except Exception as e:
        logger.error("Failed to deregister IFC export handler: {}".format(str(e)))
        return False

