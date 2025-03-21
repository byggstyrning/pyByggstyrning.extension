# -*- coding: utf-8 -*-
"""Batch processor for StreamBIM 'Run Everything' functionality."""

import os
import datetime
import sys
import clr
import imp
import time
import traceback
from pyrevit import HOST_APP, forms
from Autodesk.Revit.DB import TransactWithCentralOptions, SynchronizeWithCentralOptions, RelinquishOptions
from Autodesk.Revit.DB import Transaction, TransactionStatus, OpenOptions, DetachFromCentralOption, ModelPathUtils, SaveAsOptions

# Required assemblies
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')

# Simple file logger
class Logger:
    def __init__(self):
        log_dir = os.path.join(os.path.dirname(__file__), "log")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        self.log_file = os.path.join(log_dir, "streambim_checklist_batch_{}.log".format(datetime.datetime.now().strftime("%Y%m%d_%H%M%S")))
        with open(self.log_file, 'wb') as f:
            f.write("=== StreamBIM Checklist Batch Processing Log ===\nStarted: {}\n\n".format(datetime.datetime.now()).encode('utf-8'))
    
    def log(self, message, level="DEBUG"):
        with open(self.log_file, 'ab') as f:
            f.write((datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + " [" + level + "] " + message + "\n").encode('utf-8'))
    
    def debug(self, message): self.log(message, "DEBUG")
    def error(self, message): self.log(message, "ERROR")

# Initialize
logger = Logger()
logger.debug("Starting batch processing")

# Get the current script's directory and resolve path to the original script
current_dir = os.path.dirname(os.path.abspath(__file__))
panel_dir = os.path.dirname(current_dir)
original_script_path = os.path.join(panel_dir, "Edit.stack", "Run Everything.pushbutton", "script.py")
logger.debug("Script path: {}".format(original_script_path))

# Ask user if they want to create detached copies instead of saving directly
detach_option = forms.alert(
    "How would you like to save processed files?",
    options=["Save directly", "Create detached copies"],
    title="Batch Processing Options"
)
create_detached_copies = (detach_option == "Create detached copies")
logger.debug("Create detached copies: {}".format(create_detached_copies))

# Track current open documents to avoid conflicts
currently_open_docs = set()

# Process and save models
for model in __models__:
    open_doc = None
    temp_path = None
    
    try:
        # Track this model path to avoid reopening
        model_path = os.path.normpath(model).lower()
        if model_path in currently_open_docs:
            logger.error("Model is already being processed in another session: {}".format(model))
            continue
            
        currently_open_docs.add(model_path)
        
        # Open model with options if needed
        opening_start_time = time.time()
        logger.debug("Opening model: {}".format(model))
        
        try:
            # Close all other documents first
            app = HOST_APP.app
            
            # Create a list of documents to close (all except the one we're processing)
            docs_to_close = []
            for open_doc_id in app.Documents:
                if open_doc_id.PathName.lower() != model_path:
                    docs_to_close.append(open_doc_id)
            
            # Close other documents
            for doc_to_close in docs_to_close:
                logger.debug("Closing document: {}".format(doc_to_close.PathName))
                try:
                    # Make sure we're not in a view of the document we're closing
                    if HOST_APP.uiapp.ActiveUIDocument and HOST_APP.uiapp.ActiveUIDocument.Document.PathName == doc_to_close.PathName:
                        HOST_APP.uiapp.ActiveUIDocument = None
                    app.CloseDocument(doc_to_close)
                except Exception as ex:
                    logger.error("Error closing document: {} - {}".format(doc_to_close.PathName, ex))
            
            # Now open the new document, with detach option if requested
            if create_detached_copies:
                # Create a detached copy
                model_path_obj = ModelPathUtils.ConvertUserVisiblePathToModelPath(model)
                open_options = OpenOptions()
                open_options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
                open_options.Audit = False
                
                logger.debug("Opening with detach option")
                doc = app.OpenDocumentFile(model_path_obj, open_options)
                # We'll set the uidoc later if needed
                uidoc = HOST_APP.uiapp.ActiveUIDocument
            else:
                # Standard open
                uidoc = HOST_APP.uiapp.OpenAndActivateDocument(model)
                doc = uidoc.Document
                
            open_doc = doc  # Keep track of opened document
            logger.debug("Opening model took: {:.2f} seconds".format(time.time() - opening_start_time))
            
            if create_detached_copies:
                # Generate temporary path for detached copy
                file_dir = os.path.dirname(model)
                file_name = os.path.basename(model)
                name_part, ext = os.path.splitext(file_name)
                temp_path = os.path.join(file_dir, "{}_SBImport_{}{}".format(
                    name_part, 
                    datetime.datetime.now().strftime("%Y%m%d_%H%M%S"),
                    ext
                ))
                logger.debug("Will save detached copy to: {}".format(temp_path))
        except Exception as open_ex:
            logger.error("Error during document open: {}".format(open_ex))
            logger.error("Open error trace: {}".format(traceback.format_exc()))
            raise
        
        try:
            processing_start_time = time.time()
            original_script = imp.load_source('run_everything_script', original_script_path)
            processor = original_script.RunEverythingProcessor()
            
            if processor.try_automatic_login():
                logger.debug("Logged in to StreamBIM")
                
                # Check for open transactions and close them
                while doc.IsModifiable:
                    logger.debug("Document has open transaction, attempting to close it")
                    try:
                        # Try to get current transaction and close it
                        if doc.GetTransactionNames().Count > 0:
                            last_tx_name = doc.GetTransactionNames().get_Item(doc.GetTransactionNames().Count - 1)
                            logger.debug("Found open transaction: {}".format(last_tx_name))
                            doc.RollbackTransaction()
                            logger.debug("Rolled back transaction: {}".format(last_tx_name))
                        else:
                            # No named transaction but still modifiable - create and rollback dummy transaction
                            logger.debug("No named transaction found but document is modifiable")
                            dummy_tx = Transaction(doc, "CleanupDummy")
                            dummy_tx.Start()
                            dummy_tx.RollBack()
                            logger.debug("Rolled back dummy transaction")
                            break
                    except Exception as tx_ex:
                        logger.error("Error cleaning up transactions: {}".format(tx_ex))
                        break
                
                # Run processor in its own transaction
                logger.debug("Starting import transaction")
                tx = Transaction(doc, "StreamBIM Import")
                tx.Start()
                
                try:
                    processor.run_import_configurations()
                    logger.debug("Processed model: {}".format(doc.Title))
                    
                    # Calculate processing time
                    logger.debug("Processing time: {:.2f} seconds".format(time.time() - processing_start_time))
                    
                    # Commit the transaction before saving
                    logger.debug("Committing import transaction")
                    tx_result = tx.Commit()
                    
                    if tx_result == TransactionStatus.Committed:
                        logger.debug("Transaction committed successfully")
                    else:
                        logger.error("Transaction failed to commit with status: {}".format(tx_result))
                        continue  # Skip saving if transaction failed
                        
                except Exception as proc_ex:
                    logger.error("Error during processing: {}".format(proc_ex))
                    if tx.HasStarted():
                        logger.debug("Rolling back transaction due to error")
                        tx.RollBack()
                    raise
                
                # Save/Sync after all transactions are closed
                saving_start_time = time.time()
                try:
                    if create_detached_copies:
                        # Save as a detached copy
                        logger.debug("Saving as detached copy to: {}".format(temp_path))
                        save_options = SaveAsOptions()
                        save_options.OverwriteExistingFile = True
                        save_options.Compact = True
                        doc.SaveAs(temp_path, save_options)
                        logger.debug("Successfully saved detached copy")
                    elif doc.IsWorkshared:
                        logger.debug("Model is workshared, syncing with central")
                        # Configure sync options
                        trans_options = TransactWithCentralOptions()
                        sync_options = SynchronizeWithCentralOptions()
                        relinq_options = RelinquishOptions(True)
                        relinq_options.StandardWorksets = True
                        relinq_options.ViewWorksets = True
                        relinq_options.FamilyWorksets = True
                        relinq_options.UserWorksets = True
                        relinq_options.CheckedOutElements = True
                        sync_options.SetRelinquishOptions(relinq_options)
                        sync_options.Compact = True
                        
                        logger.debug("Starting sync with central")
                        doc.SynchronizeWithCentral(trans_options, sync_options)
                        logger.debug("Sync completed successfully in {:.2f} seconds".format(time.time() - saving_start_time))
                    else:
                        logger.debug("Model is not workshared, saving locally")
                        save_options = SaveAsOptions()
                        save_options.OverwriteExistingFile = True
                        save_options.Compact = True
                        doc.SaveAs(doc.PathName, save_options)
                        logger.debug("Save completed successfully in {:.2f} seconds".format(time.time() - saving_start_time))
                    
                    logger.debug("Saved: {}".format(doc.Title))
                except Exception as save_ex:
                    logger.error("Error during save/sync: {}".format(save_ex))
                    logger.error("Save/sync error trace: {}".format(traceback.format_exc()))
            else:
                logger.error("StreamBIM login failed")
        except Exception as e:
            logger.error("Error: {}".format(e))
            logger.error("Trace: {}".format(traceback.format_exc()))
    except Exception as e:
        logger.error("Failed to process {}: {}".format(model, e))
    finally:
        # Try to close document
        try:
            if open_doc:
                logger.debug("Attempting to close document")
                open_doc_path = open_doc.PathName
                
                # Make sure document is not active
                if HOST_APP.uiapp.ActiveUIDocument and HOST_APP.uiapp.ActiveUIDocument.Document.PathName == open_doc_path:
                    HOST_APP.uiapp.ActiveUIDocument = None
                
                HOST_APP.app.CloseDocument(open_doc)
                logger.debug("Document closed: {}".format(open_doc_path))
                
                # Remove from tracking set
                model_path = os.path.normpath(model).lower()
                if model_path in currently_open_docs:
                    currently_open_docs.remove(model_path)
        except Exception as close_ex:
            logger.error("Error closing document: {}".format(close_ex))

logger.debug("Batch processing complete")