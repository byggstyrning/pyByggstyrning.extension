# -*- coding: utf-8 -*-
"""Batch processor for StreamBIM 'Run Everything' functionality."""

import os
import datetime
import sys
import clr
import imp
import time
import traceback
from pyrevit import HOST_APP
from Autodesk.Revit.DB import TransactWithCentralOptions, SynchronizeWithCentralOptions, RelinquishOptions
from Autodesk.Revit.DB import SaveAsOptions

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

# Track current open documents to avoid conflicts
currently_open_docs = set()

# Process and save models
for model in __models__:
    try:
        # Open model
        logger.debug("Opening model: {}".format(model))
        uidoc = HOST_APP.uiapp.OpenAndActivateDocument(model)
        doc = uidoc.Document
        
        # Process using original script
        try:
            original_script = imp.load_source('run_everything_script', original_script_path)
            processor = original_script.RunEverythingProcessor()
            
            if processor.try_automatic_login():
                logger.debug("Logged in to StreamBIM")
                
                # Process the model - let the processor handle its own transactions
                processing_start_time = time.time()
                try:
                    processor.run_import_configurations()
                    logger.debug("Processed model: {}".format(doc.Title))
                    logger.debug("Processing time: {:.2f} seconds".format(time.time() - processing_start_time))
                except Exception as proc_ex:
                    logger.error("Error during processing: {}".format(proc_ex))
                    logger.error("Processing error trace: {}".format(traceback.format_exc()))
                    raise
                
                # Save/Sync
                if doc.IsWorkshared:
                    logger.debug("Syncing with central")
                    trans_options = TransactWithCentralOptions()
                    sync_options = SynchronizeWithCentralOptions()
                    sync_options.SetRelinquishOptions(RelinquishOptions(False))
                    doc.SynchronizeWithCentral(trans_options, sync_options)
                else:
                    logger.debug("Saving model")
                    doc.Save()
                logger.debug("Saved: {}".format(doc.Title))
            else:
                logger.error("StreamBIM login failed")
        except Exception as e:
            logger.error("Error: {}".format(e))
            logger.error("Trace: {}".format(traceback.format_exc()))
    except Exception as e:
        logger.error("Failed to process {}: {}".format(model, e))

logger.debug("Batch processing complete")