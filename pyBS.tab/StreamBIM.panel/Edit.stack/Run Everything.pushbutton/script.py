# -*- coding: utf-8 -*-
__title__ = "Run\nEverything"
__author__ = "Byggstyrning AB"
__doc__ = """Run all saved StreamBIM checklist configurations.

This tool applies all saved mapping configurations to all elements
in the model that have an IfcGUID parameter without any user interaction."""

import os
import sys
import clr

# Add the extension directory to the path - FIXED PATH RESOLUTION
import os.path as op
script_path = __file__
script_dir = op.dirname(script_path)
stack_dir = op.dirname(script_dir)
panel_dir = op.dirname(stack_dir)
tab_dir = op.dirname(panel_dir)
extension_dir = op.dirname(tab_dir)
lib_path = op.join(extension_dir, 'lib')

if lib_path not in sys.path:
    sys.path.append(lib_path)

# Try direct import from current directory's parent path
sys.path.append(op.dirname(op.dirname(panel_dir)))

clr.AddReference('RevitAPI')

from pyrevit import script
from pyrevit import revit

# Import StreamBIM API + shared engine
from streambim import streambim_api
from streambim.streambim_api import load_configs_with_pickle
from streambim.streambim_api import get_saved_project_id
from streambim.run_engine import (
    ConfigItem,
    ChecklistMetadataCache,
    config_from_dict,
    get_property_value,
    set_parameter_value,
    run_config,
)

# Initialize logger
logger = script.get_logger()


class RunEverythingProcessor:
    """Run Everything processor that runs all checks without UI.

    The heavy lifting lives in lib/streambim/run_engine.py, shared with the
    dockable CDE panel; this class keeps the public API used by the
    Batch Importer Tool (try_automatic_login / run_import_configurations).
    """

    def __init__(self):
        """Initialize the processor."""
        # Initialize StreamBIM API client
        self.api_client = streambim_api.StreamBIMClient()

        # Initialize config list
        self.configs = []

        # Cache for checklist metadata (runtime only, not persisted)
        self.metadata_cache = ChecklistMetadataCache(self.api_client)

        # Load configurations
        self.load_configurations()

        # Log status
        logger.info("Found {} configurations to process".format(len(self.configs)))

    def load_configurations(self):
        """Load all mapping configurations from storage."""
        logger.info("Loading configurations from storage...")
        loaded_configs = load_configs_with_pickle(revit.doc)

        if loaded_configs:
            logger.info("Found {} configurations in storage".format(len(loaded_configs)))
            for config_dict in loaded_configs:
                logger.debug("Processing config: checklist_id={}, property={}, parameter={}".format(
                    config_dict.get('checklist_id'),
                    config_dict.get('streambim_property'),
                    config_dict.get('revit_parameter')
                ))
                self.configs.append(config_from_dict(config_dict))

            logger.info("Loaded {} configurations".format(len(self.configs)))
        else:
            logger.info("No configurations found in storage")

    def try_automatic_login(self):
        """Attempt to automatically log in using saved tokens."""
        # Load tokens from file first
        self.api_client.load_tokens()

        # Check if token exists
        if self.api_client.idToken:
            logger.info("Found saved StreamBIM login...")

            # Try to load saved project ID
            saved_project_id = get_saved_project_id(revit.doc)
            if saved_project_id:
                self.api_client.set_current_project(saved_project_id)
                logger.info("Using saved project ID: {}".format(saved_project_id))

            return True
        else:
            logger.error("No saved StreamBIM login found. Please log in using the ChecklistImporter first.")
            return False

    def get_checklist_metadata(self, checklist_id):
        """Get checklist metadata (group-by and building_id) from cache or API."""
        return self.metadata_cache.get(checklist_id)

    def process_single_configuration(self, config, config_index, total_configs):
        """Process a single configuration with its own transaction.
        Returns a tuple of (processed_count, updated_count)."""
        processed_count, updated_count = run_config(
            self.api_client, revit.doc, config, self.metadata_cache)
        logger.info("Completed configuration {}/{}: {} - Processed: {}, Updated: {}".format(
            config_index + 1, total_configs, config.DisplayName,
            processed_count, updated_count))
        return (processed_count, updated_count)

    def run_import_configurations(self):
        """Run import for all configurations."""
        if not self.configs:
            logger.info("No configurations to process. Exiting.")
            return

        logger.info("Starting batch import process for {} configurations".format(len(self.configs)))

        try:
            total_processed = 0
            total_updated = 0

            for i, config in enumerate(self.configs):
                logger.info("==== Processing configuration {}/{}: {} ====".format(
                    i + 1, len(self.configs), config.DisplayName
                ))
                logger.info("Checklist: {} (ID: {})".format(config.ChecklistName, config.checklist_id))
                logger.info("Property: {} -> Parameter: {}".format(config.streambim_property, config.revit_parameter))
                logger.info("Mapping enabled: {}".format(config.mapping_enabled))

                if not config.checklist_id:
                    logger.info("Skipping configuration - no checklist ID")
                    config.elements_processed = 0
                    config.elements_updated = 0
                    continue

                processed_count, updated_count = self.process_single_configuration(config, i, len(self.configs))

                total_processed += processed_count
                total_updated += updated_count

            logger.info("Batch import completed. Processed {} configurations. Updated {}/{} elements.".format(
                len(self.configs), total_updated, total_processed))

        except Exception as e:
            logger.error("Error running batch import: {}".format(str(e)))
            import traceback
            logger.error("Stack trace: {}".format(traceback.format_exc()))


# Main execution
if __name__ == '__main__':
    logger.info("Starting Run Everything script...")

    # Create processor and run without UI
    processor = RunEverythingProcessor()

    # Check if we have valid login
    if processor.try_automatic_login():
        # Run import process
        processor.run_import_configurations()
    else:
        logger.error("Cannot proceed without StreamBIM login. Please run the ChecklistImporter tool first to log in.")
