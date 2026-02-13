# -*- coding: utf-8 -*-
"""Clash detection API client for accessing clash detection services.

This module provides functionality to communicate with clash detection APIs
and manage clash test results.
"""

import json
import urllib2
from collections import namedtuple

# Import Revit API and pyRevit modules
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

from pyrevit import script
from pyrevit import revit

# Import extensible storage using the generic implementation
import sys
import os.path as op
current_dir = op.dirname(__file__)
lib_dir = op.dirname(current_dir)
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

from extensible_storage import BaseSchema, simple_field

# Initialize logger
logger = script.get_logger()

# Define the ClashExplorerSettingsSchema using the generic BaseSchema
class ClashExplorerSettingsSchema(BaseSchema):
    """Schema for storing Clash Explorer settings using extensible storage"""
    
    # Generate a unique GUID for Clash Explorer settings
    guid = "a7f3d8e2-5c1a-4b9e-8f2d-3e4a5b6c7d8e"
    
    @simple_field(value_type="string")
    def api_url():
        """The clash detection API URL"""
        
    @simple_field(value_type="string")
    def api_key():
        """The API key for authentication"""

def get_or_create_clash_settings_storage(doc):
    """Get existing or create new Clash Explorer settings storage element.
    
    Uses the generic pattern from extensible storage.
    """
    if not doc:
        logger.error("No active document available")
        return None
        
    try:
        logger.debug("Searching for Clash Explorer settings storage...")
        data_storages = FilteredElementCollector(doc)\
            .OfClass(ExtensibleStorage.DataStorage)\
            .ToElements()
        
        # Look for our storage with the schema
        for ds in data_storages:
            try:
                entity = ds.GetEntity(ClashExplorerSettingsSchema.schema)
                if entity.IsValid():
                    logger.debug("Found existing Clash Explorer settings storage")
                    return ds
            except Exception as e:
                logger.debug("Error checking storage entity: {}".format(str(e)))
                continue
        
        logger.debug("No existing Clash Explorer settings storage found, creating new one...")
        # Create new storage in transaction
        with revit.Transaction("Create Clash Explorer Settings Storage", doc):
            new_storage = ExtensibleStorage.DataStorage.Create(doc)
            logger.debug("Created new Clash Explorer settings storage")
            return new_storage
            
    except Exception as e:
        logger.error("Error in get_or_create_clash_settings_storage: {}".format(str(e)))
        return None

def save_clash_settings(doc, api_url, api_key):
    """Save Clash Explorer settings to extensible storage.
    
    Uses the generic BaseSchema pattern with context manager.
    """
    if not doc:
        logger.error("No active document available")
        return False
        
    try:
        storage = get_or_create_clash_settings_storage(doc)
        if not storage:
            logger.error("Failed to get or create Clash Explorer storage")
            return False
        
        # Use context manager pattern from BaseSchema for automatic transaction handling
        with ClashExplorerSettingsSchema(storage) as entity:
            entity.set("api_url", api_url)
            entity.set("api_key", api_key)
                
        logger.debug("Saved Clash Explorer settings")
        return True
    except Exception as e:
        logger.error("Error saving Clash Explorer settings: {}".format(str(e)))
        return False

def load_clash_settings(doc):
    """Load Clash Explorer settings from extensible storage.
    
    Uses the generic BaseSchema pattern.
    """
    if not doc:
        logger.error("No active document available")
        return None, None
        
    try:
        logger.debug("Loading Clash Explorer settings...")
        data_storages = FilteredElementCollector(doc)\
            .OfClass(ExtensibleStorage.DataStorage)\
            .ToElements()
        
        for ds in data_storages:
            try:
                entity = ds.GetEntity(ClashExplorerSettingsSchema.schema)
                if entity.IsValid():
                    # Use the generic BaseSchema to read data
                    schema = ClashExplorerSettingsSchema(ds, update=False)
                    if schema.is_valid:
                        api_url = schema.get("api_url")
                        api_key = schema.get("api_key")
                        logger.debug("Loaded Clash Explorer settings")
                        return api_url, api_key
            except Exception as e:
                logger.debug("Error checking storage entity: {}".format(str(e)))
                continue
        
        logger.debug("No saved Clash Explorer settings found")
        return None, None
    except Exception as e:
        logger.error("Error in load_clash_settings: {}".format(str(e)))
        return None, None

# Clash API client
class ClashAPIClient:
    """Client for interacting with clash detection API."""
    
    def __init__(self, base_url=None, api_key=None):
        """Initialize the Clash API client.
        
        Args:
            base_url: Base URL for the clash detection API
            api_key: API key for authentication
        """
        self.base_url = base_url
        self.api_key = api_key
        self.last_error = None
    
    def _make_request(self, endpoint, method="GET", data=None):
        """Make an HTTP request to the API.
        
        Args:
            endpoint: API endpoint path
            method: HTTP method (GET, POST, etc.)
            data: Optional data payload for POST requests
            
        Returns:
            Parsed JSON response or None on error
        """
        if not self.base_url or not self.api_key:
            self.last_error = "API URL or API key not configured"
            return None
            
        try:
            url = "{}{}".format(self.base_url, endpoint)
            logger.debug("Making {} request to: {}".format(method, url))
            
            # Prepare request
            if data:
                req = urllib2.Request(url, data=json.dumps(data))
                req.add_header('Content-Type', 'application/json')
            else:
                req = urllib2.Request(url)
            
            # Add authentication header
            req.add_header('Authorization', 'Bearer {}'.format(self.api_key))
            req.add_header('Accept', 'application/json')
            
            # Make request
            response = urllib2.urlopen(req)
            result = json.loads(response.read())
            
            self.last_error = None
            return result
            
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Authentication failed. Please check your API key."
            elif e.code == 404:
                error_message = "API endpoint not found: {}".format(endpoint)
                
            self.last_error = error_message
            logger.error("API request error: {}".format(error_message))
            return None
        except Exception as e:
            self.last_error = str(e)
            logger.error("API request error: {}".format(str(e)))
            return None
    
    def get_clash_sets(self, project_id=None):
        """Get list of clash sets from ifcclash JSON format.
        
        Expected format: list[ClashSet] where each ClashSet contains:
        - name: str
        - mode: "intersection" | "collision" | "clearance"
        - a: list of sources
        - b: optional list of sources
        - clashes: dict of "guid_a-guid_b" -> ClashResult
        
        Args:
            project_id: Optional project ID filter
            
        Returns:
            List of clash set objects
        """
        endpoint = "/api/v1/clash-sets"
        if project_id:
            endpoint = "{}?project_id={}".format(endpoint, project_id)
            
        result = self._make_request(endpoint)
        if result:
            # ifcclash returns a list directly, not wrapped in 'data'
            if isinstance(result, list):
                return result
            return result.get('data', [])
        return []
    
    def get_clash_set(self, clash_set_name):
        """Get a specific clash set by name.
        
        Args:
            clash_set_name: Name of the clash set
            
        Returns:
            ClashSet object or None
        """
        endpoint = "/api/v1/clash-sets/{}".format(clash_set_name)
        result = self._make_request(endpoint)
        if result:
            return result
        return None
    
    def get_smart_grouped_clashes(self, max_clustering_distance=3.0):
        """Get smart-grouped clashes from ifcclash.
        
        The smart_group_clashes format returns:
        {
            "ClashSetName": [
                {
                    "ClashSetName - 1": [["guid_a", "guid_b"], ...],
                    "ClashSetName - 2": [["guid_a", "guid_b"], ...]
                }
            ]
        }
        
        Args:
            max_clustering_distance: Maximum distance for grouping clashes
            
        Returns:
            Dictionary of grouped clashes
        """
        endpoint = "/api/v1/clash-sets/smart-groups?max_distance={}".format(max_clustering_distance)
        result = self._make_request(endpoint)
        if result:
            return result
        return {}
    
    def get_clashes_by_guids(self, guids):
        """Get clashes involving specific element GUIDs.
        
        Searches through clash results for matching a_global_id or b_global_id.
        
        Args:
            guids: List of IFC GUIDs to search for
            
        Returns:
            List of clashes containing these GUIDs
        """
        endpoint = "/api/v1/clashes/search"
        data = {'guids': guids}
        result = self._make_request(endpoint, method="POST", data=data)
        if result:
            return result.get('data', [])
        return []
