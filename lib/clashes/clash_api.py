# -*- coding: utf-8 -*-
"""Clash detection API client for accessing clash detection services.

This module provides functionality to communicate with clash detection APIs
and manage clash test results.
"""

import json
import urllib2
import base64
from collections import namedtuple

# Import Revit API and pyRevit modules
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

from pyrevit import script
from pyrevit import revit

# Import extensible storage
import sys
import os.path as op
current_dir = op.dirname(__file__)
lib_dir = op.dirname(current_dir)
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

try:
    from extensible_storage import BaseSchema, simple_field
except ImportError:
    print("Warning: extensible_storage module could not be imported in clash_api.py")
    BaseSchema = object
    def simple_field(**kwargs):
        def decorator(func):
            return func
        return decorator

# Initialize logger
logger = script.get_logger()

# Define the ClashExplorerSettingsSchema
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
    """Get existing or create new Clash Explorer settings storage element."""
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
                # Check if this storage has our schema
                entity = ds.GetEntity(ClashExplorerSettingsSchema.schema)
                if entity.IsValid():
                    logger.debug("Found existing Clash Explorer settings storage")
                    return ds
            except Exception as e:
                logger.debug("Error checking storage entity: {}".format(str(e)))
                continue
        
        logger.debug("No existing Clash Explorer settings storage found, creating new one...")
        # If not found, create a new one
        with revit.Transaction("Create Clash Explorer Settings Storage", doc):
            new_storage = ExtensibleStorage.DataStorage.Create(doc)
            logger.debug("Created new Clash Explorer settings storage")
            return new_storage
            
    except Exception as e:
        logger.error("Error in get_or_create_clash_settings_storage: {}".format(str(e)))
        return None

def save_clash_settings(doc, api_url, api_key):
    """Save Clash Explorer settings to extensible storage."""
    if not doc:
        logger.error("No active document available")
        return False
        
    try:
        # Get or create data storage
        storage = get_or_create_clash_settings_storage(doc)
        if not storage:
            logger.error("Failed to get or create Clash Explorer storage")
            return False
        
        # Save to storage
        with revit.Transaction("Save Clash Explorer Settings", doc):
            with ClashExplorerSettingsSchema(storage) as entity:
                entity.set("api_url", api_url)
                entity.set("api_key", api_key)
                
        logger.debug("Saved Clash Explorer settings")
        return True
    except Exception as e:
        logger.error("Error saving Clash Explorer settings: {}".format(str(e)))
        return False

def load_clash_settings(doc):
    """Load Clash Explorer settings from extensible storage."""
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
                # Check if this storage has our schema
                entity = ds.GetEntity(ClashExplorerSettingsSchema.schema)
                if entity.IsValid():
                    schema = ClashExplorerSettingsSchema(ds)
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
    
    def get_clash_tests(self, project_id=None):
        """Get list of clash tests for a project.
        
        Args:
            project_id: Optional project ID filter
            
        Returns:
            List of clash test objects
        """
        endpoint = "/api/v1/clash-tests"
        if project_id:
            endpoint = "{}?project_id={}".format(endpoint, project_id)
            
        result = self._make_request(endpoint)
        if result:
            return result.get('data', [])
        return []
    
    def get_clash_results(self, clash_test_id):
        """Get clash results for a specific clash test.
        
        Args:
            clash_test_id: ID of the clash test
            
        Returns:
            List of clash result groups
        """
        endpoint = "/api/v1/clash-tests/{}/results".format(clash_test_id)
        result = self._make_request(endpoint)
        if result:
            return result.get('data', [])
        return []
    
    def get_clash_groups(self, clash_test_id):
        """Get clash groups for a specific clash test, sorted by count.
        
        Args:
            clash_test_id: ID of the clash test
            
        Returns:
            List of clash groups sorted by clash count (descending)
        """
        endpoint = "/api/v1/clash-tests/{}/groups".format(clash_test_id)
        result = self._make_request(endpoint)
        if result:
            groups = result.get('data', [])
            # Sort by count in descending order
            groups.sort(key=lambda x: x.get('clash_count', 0), reverse=True)
            return groups
        return []
    
    def get_clash_details(self, clash_id):
        """Get detailed information for a specific clash.
        
        Args:
            clash_id: ID of the clash
            
        Returns:
            Clash details object
        """
        endpoint = "/api/v1/clashes/{}".format(clash_id)
        result = self._make_request(endpoint)
        if result:
            return result.get('data', {})
        return {}
    
    def get_clashes_by_guids(self, guids):
        """Get clashes involving specific element GUIDs.
        
        Args:
            guids: List of IFC GUIDs to search for
            
        Returns:
            List of clash groups containing these GUIDs
        """
        endpoint = "/api/v1/clashes/search"
        data = {'guids': guids}
        result = self._make_request(endpoint, method="POST", data=data)
        if result:
            return result.get('data', [])
        return []
