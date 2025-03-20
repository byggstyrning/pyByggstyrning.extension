import json
import urllib2
import base64
from collections import namedtuple
import os
import pickle

# Import Revit API and pyRevit modules
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

from pyrevit import script
from pyrevit import revit

# Import extensible storage
# Add the parent directory to sys.path to ensure the extensible_storage module can be found
import sys
import os.path as op
current_dir = op.dirname(__file__)
lib_dir = op.dirname(current_dir)
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

try:
    from extensible_storage import BaseSchema, simple_field
except ImportError:
    # Handle the import error gracefully
    print("Warning: extensible_storage module could not be imported in streambim_api.py")
    BaseSchema = object
    def simple_field(**kwargs):
        def decorator(func):
            return func
        return decorator

# Initialize logger
logger = script.get_logger()

# Define the StreamBIMSettingsSchema
class StreamBIMSettingsSchema(BaseSchema):
    """Schema for storing StreamBIM settings and configurations using pickle serialization"""
    
    # Generate a completely new GUID to avoid conflicts with existing schemas
    guid = "82c63e54-9b4c-45c8-a290-47d3b71e8a56"
    
    @simple_field(value_type="string")
    def project_id():
        """The selected StreamBIM project ID"""
        
    @simple_field(value_type="string")
    def pickled_configs():
        """Base64 encoded pickle of all checklist configurations"""

def get_or_create_settings_storage(doc):
    """Get existing or create new StreamBIM settings storage element."""
    if not doc:
        logger.error("No active document available")
        return None
        
    try:
        logger.debug("Searching for StreamBIM settings storage...")
        data_storages = FilteredElementCollector(doc)\
            .OfClass(ExtensibleStorage.DataStorage)\
            .ToElements()
        
        # Look for our storage with the schema
        for ds in data_storages:
            try:
                # Check if this storage has our schema
                entity = ds.GetEntity(StreamBIMSettingsSchema.schema)
                if entity.IsValid():
                    logger.debug("Found existing StreamBIM settings storage")
                    return ds
            except Exception as e:
                logger.debug("Error checking storage entity: {}".format(str(e)))
                continue
        
        logger.debug("No existing StreamBIM settings storage found, creating new one...")
        # If not found, create a new one
        with revit.Transaction("Create StreamBIM Settings Storage", doc):
            new_storage = ExtensibleStorage.DataStorage.Create(doc)
            logger.debug("Created new StreamBIM settings storage")
            return new_storage
            
    except Exception as e:
        logger.error("Error in get_or_create_settings_storage: {}".format(str(e)))
        return None

def load_configs_with_pickle(doc):
    """Load configurations from StreamBIM storage using pickle serialization."""
    if not doc:
        logger.error("No active document available")
        return []
        
    try:
        # Get storage
        storage = get_or_create_settings_storage(doc)
        if not storage:
            logger.error("Failed to get or create StreamBIM storage")
            return []
            
        # Load configurations
        schema = StreamBIMSettingsSchema(storage)
        if not schema.is_valid:
            logger.debug("Invalid StreamBIM schema")
            return []
            
        pickled_configs = schema.get("pickled_configs")
        if not pickled_configs:
            logger.debug("No configurations found in storage")
            return []
            
        # Decode and unpickle
        try:
            decoded_data = base64.b64decode(pickled_configs)
            configs = pickle.loads(decoded_data)
            logger.debug("Successfully loaded {} configurations".format(len(configs)))
            return configs
        except Exception as e:
            logger.error("Error unpickling configurations: {}".format(str(e)))
            return []
    except Exception as e:
        logger.error("Error loading configurations: {}".format(str(e)))
        return []

def save_configs_with_pickle(doc, config_dicts):
    """Save configurations to StreamBIM storage using pickle serialization."""
    if not doc:
        logger.error("No active document available")
        return False
        
    try:
        # Get or create data storage
        storage = get_or_create_settings_storage(doc)
        if not storage:
            logger.error("Failed to get or create StreamBIM storage")
            return False
        
        # Pickle and encode the data
        try:
            pickled_data = pickle.dumps(config_dicts)
            encoded_data = base64.b64encode(pickled_data)
            
            # Save to storage
            with revit.Transaction("Save StreamBIM Configurations", doc):
                with StreamBIMSettingsSchema(storage) as entity:
                    # Preserve existing project_id if present
                    schema = StreamBIMSettingsSchema(storage)
                    if schema.is_valid:
                        project_id = schema.get("project_id")
                        if project_id:
                            entity.set("project_id", project_id)
                            
                    # Save configurations
                    entity.set("pickled_configs", encoded_data)
                    
            logger.debug("Saved {} configurations to StreamBIM storage".format(len(config_dicts)))
            return True
        except Exception as e:
            logger.error("Error pickling configurations: {}".format(str(e)))
            return False
    except Exception as e:
        logger.error("Error saving configurations: {}".format(str(e)))
        return False

def get_saved_project_id(doc):
    """Get the saved StreamBIM project ID from extensible storage."""
    if not doc:
        logger.error("No active document available")
        return None
        
    try:
        logger.debug("Searching for StreamBIM settings storage...")
        data_storages = FilteredElementCollector(doc)\
            .OfClass(ExtensibleStorage.DataStorage)\
            .ToElements()
        
        for ds in data_storages:
            try:
                # Check if this storage has our schema
                entity = ds.GetEntity(StreamBIMSettingsSchema.schema)
                if entity.IsValid():
                    schema = StreamBIMSettingsSchema(ds)
                    if schema.is_valid:
                        project_id = schema.get("project_id")
                        if project_id:
                            logger.debug("Found saved project ID: {0}".format(project_id))
                            return project_id
            except Exception as e:
                logger.debug("Error checking storage entity: {0}".format(str(e)))
                continue
        
        logger.debug("No saved project ID found")
        return None
    except Exception as e:
        logger.error("Error in get_saved_project_id: {0}".format(str(e)))
        return None

# StreamBIM API client
class StreamBIMClient:
    def __init__(self, base_url="https://app.streambim.com"):
        self.base_url = base_url
        self.idToken = None
        self.accessToken = None
        self.username = None
        self.projects = []
        self.current_project = None
        self.last_error = None
        
        # Load saved tokens if they exist
        self.token_file = os.path.join(os.getenv('APPDATA'), 'pyBS', 'tokens.json')
        self.load_tokens()
    
    def load_tokens(self):
        """Load saved tokens from file."""
        try:
            if os.path.exists(self.token_file):
                with open(self.token_file, 'r') as f:
                    data = json.load(f)
                    self.idToken = data.get('idToken')
                    self.accessToken = data.get('accessToken')
                    self.username = data.get('username')
        except Exception as e:
            print("Error loading tokens: {}".format(str(e)))
            self.idToken = None
            self.accessToken = None
            self.username = None
    
    def save_tokens(self):
        """Save tokens to file."""
        try:
            # Create directory if it doesn't exist
            token_dir = os.path.dirname(self.token_file)
            if not os.path.exists(token_dir):
                os.makedirs(token_dir)
                
            with open(self.token_file, 'w') as f:
                json.dump({
                    'idToken': self.idToken,
                    'accessToken': self.accessToken,
                    'username': self.username
                }, f)
        except Exception as e:
            print("Error saving tokens: {}".format(str(e)))
    
    def clear_tokens(self):
        """Clear saved tokens."""
        self.idToken = None
        self.accessToken = None
        self.username = None
        try:
            if os.path.exists(self.token_file):
                os.remove(self.token_file)
        except Exception as e:
            print("Error clearing tokens: {}".format(str(e)))
        
    def login(self, username, password):
        """Login to StreamBIM and get authentication token"""
        try:
            url = "{}/mgw/api/v2/login".format(self.base_url)
            data = json.dumps({
                "username": username,
                "password": password
            })
            
            headers = {
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8"
            }
            
            req = urllib2.Request(url, data=data.encode('utf-8'))
            for key, value in headers.items():
                req.add_header(key, value)
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read())

            if 'idToken' in result:
                self.accessToken = result['accessToken']
                self.idToken = result['idToken']
                self.username = username  # Store the username
                self.save_tokens()
                self.last_error = None
                return True
            
            self.last_error = "No token in response"
            return False
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Invalid username or password: " + e.reason + " " + e.read()
            elif e.code == 404:
                error_message = "Server URL not found"
                
            self.last_error = error_message
            print("Login error: {}".format(error_message))
            return False
        except Exception as e:
            self.last_error = str(e)
            print("Login error: {}".format(str(e)))
            return False
    
    def _decode_utf8(self, data):
        """Recursively decode UTF-8 strings in the API response."""
        if isinstance(data, dict):
            return {self._decode_utf8(key): self._decode_utf8(value) for key, value in data.items()}
        elif isinstance(data, list):
            return [self._decode_utf8(item) for item in data]
        elif isinstance(data, tuple):
            return tuple(self._decode_utf8(item) for item in data)
        elif isinstance(data, set):
            return {self._decode_utf8(item) for item in data}
        elif isinstance(data, str):
            try:
                return data.decode('utf-8')
            except UnicodeError:
                return data
        elif isinstance(data, unicode):
            return data
        else:
            return data
        
    def get_projects(self):
        """Get list of available projects"""
        if not self.idToken:
            self.last_error = "Not authenticated"
            return []
            
        try:
            url = "{}/mgw/api/v3/project-links?filter%5Bactive%5D=true".format(self.base_url)
            req = urllib2.Request(url)
            req.add_header('Authorization', 'Bearer {}'.format(self.idToken))
            req.add_header('Accept', 'application/vnd.api+json')
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read())
            
            # Decode UTF-8 strings in the response
            result = self._decode_utf8(result)
            
            self.projects = result.get('data', [])
            return self.projects
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Authentication failed. Please log in again."
                
            self.last_error = error_message
            print("Error getting projects: {}".format(error_message))
            return []
        except Exception as e:
            self.last_error = str(e)
            print("Error getting projects: {}".format(str(e)))
            return []
    
    def set_current_project(self, project_id):
        """Set current project by ID"""
        self.current_project = project_id
    
    def get_checklists(self):
        """Get available checklists for the current project"""
        if not self.idToken or not self.current_project:
            self.last_error = "Not authenticated or no project selected"
            return []
            
        try:
            url = "{}/project-{}/api/v1/v2/checklists?filter[isDraft]=false".format(
                self.base_url, self.current_project
            )
            
            req = urllib2.Request(url)
            req.add_header('Authorization', 'Bearer {}'.format(self.idToken))
            req.add_header('Accept', 'application/vnd.api+json')
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read())
            
            # Decode UTF-8 strings in the response
            result = self._decode_utf8(result)
            
            return result.get('data', [])
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Authentication failed. Please log in again."
                
            self.last_error = error_message
            print("Error getting checklists: {}".format(error_message))
            return []
        except Exception as e:
            self.last_error = str(e)
            print("Error getting checklists: {}".format(str(e)))
            return []
    
    def get_checklist_items(self, checklist_id, limit=10000):
        """Get checklist items for a specific checklist
        
        Args:
            checklist_id: ID of the checklist to fetch items from
            limit: Maximum number of items to fetch. Use 0 for no limit.
        """
        if not self.idToken or not self.current_project:
            self.last_error = "Not authenticated or no project selected"
            return []
        
        try:
            # Create a query object and encode it
            query = {
                "key": "object",
                "sort": {"field": "status", "descending": False},
                "page": {"skip": 0, "limit": limit if limit > 0 else 100000},  # Use large number for no limit
                "filter": {"checklist": checklist_id},
                "timeZone": "Europe/Stockholm",
                "format": "json",
                "filename": ""
            }
            
            # Encode the query as a base64 string
            encoded_query = base64.b64encode(json.dumps(query))
            
            url = "{}/project-{}/api/v1/checklists/export/json/?query={}".format(
                self.base_url, self.current_project, encoded_query
            )
            
            req = urllib2.Request(url)
            req.add_header('Authorization', 'Bearer {}'.format(self.idToken))
            req.add_header('Accept', '*/*')
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read())
            
            # Decode UTF-8 strings in the response
            result = self._decode_utf8(result)
            
            return result.get('data', [])
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Authentication failed. Please log in again."
                
            self.last_error = error_message
            print("Error getting checklist items: {}".format(error_message))
            return []
        except Exception as e:
            self.last_error = str(e)
            print("Error getting checklist items: {}".format(str(e)))
            return [] 