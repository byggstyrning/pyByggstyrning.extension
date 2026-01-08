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
        data_storages = FilteredElementCollector(doc)\
            .OfClass(ExtensibleStorage.DataStorage)\
            .ToElements()
        
        # Look for our storage with the schema
        for ds in data_storages:
            try:
                # Check if this storage has our schema
                entity = ds.GetEntity(StreamBIMSettingsSchema.schema)
                if entity.IsValid():
                    return ds
            except Exception as e:
                continue
        
        # If not found, create a new one
        with revit.Transaction("Create StreamBIM Settings Storage", doc):
            new_storage = ExtensibleStorage.DataStorage.Create(doc)
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
            return []
            
        pickled_configs = schema.get("pickled_configs")
        if not pickled_configs:
            return []
            
        # Decode and unpickle
        try:
            decoded_data = base64.b64decode(pickled_configs)
            configs = pickle.loads(decoded_data)
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
                            return project_id
            except Exception as e:
                continue
        
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
        self.mfa_session = None  # Store MFA session token
        
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
            pass
    
    def clear_tokens(self):
        """Clear saved tokens."""
        self.idToken = None
        self.accessToken = None
        self.username = None
        try:
            if os.path.exists(self.token_file):
                os.remove(self.token_file)
        except Exception as e:
            pass
        
    def login(self, username, password):
        """Login to StreamBIM and get authentication token.
        
        Returns:
            dict with keys:
            - 'success': bool - True if login successful
            - 'requires_mfa': bool - True if MFA challenge is required
            - 'session': str - MFA session token if requires_mfa is True
            - 'result': str - Response result field
        """
        try:
            # Use new auth endpoint
            url = "{}/auth/v1/login".format(self.base_url)
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

            # Check for MFA challenge
            if result.get('result') == 'CHALLENGE_REQUESTED' and result.get('session'):
                self.last_error = None
                return {
                    'success': False,
                    'requires_mfa': True,
                    'session': result.get('session'),
                    'result': result.get('result')
                }
            
            # Check for successful login with tokens
            if 'idToken' in result or 'accessToken' in result:
                self.accessToken = result.get('accessToken')
                self.idToken = result.get('idToken')
                self.username = username  # Store the username
                self.save_tokens()
                self.last_error = None
                return {
                    'success': True,
                    'requires_mfa': False,
                    'session': None,
                    'result': result.get('result', 'SUCCESS')
                }
            
            self.last_error = "No token in response"
            return {
                'success': False,
                'requires_mfa': False,
                'session': None,
                'result': result.get('result', 'UNKNOWN')
            }
        except urllib2.HTTPError as e:
            error_body = ""
            try:
                error_body = e.read()
                # Try to parse error response for MFA challenge
                error_data = json.loads(error_body)
                if error_data.get('result') == 'CHALLENGE_REQUESTED' and error_data.get('session'):
                    self.last_error = None
                    return {
                        'success': False,
                        'requires_mfa': True,
                        'session': error_data.get('session'),
                        'result': error_data.get('result')
                    }
            except:
                pass
            
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Invalid username or password"
            elif e.code == 404:
                error_message = "Server URL not found"
                
            self.last_error = error_message
            logger.error("Login error: {}".format(error_message))
            return {
                'success': False,
                'requires_mfa': False,
                'session': None,
                'result': 'ERROR'
            }
        except Exception as e:
            self.last_error = str(e)
            logger.error("Login error: {}".format(str(e)))
            return {
                'success': False,
                'requires_mfa': False,
                'session': None,
                'result': 'ERROR'
            }
    
    def verify_mfa(self, username, session, code):
        """Verify MFA code and complete login.
        
        Args:
            username: Username for the account
            session: MFA session token from login response
            code: MFA verification code
            
        Returns:
            dict with keys:
            - 'success': bool - True if verification successful
            - 'accessToken': str - Access token if successful
            - 'idToken': str - ID token if successful
        """
        try:
            url = "{}/auth/v1/mfa/verify".format(self.base_url)
            data = json.dumps({
                "username": username,
                "session": session,
                "code": code
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

            if 'idToken' in result and 'accessToken' in result:
                self.accessToken = result['accessToken']
                self.idToken = result['idToken']
                self.username = username  # Store the username
                self.save_tokens()
                self.last_error = None
                return {
                    'success': True,
                    'accessToken': result['accessToken'],
                    'idToken': result['idToken']
                }
            
            self.last_error = "MFA verification failed: No tokens in response"
            return {
                'success': False,
                'accessToken': None,
                'idToken': None
            }
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                try:
                    error_data = json.loads(e.read())
                    error_message = error_data.get('message', 'Invalid MFA code')
                except:
                    error_message = "Invalid MFA code"
            elif e.code == 404:
                error_message = "MFA verification endpoint not found"
                
            self.last_error = error_message
            logger.error("MFA verification error: {}".format(error_message))
            return {
                'success': False,
                'accessToken': None,
                'idToken': None
            }
        except Exception as e:
            self.last_error = str(e)
            logger.error("MFA verification error: {}".format(str(e)))
            return {
                'success': False,
                'accessToken': None,
                'idToken': None
            }
    
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
            except (UnicodeError, AttributeError):
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
            return []
        except Exception as e:
            self.last_error = str(e)
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
            return []
        except Exception as e:
            self.last_error = str(e)
            return []
    
    def get_checklist_items(self, checklist_id, checklist_item=None, limit=10000):
        """Get checklist items for a specific checklist
        
        Args:
            checklist_id: ID of the checklist to fetch items from
            checklist_item: Optional specific checklist item/property to filter by
            limit: Maximum number of items to fetch. Use 0 for no limit.
        """
        if not self.idToken or not self.current_project:
            self.last_error = "Not authenticated or no project selected"
            return []
        
        try:
            # Create base query object
            query = {
                "key": "object",
                "sort": {"field": "title", "descending": False},
                "page": {"skip": 0, "limit": limit if limit > 0 else 100000},
                "filter": {
                    "checklist": checklist_id
                },
                "timeZone": "Europe/Stockholm",
                "format": "json",
                "filename": ""
            }

            # Add checklist item filter if specified
            if checklist_item:
                # Convert checklist_item to UTF-8 if it's not already
                if isinstance(checklist_item, str):
                    checklist_item = checklist_item.decode('utf-8') if not isinstance(checklist_item, unicode) else checklist_item
                
                # Add to filter but don't override existing checklist filter
                query["filter"].update({
                    "properties": {checklist_item: {"$exists": True}}
                })

            # Convert the entire query to UTF-8 JSON
            query_json = json.dumps(query, ensure_ascii=False).encode('utf-8')
            
            # Encode the query as a base64 string
            encoded_query = base64.b64encode(query_json)
            
            url = "{}/project-{}/api/v1/checklists/export/json/?query={}".format(
                self.base_url, self.current_project, encoded_query
            )
            
            req = urllib2.Request(url)
            req.add_header('Authorization', 'Bearer {}'.format(self.idToken))
            req.add_header('Accept', '*/*')
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read().decode('utf-8'))
            
            # Decode UTF-8 strings in the response
            result = self._decode_utf8(result)

            return result.get('data', [])
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Authentication failed. Please log in again."
            elif e.code == 500:
                error_message = "Server error. The query might be malformed: {}".format(e.read())
                
            self.last_error = error_message
            return []
        except UnicodeError as e:
            self.last_error = "Unicode error: {}".format(str(e))
            return []
        except Exception as e:
            self.last_error = str(e)
            return []
    
    def create_ifc_search(self, checklist_id, building_id, checklist_value):
        """Create an IFC search for a grouped checklist value.
        
        Args:
            checklist_id: ID of the checklist
            building_id: ID of the building
            checklist_value: The group key (checklistValue) to search for
            
        Returns:
            searchId string if successful, None otherwise
        """
        if not self.idToken or not self.current_project:
            self.last_error = "Not authenticated or no project selected"
            return None
        
        try:
            url = "{}/project-{}/api/v1/ifc-searches".format(
                self.base_url, self.current_project
            )
            
            # Build rules with only checklistValue (no @kind or @Document Id needed)
            rules = [[{
                "propKey": "checklistValue",
                "propValue": checklist_value,
                "buildingId": building_id,
                "checklistId": checklist_id
            }]]
            
            data = json.dumps({"rules": rules}, ensure_ascii=False).encode('utf-8')
            
            req = urllib2.Request(url, data=data)
            req.add_header('Authorization', 'Bearer {}'.format(self.idToken))
            req.add_header('Content-Type', 'application/json')
            req.add_header('Accept', '*/*')
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read().decode('utf-8'))
            
            # Decode UTF-8 strings in the response
            result = self._decode_utf8(result)
            
            search_id = result.get('searchId')
            if search_id:
                return str(search_id)
            else:
                self.last_error = "No searchId in response"
                return None
                
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Authentication failed. Please log in again."
            elif e.code == 500:
                error_message = "Server error: {}".format(e.read())
                
            self.last_error = error_message
            logger.error("Error creating IFC search: {}".format(error_message))
            return None
        except Exception as e:
            self.last_error = str(e)
            logger.error("Error creating IFC search: {}".format(str(e)))
            return None
    
    def get_ifc_object_refs(self, search_id, limit=500):
        """Get IFC object references for a search ID.
        
        Args:
            search_id: The search ID from create_ifc_search
            limit: Maximum number of results to fetch (default 500)
            
        Returns:
            List of IFC object ref data dictionaries
        """
        if not self.idToken or not self.current_project:
            self.last_error = "Not authenticated or no project selected"
            return []
        
        try:
            url = "{}/project-{}/api/v1/v2/ifc-object-refs-sets?page[limit]={}&searchId={}".format(
                self.base_url, self.current_project, limit, search_id
            )
            
            req = urllib2.Request(url)
            req.add_header('Authorization', 'Bearer {}'.format(self.idToken))
            req.add_header('Accept', 'application/vnd.api+json')
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read().decode('utf-8'))
            
            # Decode UTF-8 strings in the response
            result = self._decode_utf8(result)
            
            data = result.get('data', [])
            return data
            
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Authentication failed. Please log in again."
                
            self.last_error = error_message
            logger.error("Error getting IFC object refs: {}".format(error_message))
            return []
        except Exception as e:
            self.last_error = str(e)
            logger.error("Error getting IFC object refs: {}".format(str(e)))
            return []
    
    def resolve_group_key_to_ifc_guids(self, checklist_id, building_id, group_key):
        """Resolve a grouped checklist group key to a list of IFC GUIDs.
        
        This method caches results per (checklist_id, building_id, group_key) to avoid
        redundant API calls during a single import run.
        
        Args:
            checklist_id: ID of the checklist
            building_id: ID of the building
            group_key: The group key (object value from exported checklist items)
            
        Returns:
            List of IFC GUID strings (matching Revit IfcGUID parameter values)
        """
        # Initialize cache if it doesn't exist
        if not hasattr(self, '_group_key_cache'):
            self._group_key_cache = {}
        
        # Check cache first
        cache_key = (checklist_id, building_id, group_key)
        if cache_key in self._group_key_cache:
            return self._group_key_cache[cache_key]
        
        # Resolve via API
        search_id = self.create_ifc_search(checklist_id, building_id, group_key)
        
        if not search_id:
            self._group_key_cache[cache_key] = []
            return []
        
        # Get IFC object refs (use high limit to get all results)
        object_refs = self.get_ifc_object_refs(search_id, limit=10000)
        
        # Extract IDs (these match Revit IfcGUID parameter values)
        ifc_guids = []
        for ref in object_refs:
            ref_id = ref.get('id')
            if ref_id:
                ifc_guids.append(ref_id)
        
        # Cache the result
        self._group_key_cache[cache_key] = ifc_guids
        
        return ifc_guids 