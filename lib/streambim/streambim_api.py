import json
import urllib2
import base64
from collections import namedtuple
import os


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
                "sort": {"field": "title", "descending": False},
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