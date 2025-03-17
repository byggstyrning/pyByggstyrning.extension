import json
import urllib2
import base64
from collections import namedtuple

# StreamBIM API client
class StreamBIMClient:
    def __init__(self, base_url="https://app.streambim.com"):
        self.base_url = base_url
        self.token = None
        self.projects = []
        self.current_project = None
        self.last_error = None
        
    def login(self, username, password):
        """Login to StreamBIM and get authentication token"""
        try:
            url = "{}/mgw/api/v2/login".format(self.base_url)
            data = json.dumps({
                "username": username,
                "password": password
            })
            
            req = urllib2.Request(url)
            req.add_header('Content-Type', 'application/json')
            
            response = urllib2.urlopen(req, data)
            result = json.loads(response.read())
            
            if 'token' in result:
                self.token = result['token']
                self.last_error = None
                return True
            
            self.last_error = "No token in response"
            return False
        except urllib2.HTTPError as e:
            error_message = "HTTP Error: {} - {}".format(e.code, e.reason)
            if e.code == 401:
                error_message = "Invalid username or password"
            elif e.code == 404:
                error_message = "Server URL not found"
                
            self.last_error = error_message
            print("Login error: {}".format(error_message))
            return False
        except Exception as e:
            self.last_error = str(e)
            print("Login error: {}".format(str(e)))
            return False
    
    def get_projects(self):
        """Get list of available projects"""
        if not self.token:
            self.last_error = "Not authenticated"
            return []
            
        try:
            url = "{}/mgw/api/v2/projects".format(self.base_url)
            req = urllib2.Request(url)
            req.add_header('Authorization', 'Bearer {}'.format(self.token))
            req.add_header('Accept', 'application/json')
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read())
            
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
        if not self.token or not self.current_project:
            self.last_error = "Not authenticated or no project selected"
            return []
            
        try:
            url = "{}/project-{}/api/v1/v2/checklists?filter[isDraft]=false".format(
                self.base_url, self.current_project
            )
            
            req = urllib2.Request(url)
            req.add_header('Authorization', 'Bearer {}'.format(self.token))
            req.add_header('Accept', 'application/vnd.api+json')
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read())
            
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
    
    def get_checklist_items(self, checklist_id, limit=1000):
        """Get checklist items for a specific checklist"""
        if not self.token or not self.current_project:
            self.last_error = "Not authenticated or no project selected"
            return []
        
        try:
            # Create a query object and encode it
            query = {
                "key": "object",
                "sort": {"field": "title", "descending": False},
                "page": {"skip": 0, "limit": limit},
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
            req.add_header('Authorization', 'Bearer {}'.format(self.token))
            req.add_header('Accept', '*/*')
            
            response = urllib2.urlopen(req)
            result = json.loads(response.read())
            
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