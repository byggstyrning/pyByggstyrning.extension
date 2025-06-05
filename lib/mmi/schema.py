"""Module for defining the MMI parameter mapping schema."""

import System
from pyrevit import script
from extensible_storage import BaseSchema, simple_field, ES

# Initialize logger
logger = script.get_logger()

class MMIParameterSchema(BaseSchema):
    """Schema for storing MMI parameter mapping preferences and settings."""
    
    # Schema identification and access levels
    guid = "8844cb2d-4234-4bf0-8361-b3da4d64234d"  # Updated GUID for new version
    vendor = "pyByggstyrning"
    application = "MMIParameterMapping"
    read_access_level = ES.AccessLevel.Public
    write_access_level = ES.AccessLevel.Public
    
    @simple_field(value_type="string")
    def schema_version():
        """Current schema version."""
        return "1.1"
    
    @simple_field(value_type="string")
    def mmi_parameter_name():
        """The selected MMI parameter name."""
        return None
    
    @simple_field(value_type="string")
    def last_used_date():
        """Track when the parameter was last used."""
        return None
    
    @simple_field(value_type="boolean")
    def is_validated():
        """Whether the parameter has been validated."""
        return False
    
    @simple_field(value_type="string")
    def pm_id():
        """The PM-ID associated with this element."""
        return None

    @simple_field(value_type="boolean")
    def validate_mmi():
        """Whether to validate MMI on change."""
        return False
        
    @simple_field(value_type="boolean")
    def pin_elements():
        """Whether to pin elements >= 400."""
        return False
        
    @simple_field(value_type="boolean")
    def warn_on_move():
        """Whether to warn when moving elements > 400."""
        return False
        
    @simple_field(value_type="boolean")
    def check_mmi_after_sync():
        """Whether to check MMI on user-owned elements after sync."""
        return False