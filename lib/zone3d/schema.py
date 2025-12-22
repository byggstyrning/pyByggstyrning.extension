# -*- coding: utf-8 -*-
"""Schema for storing 3D Zone configurations using extensible storage."""

import System
from pyrevit import script
from extensible_storage import BaseSchema, simple_field, ES

# Initialize logger
logger = script.get_logger()

class Zone3DConfigSchema(BaseSchema):
    """Schema for storing 3D Zone parameter mapping configurations using pickle serialization"""
     
    guid = "52fc2611-774e-45e3-9f3d-c57fa3856760"
    
    @simple_field(value_type="string")
    def schema_version():
        """Current schema version."""
        return "0.1"
    
    @simple_field(value_type="string")
    def pickled_configs():
        """Base64 encoded pickle of all zone mapping configurations"""







