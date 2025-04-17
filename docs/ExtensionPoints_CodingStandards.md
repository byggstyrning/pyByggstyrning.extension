# Extension Points Analysis and Coding Standards
*for pyByggstyrning Extension*

## Table of Contents
1. [Introduction](#introduction)
2. [Extension Points](#extension-points)
3. [Architectural Patterns](#architectural-patterns)
4. [Coding Standards](#coding-standards)
5. [Error Handling and Logging](#error-handling-and-logging)
6. [Feature Development Template](#feature-development-template)
7. [Refactoring Recommendations](#refactoring-recommendations)

## Introduction

This document serves as a guide for maintaining and extending the pyByggstyrning extension. It identifies extension points for adding new features, describes the architectural patterns used throughout the codebase, establishes coding standards for consistency, and provides templates for feature development.

## Extension Points

The following components have been identified as key extension points for adding new functionality to the pyByggstyrning extension:

### 1. Extensible Storage Schema System

The BaseSchema implementation provides a powerful extension point for storing custom data:

```python
# Example of creating a new schema
class MyCustomSchema(BaseSchema):
    """Schema for storing custom data."""
    
    guid = "00000000-0000-0000-0000-000000000000"  # Replace with a unique GUID
    vendor = "pyByggstyrning"
    application = "MyCustomFeature"
    read_access_level = ES.AccessLevel.Public
    write_access_level = ES.AccessLevel.Public
    
    @simple_field(value_type="string")
    def schema_version():
        """Current schema version."""
        return "1.0"
        
    @simple_field(value_type="string")
    def my_custom_field():
        """Custom data field."""
        return None
```

### 2. UI Panel Integration

New functionality can be added through the pyRevit panel system by:

1. Creating a new pushbutton script in the appropriate panel directory
2. Setting standard attributes to integrate with the UI system

```python
__title__ = "New Feature"
__author__ = "Your Name"
__doc__ = """Tooltip description for the new feature button."""
__context__ = 'Selection'  # Activate only when elements are selected
```

### 3. Event Subscription System

Custom event handlers can be added in the startup.py file to respond to Revit events:

```python
# Example event subscription in startup.py
HOST_APP.app.DocumentOpening += \
    EventHandler[Events.DocumentOpeningEventArgs](
        my_custom_doc_opening_handler
    )
```

### 4. Parameter Management

The parameter access and modification system can be extended with new parameter types and handling logic:

```python
# Example of extending parameter handling
def handle_custom_parameter(element, param_name, value):
    """Custom parameter handling logic."""
    param = element.LookupParameter(param_name)
    if not param:
        return False
        
    # Custom handling logic here
    return True
```

### 5. StreamBIM Integration API

The StreamBIM integration provides extension points for connecting with external systems:

```python
# Example of extending StreamBIM integration
def import_custom_data_from_streambim(project_id, element_id):
    """Import custom data from StreamBIM."""
    # Custom integration logic
    pass
```

### 6. Configuration System

The extension can be configured through a standardized configuration system:

```python
# Example of accessing configuration
def get_custom_config(config_key, default_value=None):
    """Get custom configuration value."""
    config = get_extension_config()
    return config.get(config_key, default_value)
```

## Architectural Patterns

The pyByggstyrning extension employs several architectural patterns:

### 1. Model-View-Controller (MVC) Pattern

- **Model**: Data structures like MMIParameterSchema represent application data
- **View**: WPF interfaces in panel scripts handle visual representation
- **Controller**: Script logic in pushbutton scripts manage user interactions

### 2. Command Pattern

UI buttons implement the command pattern:
- Each button script encapsulates a specific operation
- Commands have a clear single responsibility
- Transaction management handles operation atomicity

### 3. Observer Pattern

Event subscriptions follow the observer pattern:
- Revit events trigger registered handlers
- Event handlers react to document/application changes
- Loose coupling between event source and handlers

### 4. Factory Pattern

Schema and field creation uses factory patterns:
- Field decorators act as factories for field descriptors
- Schema metaclass acts as a factory for schema instances
- Utility methods provide factory functions for component creation

### 5. Descriptor Pattern

Field access uses the descriptor pattern:
- FieldDescriptor implements the descriptor protocol
- Property-like access to schema fields
- Type validation and conversion handled transparently

### 6. Context Manager Pattern

Transaction management uses context managers:
- with statements ensure proper transaction lifecycle
- Automatic resource cleanup on exceptions
- Simplified transaction boundary definition

## Coding Standards

To maintain consistency across the codebase, follow these standards:

### 1. Naming Conventions

- **Classes**: Use PascalCase for class names (e.g., `BaseSchema`, `MMIParameterSchema`)
- **Functions/Methods**: Use snake_case for function and method names (e.g., `get_field`, `update_schema_entities`)
- **Variables**: Use snake_case for variable names (e.g., `field_descriptor`, `schema_builder`)
- **Constants**: Use UPPER_SNAKE_CASE for constants (e.g., `DEFAULT_SCHEMA_VERSION`, `MMI_PARAMETER_NAME`)
- **Private Members**: Prefix with underscore for private methods/variables (e.g., `_wrapped`, `_get_field`)

### 2. Documentation

- Use Google-style docstrings for all classes, methods, and functions:

```python
def function_name(param1, param2):
    """Short description of function.
    
    More detailed description of function.
    
    Args:
        param1 (type): Description of param1.
        param2 (type): Description of param2.
        
    Returns:
        type: Description of return value.
        
    Raises:
        ExceptionType: When and why this exception is raised.
    """
    pass
```

- Include examples for complex functionality
- Keep docstrings updated when implementation changes
- Document all public APIs thoroughly

### 3. Code Structure

- Limit function length to maximum 30-40 lines
- Use clear responsibility separation between modules
- Group related functionality in classes or modules
- Use inheritance judiciously, favor composition where appropriate
- Keep classes focused on a single responsibility

### 4. Import Conventions

- Group imports in the following order:
  1. Standard library imports
  2. Third-party imports (like System imports)
  3. pyRevit and Revit API imports
  4. Local module imports
- Use explicit imports rather than wildcard imports
- Import only what is needed

```python
# Example import structure
import sys
from datetime import datetime

import System
from System.Collections.Generic import List

from pyrevit import DB, forms, revit, script

from extensible_storage import BaseSchema, simple_field
```

### 5. Coding Style

- Use 4 spaces for indentation (not tabs)
- Maximum line length: 79 characters
- Use blank lines to separate logical sections of code
- Use consistent spacing around operators
- Follow PEP 8 guidelines for Python code style

## Error Handling and Logging

### 1. Exception Handling

- Use specific exception types rather than catching all exceptions
- Handle exceptions at appropriate levels:
  - Wrap low-level exceptions in domain-specific exceptions
  - Log exceptions with appropriate context
  - Present user-friendly error messages
- Use context managers for resource cleanup

```python
try:
    # Operation that might fail
    result = perform_operation()
except SpecificException as ex:
    # Log the exception
    logger.error("Operation failed: %s", ex)
    # Provide user feedback
    forms.alert("The operation could not be completed.")
    # Return appropriate fallback or error state
    return None
```

### 2. Logging System

- Use the script logger for consistent log formatting:

```python
logger = script.get_logger()
logger.debug("Detailed debugging information")
logger.info("General information about operation progress")
logger.warning("Warning about potential issues")
logger.error("Error information when operations fail")
```

- Include appropriate context in log messages:
  - Operation being performed
  - Relevant identifiers (element IDs, GUIDs, etc.)
  - Parameter values where appropriate
- Use log levels appropriately:
  - DEBUG: Detailed troubleshooting information
  - INFO: Confirmation that things are working as expected
  - WARNING: Indication of potential issues
  - ERROR: Error conditions that should be investigated

### 3. User Feedback

- Provide clear feedback for user-initiated operations:
  - Success messages for completed operations
  - Error dialogs for failed operations
  - Progress indicators for long-running operations
- Use appropriate dialog types:
  - Simple alerts for informational messages
  - Task dialogs for more complex interactions
  - Custom forms for complex feedback

```python
# Success feedback
forms.alert("Operation completed successfully.", title="Success")

# Error feedback
forms.alert("The operation failed: {}".format(error_msg), 
            title="Error", exitscript=True)

# Progress feedback
with forms.ProgressBar(title="Processing Elements") as pb:
    for i, element in enumerate(elements):
        pb.update_progress(i, len(elements))
        process_element(element)
```

## Feature Development Template

### 1. Feature Development Process

Follow this process for adding new features:

1. **Analysis**: Understand the requirements and how they fit into existing architecture
2. **Design**: Plan the implementation with consideration for extension points
3. **Implementation**: Write the code following the coding standards
4. **Testing**: Verify the feature works correctly
5. **Documentation**: Update documentation to include the new feature
6. **Review**: Have the feature reviewed by team members

### 2. Feature Implementation Template

Use this template for implementing new features:

```python
"""Module for implementing [Feature Name].

This module provides functionality for [brief description].
"""

# Standard imports
import sys
from datetime import datetime

# Third-party imports
import System
from System.Collections.Generic import List

# pyRevit imports
from pyrevit import DB, forms, revit, script

# Local imports
from extensible_storage import BaseSchema, simple_field, ES

# Initialize logger
logger = script.get_logger()

# Constants
FEATURE_NAME = "MyFeature"
DEFAULT_VALUE = "Default"

# Feature schema definition
class FeatureSchema(BaseSchema):
    """Schema for storing feature-specific data."""
    
    guid = "00000000-0000-0000-0000-000000000000"  # Generate a unique GUID
    vendor = "pyByggstyrning"
    application = FEATURE_NAME
    read_access_level = ES.AccessLevel.Public
    write_access_level = ES.AccessLevel.Public
    
    @simple_field(value_type="string")
    def schema_version():
        """Current schema version."""
        return "1.0"
    
    # Add feature-specific fields here

# Utility functions
def feature_specific_function(param):
    """Performs a feature-specific operation.
    
    Args:
        param: The parameter to process.
        
    Returns:
        The result of the operation.
    """
    # Implementation here
    return result

# Main feature class
class FeatureImplementation:
    """Main implementation of the feature."""
    
    def __init__(self, doc):
        """Initialize with the active document.
        
        Args:
            doc: The Revit document.
        """
        self.doc = doc
        self.logger = logger
    
    def execute(self):
        """Execute the main feature functionality."""
        try:
            # Implementation here
            self._perform_operation()
            return True
        except Exception as ex:
            self.logger.error("Feature execution failed: %s", ex)
            forms.alert("The operation failed: {}".format(ex))
            return False
    
    def _perform_operation(self):
        """Internal helper method."""
        pass

# Script entry point (for pushbutton scripts)
if __name__ == "__main__":
    # Get the current document
    doc = revit.doc
    
    # Initialize the feature
    feature = FeatureImplementation(doc)
    
    # Execute with proper transaction management
    with revit.Transaction("Execute Feature", doc):
        success = feature.execute()
    
    # Provide feedback
    if success:
        forms.alert("Feature executed successfully.", title="Success")
```

### 3. UI Integration Template

For features requiring UI interaction, use this template:

```python
"""UI implementation for [Feature Name].

This module provides the user interface for [brief description].
"""

# Standard imports
import System
from System import Windows

# pyRevit imports
from pyrevit import forms, script

# Initialize logger
logger = script.get_logger()

# Define WPF form
class FeatureForm(forms.WPFWindow):
    """WPF form for the feature."""
    
    def __init__(self, xaml_file_name):
        """Initialize the WPF form.
        
        Args:
            xaml_file_name: XAML file defining the UI.
        """
        forms.WPFWindow.__init__(self, xaml_file_name)
        self.setup_form()
    
    def setup_form(self):
        """Set up form controls and event handlers."""
        # Connect event handlers
        self.submit_button.Click += self.handle_submit
        
        # Initialize UI state
        self.populate_data()
    
    def populate_data(self):
        """Populate the form with data."""
        pass
    
    def handle_submit(self, sender, args):
        """Handle submit button click."""
        try:
            # Process form data
            self.process_data()
            self.Close()
        except Exception as ex:
            logger.error("Form submission failed: %s", ex)
            forms.alert("The operation failed: {}".format(ex))
    
    def process_data(self):
        """Process the form data."""
        pass

# Create and show form
if __name__ == "__main__":
    form = FeatureForm("FeatureForm.xaml")
    form.ShowDialog()
```

## Refactoring Recommendations

Based on the analysis of the codebase, the following areas are recommended for refactoring to improve extensibility:

### 1. Schema Version Management

Create a more robust schema versioning system that can:
- Automatically migrate schema data between versions
- Track schema compatibility
- Provide version upgrade paths

### 2. UI Component Library

Develop a reusable UI component library that includes:
- Standard dialog templates
- Common UI controls for parameter selection
- Consistent styling and interaction patterns

### 3. Parameter Management API

Create a unified parameter management API that abstracts:
- Parameter discovery and categorization
- Type-safe parameter access and modification
- Parameter validation and mapping

### 4. Enhanced Testing Framework

Implement a comprehensive testing framework with:
- Unit tests for core functionality
- Integration tests for Revit API interactions
- Mock objects for testing without Revit

### 5. Documentation Generation

Add automatic documentation generation from:
- Code docstrings
- Schema definitions
- UI component specifications

### 6. Configuration Management

Enhance the configuration system with:
- Environment-specific settings
- User preference management
- Feature flags for experimental features

### 7. Dependency Injection

Introduce a dependency injection pattern to:
- Improve testability of components
- Reduce direct dependencies between modules
- Enable easier component replacement

These refactoring recommendations aim to enhance the maintainability, extensibility, and testability of the pyByggstyrning extension while preserving its core functionality and architectural patterns. 