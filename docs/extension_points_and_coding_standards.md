# pyByggstyrning Extension Points Analysis and Coding Standards

## Table of Contents

1. [Introduction](#introduction)
2. [Extension Architecture Overview](#extension-architecture-overview)
3. [Extension Points](#extension-points)
4. [Coding Standards](#coding-standards)
5. [Error Handling Patterns](#error-handling-patterns)
6. [Feature Implementation Template](#feature-implementation-template)
7. [Refactoring Opportunities](#refactoring-opportunities)
8. [Appendix: Sample Implementation](#appendix-sample-implementation)

## Introduction

This document provides a comprehensive analysis of the pyByggstyrning extension architecture, identifies potential extension points for new features, and documents coding standards to maintain consistency across the codebase. It serves as a guide for developers who want to extend or modify the extension.

The document is based on an extensive analysis of the existing codebase, including module structure, UI components, Revit API integration points, and data models. It provides recommendations for consistent development practices and identifies areas for potential refactoring to improve extensibility.

## Extension Architecture Overview

The pyByggstyrning extension follows a modular architecture organized around the following core components:

- **Core Libraries**: Foundational modules for Revit API interaction, including the extensible storage framework
- **UI Components**: Ribbon panels, buttons, and dialog implementations
- **Domain Models**: Data models and business logic for MMI and parameter management
- **Integration Points**: Services and handlers for interacting with Revit and external systems

The extension uses several architectural patterns:

- **Command Pattern**: For encapsulating operations triggered by UI elements
- **Observer Pattern**: For event handling and responding to Revit document changes
- **Factory Pattern**: For creating schema and field instances
- **Context Manager Pattern**: For transaction management
- **Descriptor Pattern**: For Python property-like field access

## Extension Points

The pyByggstyrning extension has been designed with several key extension points that can be leveraged to add new functionality without modifying the core codebase. These extension points are organized by architectural layer.

### 1. Schema Extension Points

The Extensible Storage framework provides a robust base for creating custom data schemas:

#### 1.1 Custom Schema Classes

New schemas can be created by extending the `BaseSchema` class:

```python
class CustomParameterSchema(BaseSchema):
    """Schema for custom parameter mapping."""
  
    guid = "YOUR-UNIQUE-GUID-HERE"
    vendor = "YourVendorName"
    application = "YourApplicationName"
    read_access_level = ES.AccessLevel.Public
    write_access_level = ES.AccessLevel.Public
  
    @simple_field(value_type="string")
    def schema_version():
        """Current schema version."""
        return "1.0"
      
    @simple_field(value_type="string")
    def custom_parameter_name():
        """The custom parameter name."""
        return None
```

#### 1.2 Field Type Extensions

The field system allows for adding custom field types:

- Extension point: Add new field factory functions similar to `simple_field`, `array_field`, and `map_field`
- Use the `schema_field` decorator to create specialized field types
- Implement custom field validation logic

### 2. UI Extension Points

#### 2.1 Panel Extensions

New functionality can be added to existing panels or new panels can be created:

- Each panel is contained in a directory structure within `pyBS.tab/<PanelName>.panel/`
- New buttons can be added by creating a new `.pushbutton` directory with a `script.py` file
- Existing panels can be extended by adding new push buttons or pull-down buttons

#### 2.2 Custom Dialog Components

The extension uses WPF for complex UI interactions and provides several patterns for creating reusable dialogs:

- Inherit from `WPFWindow` for creating custom dialogs
- Use data binding with observable collections for dynamic UI updates
- Implement `INotifyPropertyChanged` for two-way data binding
- Create custom WPF controls for specialized functionality

Example extension point for a custom dialog:

```python
class CustomParameterDialog(WPFWindow):
    """Custom dialog for parameter selection."""
  
    def __init__(self):
        # Load XAML file (defined in the same directory)
        xaml_file = os.path.join(os.path.dirname(__file__), "CustomDialog.xaml")
        wpf.LoadComponent(self, xaml_file)
      
        # Initialize external event handlers if needed
        self.apply_handler = CustomEventHandler(self)
        self.ext_event = UI.ExternalEvent.Create(self.apply_handler)
      
        # Initialize UI components and data
        self.setup_ui_components()
```

#### 2.3 Custom Event Handlers

External event handlers provide a thread-safe way to modify the Revit database from UI actions:

```python
class CustomEventHandler(UI.IExternalEventHandler):
    """External event handler for custom operations."""
  
    def __init__(self, ui_reference):
        self.ui_reference = ui_reference
      
    def Execute(self, uiapp):
        try:
            doc = uiapp.ActiveUIDocument.Document
            # Implement custom logic here
            with revit.Transaction("Custom Operation", doc):
                # Modify Revit database
                pass
        except Exception as e:
            self.log_exception()
          
    def GetName(self):
        return "Custom Event Handler"
      
    def log_exception(self):
        import traceback
        logger.error("Error in CustomEventHandler: {}".format(traceback.format_exc()))
```

### 3. Parameter Management Extension Points

#### 3.1 Parameter Discovery Extensions

The parameter management system can be extended through:

- Custom parameter filters and categorization
- Parameter mapping and transformation functions
- Value extraction and conversion utilities

#### 3.2 Parameter Visualization Extensions

The ColorElements tool provides extension points for custom visualization:

- Custom color schemes and pattern generators
- Parameter value analyzers and distributors
- Specialized view filters and overrides

### 4. Integration Extension Points

#### 4.1 External API Integration

New external service integrations can be added:

- HTTP API clients for web services
- Authentication and token management
- Data mapping between external systems and Revit

#### 4.2 Revit API Integration

Custom Revit API interactions can be implemented through:

- Transaction wrappers with retry logic
- Element collectors and filters
- Parameter access and modification utilities
- View and graphic management tools

### 5. Configuration Extension Points

#### 5.1 User Preferences

Extend configuration options:

- Create new schema classes for storing user preferences
- Add UI components for preference management
- Implement configuration validators and migrators

## Coding Standards

To maintain consistency across the pyByggstyrning extension, developers should adhere to the following coding standards based on the analysis of the existing codebase.

### 1. File Organization

#### 1.1 Directory Structure

Follow the established directory structure pattern:

```
pyByggstyrning.extension/
├── lib/                             # Core libraries
│   ├── extensible_storage/          # Extensible storage framework
│   ├── mmi_schema.py                # Schema definitions
│   └── utils.py                     # Utility functions
├── pyBS.tab/                        # UI components
│   ├── MMI.panel/                   # MMI-related functionality
│   │   ├── 200.pushbutton/          # Button for setting MMI value 200
│   │   │   └── script.py            # Button implementation
│   │   ├── View.panel/              # View-related functionality
│   │   │   ├── ColorElements.pushbutton/ # Button for colorizing elements
│   │   │   │   ├── script.py            # Main script
│   │   │   │   └── ColorElements.xaml   # WPF UI definition
│   │   ├── StreamBIM.panel/         # StreamBIM integration
│   └── dev.panel/                   # Development tools
├── coloringschemas/                 # Storage for coloring schemas
└── docs/                            # Documentation
```

#### 1.2 File Structure

Organize Python files consistently with the following sections:

1. File header with encoding and documentation
2. Imports (grouped by standard library, .NET, then Revit)
3. Constants and configuration
4. Helper functions
5. Class definitions
6. Main execution code (at the end of the file)

```python
# -*- coding: utf-8 -*-
__title__ = "Feature Name"
__author__ = "Byggstyrning AB"
__doc__ = """Feature description and usage instructions.
More detailed information about the feature.
"""

# Standard library imports
import os
import sys
from collections import defaultdict

# .NET imports
import clr
from System import Object
from System.Collections.Generic import List

# Revit API imports
from pyrevit import HOST_APP, revit, DB, UI
from pyrevit.forms import WPFWindow

# Constants and configuration
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
DEFAULT_VALUES = {"param1": "value1", "param2": "value2"}

# Helper functions
def helper_function():
    """Helper function documentation."""
    pass

# Class definitions
class MainClass(object):
    """Main class documentation."""
    
    def __init__(self):
        """Initialize the class."""
        pass

# Main execution
if __name__ == '__main__':
    # Main code here
    pass
```

### 2. Naming Conventions

#### 2.1 File and Directory Names

- Use lowercase with underscores for Python files: `file_name.py`
- Use PascalCase for XAML files: `DialogName.xaml`
- Use descriptive names for pushbutton directories: `FeatureName.pushbutton`

#### 2.2 Class Names

- Use PascalCase for class names: `ClassName`
- Use descriptive names that indicate the class's purpose: `ParameterInfo`, `CategoryItem`

#### 2.3 Function and Method Names

- Use lowercase with underscores: `function_name`
- Use verbs to describe actions: `get_parameter_values`, `process_category_selection`
- Use self-explanatory names that describe what the function does

#### 2.4 Variable Names

- Use lowercase with underscores: `variable_name`
- Use descriptive names that indicate the variable's purpose: `parameter_info`, `category_list`
- Avoid single-letter variable names except for loop indices

#### 2.5 Constants

- Use uppercase with underscores: `CONSTANT_NAME`
- Group related constants together: `CAT_EXCLUDED`

### 3. Documentation

#### 3.1 File Documentation

Each script file should include:

- Encoding declaration: `# -*- coding: utf-8 -*-`
- Title: `__title__ = "Feature Name"`
- Author: `__author__ = "Byggstyrning AB"`
- Documentation: `__doc__ = """Description..."""`

#### 3.2 Function and Class Documentation

Use Google-style docstrings for all functions and classes:

```python
def function_name(param1, param2):
    """Short description of what the function does.
    
    Longer description with more details if needed.
    
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

#### 3.3 Code Comments

- Use inline comments sparingly and only when necessary to explain complex logic
- Keep comments up-to-date with code changes
- Use TODO comments to mark areas for future improvement: `# TODO: Implement feature X`

### 4. Code Style

#### 4.1 Imports

Organize imports in three groups, separated by blank lines:
1. Standard library imports
2. .NET imports
3. Revit API imports

```python
# Standard library imports
import os
import sys

# .NET imports
import clr
from System import Object

# Revit API imports
from pyrevit import revit, DB
```

#### 4.2 Indentation and Line Length

- Use 4 spaces for indentation (no tabs)
- Limit line length to 79-100 characters
- Use line continuation for long lines:

```python
long_variable_name = some_function_with_a_long_name(
    argument1,
    argument2,
    argument3
)
```

#### 4.3 White Space

- Use blank lines to separate logical sections of code
- Use a single blank line between function and class definitions
- Use a single space around operators: `x = 1 + 2`

#### 4.4 String Formatting

- Use string format method for complex strings:
  ```python
  message = "Parameter '{}' has value '{}'".format(param_name, param_value)
  ```
- Use f-strings where appropriate (if using Python 3.6+):
  ```python
  message = f"Parameter '{param_name}' has value '{param_value}'"
  ```

### 5. OOP Practices

#### 5.1 Class Structure

- Use proper encapsulation with clear method visibility
- Implement properties using decorators for controlled attribute access
- Use inheritance appropriately, preferring composition when suitable

```python
class ExampleClass(object):
    """Example class with proper structure."""
    
    def __init__(self, value):
        """Initialize with a value."""
        self._value = value
        self._calculated = None
    
    @property
    def value(self):
        """Get the stored value."""
        return self._value
    
    @value.setter
    def value(self, new_value):
        """Set the value with validation."""
        if new_value < 0:
            raise ValueError("Value must be positive")
        self._value = new_value
        self._calculated = None
    
    def calculate(self):
        """Calculate a result based on the value."""
        if self._calculated is None:
            self._calculated = self._value * 2
        return self._calculated
```

#### 5.2 Interface Implementation

For implementing .NET interfaces like `IExternalEventHandler` or `INotifyPropertyChanged`:

- Implement all required methods of the interface
- Follow the established naming conventions of the interface
- Add appropriate documentation to explain the purpose of each method

```python
class PropertyChangedImplementation(Object, INotifyPropertyChanged):
    """Implementation of INotifyPropertyChanged."""
    
    def __init__(self):
        """Initialize the object."""
        self._property_changed = None
    
    def add_PropertyChanged(self, handler):
        """Add a property changed event handler."""
        if self._property_changed is None:
            self._property_changed = handler
        else:
            self._property_changed += handler
    
    def remove_PropertyChanged(self, handler):
        """Remove a property changed event handler."""
        if self._property_changed is not None:
            self._property_changed -= handler
    
    def OnPropertyChanged(self, property_name):
        """Notify listeners that a property value has changed."""
        if self._property_changed is not None:
            self._property_changed(self, PropertyChangedEventArgs(property_name))
```
