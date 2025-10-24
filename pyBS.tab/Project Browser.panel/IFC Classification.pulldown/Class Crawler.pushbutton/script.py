# -*- coding: utf-8 -*-
"""
Class Crawler
Sends element types from the active view to a classification service
and displays the results.
"""

__title__ = "Class Crawler"
__author__ = "pyByggstyrning"
__doc__ = """Send element types from active view to Byggstyrning's IFC classification service"""

import clr
import json
from collections import OrderedDict

# .NET imports
clr.AddReference('System')
clr.AddReference('System.Net')
clr.AddReference('System.Web')
from System import Uri
from System.Net import WebClient, WebRequest, WebHeaderCollection
from System.Text import Encoding
from System.IO import StreamReader

# Revit imports
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# PyRevit imports
import pyrevit
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms
from pyrevit.revit.db import query

# Get current document and logger
doc = revit.doc
logger = script.get_logger()

# Classification endpoint
CLASSIFICATION_URL = "https://n8n.byggstyrning.se/webhook/classification"

# Global storage for results (needed for button callbacks)
_stored_results = []

class ElementTypeInfo(object):
    """Container for element type information"""
    
    def __init__(self, element_type):
        self.element_type = element_type
        self.category = self._get_category_name(element_type)
        self.family = self._get_family_name(element_type)
        self.type_name = self._get_type_name(element_type)
        self.manufacturer = self._get_parameter_value(element_type, "Manufacturer")
        
    def _get_category_name(self, element_type):
        """Get the category name"""
        try:
            if element_type.Category:
                return element_type.Category.Name
        except:
            pass
        return "Unknown"
    
    def _get_family_name(self, element_type):
        """Get the family name"""
        try:
            if hasattr(element_type, 'FamilyName'):
                return element_type.FamilyName
            elif hasattr(element_type, 'Family') and element_type.Family:
                return element_type.Family.Name
        except:
            pass
        return "Unknown"
    
    def _get_type_name(self, element_type):
        """Get the type name safely using PyRevit's method"""
        try:
            # Use PyRevit's get_name function which handles different Revit versions
            return query.get_name(element_type)
        except:
            try:
                # Fallback to direct access
                return element_type.Name
            except:
                return "Unknown Type"
    
    def _get_parameter_value(self, element, param_name):
        """Get parameter value by name"""
        try:
            param = element.LookupParameter(param_name)
            if param and param.HasValue:
                if param.StorageType == StorageType.String:
                    return param.AsString() or ""
                elif param.StorageType == StorageType.Integer:
                    return str(param.AsInteger())
                elif param.StorageType == StorageType.Double:
                    return str(param.AsDouble())
        except:
            pass
        return ""
    
    def to_dict(self):
        """Convert to dictionary for API request"""
        return {
            "Category": self.category,
            "Family": self.family,
            "Type": self.type_name,
            "Manufacturer": self.manufacturer
        }
    
    def __str__(self):
        """String representation for the selection list"""
        return "{} - {} - {}".format(self.category, self.family, self.type_name)

def get_element_types_in_view():
    """Get all element types visible in the active view"""
    active_view = doc.ActiveView
    if not active_view:
        forms.alert("No active view found.", title="Error")
        return []
    
    logger.info("Getting element types from view: {}".format(active_view.Name))
    
    # Get all elements in view
    collector = FilteredElementCollector(doc, active_view.Id)
    elements = collector.WhereElementIsNotElementType().ToElements()
    
    # Extract unique element types
    type_ids = set()
    for element in elements:
        try:
            type_id = element.GetTypeId()
            if type_id and type_id != ElementId.InvalidElementId:
                type_ids.add(type_id)
        except:
            continue
    
    # Get the actual type elements and validate them
    element_types = []
    for type_id in type_ids:
        try:
            element_type = doc.GetElement(type_id)
            if element_type and isinstance(element_type, ElementType):
                type_info = ElementTypeInfo(element_type)
                element_types.append(type_info)
                logger.debug("Found ElementType: ID={}".format(type_id))
            else:
                if element_type:
                    logger.debug("Element {} is not an ElementType: {}".format(
                        type_id, type(element_type).__name__))
        except Exception as e:
            logger.debug("Error processing type {}: {}".format(type_id, str(e)))
            continue
    
    logger.info("Found {} unique element types".format(len(element_types)))
    return element_types

def extract_classification_data(classification):
    """Extract IFC class, predefined type, and reasoning from classification response"""
    ifc_class = None
    predefined_type = None
    reasoning = None
    
    try:
        if isinstance(classification, dict):
            # Handle the new API response format with 'output' field
            output_data = classification.get('output', {})
            if output_data:
                ifc_class = output_data.get('Class', 'Not classified')
                predefined_type = output_data.get('PredefinedType', 'Not specified')
                reasoning = output_data.get('Reasoning', 'No reasoning provided')
            else:
                # Fallback to direct dictionary response
                ifc_class = classification.get('ifc_class', 'Not classified')
                predefined_type = classification.get('predefined_type', 'Not specified')
                reasoning = classification.get('reasoning', 'No reasoning provided')
        elif isinstance(classification, list) and len(classification) > 0:
            # Handle legacy list format
            class_output = classification[0].get('output', {})
            ifc_class = class_output.get('Class', 'Not classified')
            predefined_type = class_output.get('PredefinedType', class_output.get('Type', 'Not specified'))
            reasoning = class_output.get('Reasoning', 'No reasoning provided')
    except Exception as e:
        logger.error("Error extracting classification data: {}".format(str(e)))
        ifc_class = 'Error'
        predefined_type = 'Error'
        reasoning = 'Error parsing data'
    
    return ifc_class, predefined_type, reasoning

def html_escape(text):
    """Escape special HTML characters"""
    if not text:
        return ""
    text = str(text)
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    text = text.replace("'", "&#39;")
    return text

def create_tooltip_cell(content, tooltip):
    """Create HTML cell content with hover tooltip"""
    if not tooltip or tooltip == "No reasoning provided":
        return str(content)
    escaped_tooltip = html_escape(tooltip)
    escaped_content = html_escape(content)
    return '<span title="{}">{}</span>'.format(escaped_tooltip, escaped_content)

def send_classification_request(type_data):
    """Send classification request to the API"""
    try:
        # Create web client with proper encoding
        client = WebClient()
        client.Headers.Add("Content-Type", "application/json; charset=utf-8")
        client.Encoding = Encoding.UTF8
        
        # Convert data to JSON with proper encoding
        json_data = json.dumps(type_data, ensure_ascii=False)
        logger.debug("Sending data: {}".format(json_data))
        
        # Send request with UTF-8 encoding
        response_bytes = client.UploadData(CLASSIFICATION_URL, "POST", Encoding.UTF8.GetBytes(json_data))
        response = Encoding.UTF8.GetString(response_bytes)
        
        # Parse response
        response_data = json.loads(response)
        logger.debug("Received response: {}".format(response))
        
        return response_data
        
    except Exception as e:
        logger.error("Error sending classification request: {}".format(str(e)))
        return None
    finally:
        if 'client' in locals():
            client.Dispose()

def main():
    """Main function"""
    try:
        # Get element types from active view
        element_types = get_element_types_in_view()
        
        if not element_types:
            forms.alert("No element types found in the active view.", 
                       title="No Types Found")
            return
        
        # Show selection dialog
        selected_types = forms.SelectFromList.show(
            element_types,
            title="Select Element Types to Classify",
            width=600,
            height=400,
            multiselect=True,
            button_name="Classify Selected"
        )
        
        if not selected_types:
            return
        
        logger.info("User selected {} types for classification".format(len(selected_types)))
        
        # Process selected types
        results = []
        
        with forms.ProgressBar(title="Classifying Types...") as pb:
            for i, type_info in enumerate(selected_types):
                pb.update_progress(i + 1, len(selected_types))
                logger.debug("Classifying: {}".format(str(type_info)))
                
                # Send classification request
                type_data = type_info.to_dict()
                response = send_classification_request(type_data)
                
                if response:
                    results.append({
                        'type_info': type_info,
                        'classification': response
                    })
                else:
                    logger.warning("Failed to classify: {}".format(str(type_info)))
        
        # Display results in interactive table
        if results:
            display_results(results)
            
            # Ask if user wants to apply classifications after reviewing table
            apply_changes = forms.alert(
                "Apply these classifications to the element types?",
                title="Apply Classifications",
                ok=False,
                yes=True,
                no=True
            )
            
            if apply_changes:
                # Get output for logging
                output = script.get_output()
                
                # Use progress bar for applying parameters
                with forms.ProgressBar(title="Applying IFC Classifications...") as pb:
                    pb.update_progress(0, len(results))
                    updated_count = apply_classifications_with_progress(results, pb)
                
                # Log results to output window
                output.print_md("---")
                if updated_count > 0:
                    output.log_success("Successfully applied classifications to {} element types".format(updated_count))
                    output.print_md("**Status:** {} classifications applied successfully".format(updated_count))
                else:
                    output.log_warning("No classifications were applied")
                    output.print_md("**Status:** No changes were made")
                
                logger.info("Applied classifications to {} element types.".format(updated_count))
        else:
            forms.alert("No classification results received.", title="No Results")
            
    except Exception as e:
        logger.error("Error in main function: {}".format(str(e)))
        forms.alert("An error occurred: {}".format(str(e)), title="Error")

def display_results(results):
    """Display classification results in a custom HTML table with interactive buttons"""
    global _stored_results
    _stored_results = results
    
    output = script.get_output()
    output.print_md("# IFC Classification Results")
    output.print_md("---")
    
    # Build HTML table manually for full control
    html_parts = []
    
    # Add CSS styling
    html_parts.append("""
    <style>
        .classification-table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-family: Arial, sans-serif;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .classification-table thead {
            background-color: #2196F3;
            color: white;
        }
        .classification-table th {
            padding: 12px 8px;
            text-align: left;
            font-weight: 600;
            border: 1px solid #ddd;
        }
        .classification-table th.center {
            text-align: center;
        }
        .classification-table td {
            padding: 10px 8px;
            border: 1px solid #ddd;
            vertical-align: middle;
        }
        .classification-table td.center {
            text-align: center;
        }
        .classification-table tbody tr:nth-child(odd) {
            background-color: #f9f9f9;
        }
        .classification-table tbody tr:nth-child(even) {
            background-color: #ffffff;
        }
        .classification-table tbody tr:hover {
            background-color: #e3f2fd;
        }
        .tooltip-cell {
            cursor: help;
            border-bottom: 1px dotted #666;
        }
        .status-ready {
            padding: 6px 16px;
            background-color: #4CAF50;
            color: white;
            border-radius: 4px;
            font-size: 13px;
            font-weight: bold;
            display: inline-block;
        }
        .info-box {
            margin: 20px 0;
            padding: 15px;
            background-color: #E3F2FD;
            border-left: 4px solid #2196F3;
            border-radius: 4px;
        }
        .summary {
            margin: 20px 0;
            padding: 15px;
            background-color: #FFF3CD;
            border-left: 4px solid #FFC107;
            border-radius: 4px;
            font-weight: 600;
            color: #856404;
        }
        .apply-notice {
            margin: 20px 0;
            padding: 15px;
            background-color: #D1ECF1;
            border-left: 4px solid #17A2B8;
            border-radius: 4px;
            color: #0C5460;
        }
    </style>
    """)
    
    # Add info box
    html_parts.append("""
    <div class="info-box">
        <strong style="color: #1976D2;">📋 Review Classification Results</strong>
        <p style="margin: 10px 0 0 0; color: #555;">
            • Hover over <span style="border-bottom: 1px dotted #666;">IFC Class names</span> to see AI reasoning<br>
            • Review the suggested classifications in the table below<br>
            • After reviewing, you'll be asked to confirm applying these changes
        </p>
    </div>
    """)
    
    # Start table
    html_parts.append("""
    <table class="classification-table">
        <thead>
            <tr>
                <th style="width: 12%;">Category</th>
                <th style="width: 18%;">Family</th>
                <th style="width: 20%;">Type</th>
                <th style="width: 12%;">Manufacturer</th>
                <th style="width: 15%;">IFC Class</th>
                <th style="width: 13%;">Predefined Type</th>
                <th class="center" style="width: 10%;">Status</th>
            </tr>
        </thead>
        <tbody>
    """)
    
    # Add table rows
    for i, result in enumerate(results):
        type_info = result['type_info']
        classification = result['classification']
        
        # Extract classification data
        ifc_class, predefined_type, reasoning = extract_classification_data(classification)
        
        # Escape data for HTML
        category = html_escape(type_info.category)
        family = html_escape(type_info.family)
        type_name = html_escape(type_info.type_name)
        manufacturer = html_escape(type_info.manufacturer) if type_info.manufacturer else "-"
        ifc_class_escaped = html_escape(ifc_class)
        predefined_type_escaped = html_escape(predefined_type) if predefined_type else "-"
        reasoning_escaped = html_escape(reasoning)
        
        # Build row HTML
        html_parts.append("<tr>")
        html_parts.append("<td>{}</td>".format(category))
        html_parts.append("<td>{}</td>".format(family))
        html_parts.append("<td>{}</td>".format(type_name))
        html_parts.append("<td>{}</td>".format(manufacturer))
        
        # IFC Class with tooltip
        if reasoning and reasoning != "No reasoning provided":
            html_parts.append('<td><span class="tooltip-cell" title="{}">{}</span></td>'.format(
                reasoning_escaped, ifc_class_escaped))
        else:
            html_parts.append("<td>{}</td>".format(ifc_class_escaped))
        
        html_parts.append("<td>{}</td>".format(predefined_type_escaped))
        html_parts.append('<td class="center"><span class="status-ready">✓ Ready</span></td>')
        html_parts.append("</tr>")
    
    # Close table
    html_parts.append("""
        </tbody>
    </table>
    """)
    
    # Add summary and next steps
    html_parts.append("""
    <div class="summary">
        ⚡ <strong>Total:</strong> {} type(s) successfully classified
    </div>
    <div class="apply-notice">
        <strong>⏭️ Next Step:</strong> A confirmation dialog will appear asking if you want to apply these classifications to your element types.
    </div>
    """.format(len(results)))
    
    # Output the complete HTML
    output.print_html("".join(html_parts))

def apply_classifications_with_progress(results, progress_bar):
    """Apply the classifications to element types with progress tracking"""
    if not results:
        return 0
    
    updated_count = 0
    total_count = len(results)
    
    with revit.Transaction("Apply IFC Classifications"):
        for i, result in enumerate(results):
            if not result:
                progress_bar.update_progress(i + 1, total_count)
                continue
                
            try:
                type_info = result['type_info']
                classification = result['classification']
                element_type = type_info.element_type
                
                # Update progress bar
                progress_bar.update_progress(i + 1, total_count)
                
                # Extract IFC class and predefined type from classification response
                ifc_class = None
                predefined_type = None
                
                try:
                    if isinstance(classification, dict):
                        # Handle the new API response format with 'output' field
                        output_data = classification.get('output', {})
                        if output_data:
                            ifc_class = output_data.get('Class', None)
                            if ifc_class and not ifc_class.endswith("Type"):
                                ifc_class += "Type"
                            # Handle both 'PredefinedType' and 'Type' keys
                            predefined_type = output_data.get('PredefinedType', None) or output_data.get('Type', None)
                        else:
                            # Fallback to direct dictionary response
                            ifc_class = classification.get('ifc_class', None)
                            if ifc_class and not ifc_class.endswith("Type"):
                                ifc_class += "Type"
                            predefined_type = classification.get('predefined_type', None)
                    elif isinstance(classification, list) and len(classification) > 0:
                        # Handle legacy list format
                        class_output = classification[0].get('output', {})
                        ifc_class = class_output.get('Class', None)
                        if ifc_class and not ifc_class.endswith("Type"):
                            ifc_class += "Type"
                        predefined_type = class_output.get('PredefinedType', None) or class_output.get('Type', None)
                    
                    logger.debug("Parsed classification for {}: IFC={}, Predefined={}".format(
                        type_info.type_name, ifc_class, predefined_type))
                        
                except Exception as parse_e:
                    logger.warning("Error parsing classification for {}: {}".format(
                        type_info.type_name, str(parse_e)))
                    continue
                
                # Apply IFC Class if available
                if ifc_class:
                    ifc_param = element_type.LookupParameter("Export Type to IFC As")
                    if ifc_param and not ifc_param.IsReadOnly:
                        current_value = ifc_param.AsString() if ifc_param.HasValue else ""
                        if current_value != ifc_class:
                            ifc_param.Set(ifc_class)
                            logger.debug("Set IFC Class {} -> {}".format(type_info.type_name, ifc_class))
                            updated_count += 1
                    else:
                        logger.warning("Cannot set IFC Class parameter for: {}".format(type_info.type_name))
                
                # Apply Predefined Type if available
                if predefined_type and predefined_type.strip():
                    predefined_param = element_type.LookupParameter("Type IFC Predefined Type")
                    if not predefined_param:
                        # Try alternative parameter name
                        predefined_param = element_type.LookupParameter("Predefined Type")
                    
                    if predefined_param and not predefined_param.IsReadOnly:
                        current_predefined = predefined_param.AsString() if predefined_param.HasValue else ""
                        if current_predefined != predefined_type:
                            predefined_param.Set(predefined_type)
                            logger.debug("Set Predefined Type {} -> {}".format(type_info.type_name, predefined_type))
                    else:
                        logger.debug("No writable Predefined Type parameter found for: {}".format(type_info.type_name))
                        
            except Exception as e:
                logger.error("Error applying classification to {}: {}".format(
                    type_info.type_name if 'type_info' in locals() else "Unknown", str(e)))
                progress_bar.update_progress(i + 1, total_count)
                continue
    
    return updated_count

if __name__ == '__main__':
    main()
