# -*- coding: utf-8 -*-
"""
Quick Class
Uses machine learning (CatBoost) to classify element types to IFC classes.
Note: CatBoost may not be compatible with IronPython 2.7
"""

__title__ = "Quick Class"
__author__ = "pyByggstyrning"
__doc__ = """Send element types from active view to Byggstyrning's IFC quick classification service"""

import clr
import os
import sys
import json

# .NET imports for HTTP requests
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

# CatBoost Classification API endpoints
CATBOOST_API_BASE = "https://n8n.byggstyrning.se/webhook"
CATBOOST_SINGLE_URL = CATBOOST_API_BASE + "/classify"
CATBOOST_BATCH_URL = CATBOOST_API_BASE + "/classify/batch"

class MLIFCClassifier(object):
    """Machine Learning IFC Classifier using CatBoost API"""
    
    def __init__(self):
        self.script_dir = os.path.dirname(__file__)
        self.api_available = False
        self.fallback_rules = {}
        
    def check_api_availability(self):
        """Check if CatBoost API is available"""
        try:
            # Test API connectivity with a simple request
            client = WebClient()
            client.Headers.Add("Content-Type", "application/json; charset=utf-8")
            client.Encoding = Encoding.UTF8
            
            # Simple health check payload matching API spec
            test_payload = {
                "category": "Walls",
                "family": "Basic Wall", 
                "type": "Generic - 200mm",
                "manufacturer": "Test",
                "description": ""
            }
            
            json_data = json.dumps(test_payload, ensure_ascii=False)
            response_bytes = client.UploadData(CATBOOST_SINGLE_URL, "POST", Encoding.UTF8.GetBytes(json_data))
            response = Encoding.UTF8.GetString(response_bytes)
            
            # Parse response to verify format
            response_data = json.loads(response)
            if 'result' in response_data:
                logger.info("CatBoost API is available")
                return True
            else:
                logger.warning("API response format unexpected")
                return False
            
        except Exception as e:
            logger.warning("CatBoost API not available: {}".format(str(e)))
            return False
        finally:
            if 'client' in locals():
                client.Dispose()
    
    def initialize_fallback_rules(self):
        """Initialize enhanced rule-based classification"""
        # Enhanced rule-based classification with family and type patterns
        self.fallback_rules = {
            "Walls": "IfcWall",
            "Floors": "IfcSlab", 
            "Roofs": "IfcRoof",
            "Ceilings": "IfcCovering",
            "Doors": "IfcDoor",
            "Windows": "IfcWindow",
            "Columns": "IfcColumn",
            "Structural Columns": "IfcColumn",
            "Structural Framing": "IfcBeam",
            "Stairs": "IfcStair",
            "Railings": "IfcRailing",
            "Air Terminals": "IfcAirTerminal",
            "Duct Fittings": "IfcDuctFitting",
            "Duct Accessories": "IfcDuctFitting",
            "Ducts": "IfcDuctSegment",
            "Pipe Fittings": "IfcPipeFitting",
            "Pipe Accessories": "IfcPipeFitting",
            "Pipes": "IfcPipeSegment",
            "Plumbing Fixtures": "IfcSanitaryTerminal",
            "Lighting Fixtures": "IfcLightFixture",
            "Electrical Fixtures": "IfcElectricAppliance",
            "Electrical Equipment": "IfcElectricAppliance",
            "Mechanical Equipment": "IfcUnitaryEquipment",
            "Furniture": "IfcFurniture",
            "Casework": "IfcFurniture",
            "Generic Models": "IfcBuildingElementProxy"
        }
        
        # Family/Type specific patterns
        self.family_patterns = {
            # Pattern: IFC Class
            "curtain": "IfcCurtainWall",
            "glazing": "IfcWindow", 
            "beam": "IfcBeam",
            "footing": "IfcFooting",
            "foundation": "IfcFooting",
            "slab": "IfcSlab",
            "roof": "IfcRoof",
            "covering": "IfcCovering"
        }
        logger.info("Initialized fallback rules for {} categories".format(len(self.fallback_rules)))
    
    def classify_with_catboost_api(self, elements_data):
        """Classify using CatBoost batch API"""
        try:
            # Create web client with proper encoding
            client = WebClient()
            client.Headers.Add("Content-Type", "application/json; charset=utf-8")
            client.Encoding = Encoding.UTF8
            
            # Prepare API payload matching IfcClassifyBatchRequest
            api_payload = {
                "elements": []
            }
            
            # Convert elements_data to API format
            for element_data in elements_data:
                api_payload["elements"].append({
                    "category": element_data.get("category", ""),
                    "family": element_data.get("family", ""),
                    "type": element_data.get("type_name", ""),  # Note: API uses "type" not "type_name"
                    "manufacturer": element_data.get("manufacturer", ""),
                    "description": ""  # Optional field
                })
            
            # Convert to JSON with proper encoding
            json_data = json.dumps(api_payload, ensure_ascii=False)
            logger.debug("Sending batch request: {} elements".format(len(api_payload["elements"])))
            
            # Send request to batch endpoint with UTF-8 encoding
            response_bytes = client.UploadData(CATBOOST_BATCH_URL, "POST", Encoding.UTF8.GetBytes(json_data))
            response = Encoding.UTF8.GetString(response_bytes)
            
            # Parse response
            response_data = json.loads(response)
            logger.debug("API Response received: {} results".format(len(response_data.get("results", []))))
            
            return response_data
            
        except Exception as e:
            logger.error("Error with CatBoost API classification: {}".format(str(e)))
            return None
        finally:
            if 'client' in locals():
                client.Dispose()
    
    def classify_with_rules(self, category, family, type_name, manufacturer):
        """Classify using rule-based fallback"""
        if category in self.fallback_rules:
            return self.fallback_rules[category]
        
        # Try some keyword matching
        category_lower = category.lower()
        
        # HVAC equipment
        if any(word in category_lower for word in ["hvac", "air", "duct", "terminal"]):
            return "IfcAirTerminal"
        
        # Electrical
        if any(word in category_lower for word in ["electrical", "lighting", "power"]):
            return "IfcElectricAppliance"
        
        # Plumbing
        if any(word in category_lower for word in ["plumbing", "pipe", "water"]):
            return "IfcPipeSegment"
        
        # Default to proxy for unknown types
        return "IfcBuildingElementProxy"
    
    def prepare_element_data(self, element_types):
        """Prepare element data for batch classification"""
        elements_data = []
        
        for element_type in element_types:
            try:
                # Extract features
                category = element_type.Category.Name if element_type.Category else "Unknown"
                
                # Get family name
                family = "Unknown"
                if hasattr(element_type, 'FamilyName'):
                    family = element_type.FamilyName
                elif hasattr(element_type, 'Family') and element_type.Family:
                    family = element_type.Family.Name
                
                # Get type name using PyRevit's safe method
                try:
                    type_name = query.get_name(element_type)
                except:
                    try:
                        type_name = element_type.Name
                    except:
                        type_name = "Unknown Type"
                
                # Get manufacturer
                manufacturer = ""
                manufacturer_param = element_type.LookupParameter("Manufacturer")
                if manufacturer_param and manufacturer_param.HasValue:
                    manufacturer = manufacturer_param.AsString() or ""
                
                elements_data.append({
                    'element_type': element_type,
                    'category': category,
                    'family': family,
                    'type_name': type_name,
                    'manufacturer': manufacturer
                })
                
            except Exception as e:
                try:
                    element_id = element_type.Id
                except:
                    element_id = "Unknown ID"
                logger.error("Error preparing element type {}: {}".format(element_id, str(e)))
                continue
        
        return elements_data
    
    def classify_batch(self, element_types):
        """Classify a batch of element types"""
        # Prepare element data
        elements_data = self.prepare_element_data(element_types)
        if not elements_data:
            return []
        
        results = []
        
        # Try API classification first
        if self.api_available:
            try:
                api_response = self.classify_with_catboost_api(elements_data)
                
                if api_response and 'results' in api_response:
                    # Process API results according to IfcClassifyBatchResponse format
                    api_results = api_response['results']
                    processing_time = api_response.get('processing_time_ms', 0)
                    total_elements = api_response.get('total_elements', len(api_results))
                    
                    logger.info("API processed {} elements in {:.1f}ms".format(total_elements, processing_time))
                    
                    for i, element_data in enumerate(elements_data):
                        if i < len(api_results):
                            api_result = api_results[i]
                            
                            # Extract from IfcClassificationResult format
                            ifc_class = api_result.get('ifc_class', 'IfcBuildingElementProxy')
                            predefined_type = api_result.get('predefined_type', None)
                            confidence = api_result.get('confidence', 0.0)
                            
                            results.append({
                                'element_type': element_data['element_type'],
                                'category': element_data['category'],
                                'family': element_data['family'],
                                'type_name': element_data['type_name'],
                                'manufacturer': element_data['manufacturer'],
                                'ifc_class': ifc_class,
                                'predefined_type': predefined_type,
                                'confidence': confidence,
                                'method': 'catboost_api'
                            })
                        else:
                            # Fallback for missing API results
                            ifc_class = self.classify_with_rules(
                                element_data['category'],
                                element_data['family'], 
                                element_data['type_name'],
                                element_data['manufacturer']
                            )
                            results.append({
                                'element_type': element_data['element_type'],
                                'category': element_data['category'],
                                'family': element_data['family'],
                                'type_name': element_data['type_name'],
                                'manufacturer': element_data['manufacturer'],
                                'ifc_class': ifc_class,
                                'predefined_type': None,
                                'confidence': 0.5,
                                'method': 'rules_fallback'
                            })
                    
                    return results
                    
            except Exception as e:
                logger.error("API classification failed: {}".format(str(e)))
        
        # Fallback to rule-based classification
        for element_data in elements_data:
            ifc_class = self.classify_with_rules(
                element_data['category'],
                element_data['family'],
                element_data['type_name'], 
                element_data['manufacturer']
            )
            results.append({
                'element_type': element_data['element_type'],
                'category': element_data['category'],
                'family': element_data['family'],
                'type_name': element_data['type_name'],
                'manufacturer': element_data['manufacturer'],
                'ifc_class': ifc_class,
                'predefined_type': None,
                'confidence': 0.5,
                'method': 'rules_only'
            })
        
        return results

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
    
    # Get the actual type elements
    element_types = []
    for type_id in type_ids:
        try:
            element_type = doc.GetElement(type_id)
            if element_type and isinstance(element_type, ElementType):
                element_types.append(element_type)
                logger.debug("Found ElementType: ID={}".format(type_id))
            else:
                if element_type:
                    logger.debug("Element {} is not an ElementType: {}".format(
                        type_id, type(element_type).__name__))
        except Exception as get_error:
            logger.debug("Error getting element {}: {}".format(type_id, str(get_error)))
            continue
    
    logger.info("Found {} unique element types".format(len(element_types)))
    return element_types

def apply_classifications(classification_results):
    """Apply the ML classifications to element types"""
    if not classification_results:
        return
    
    updated_count = 0
    
    with revit.Transaction("Apply ML IFC Classifications"):
        for result in classification_results:
            if not result:
                continue
                
            try:
                element_type = result['element_type']
                ifc_class = result['ifc_class']
                
                # Get the IFC export parameter
                ifc_param = element_type.LookupParameter("Export Type to IFC As")
                
                if ifc_param and not ifc_param.IsReadOnly:
                    # Only set if not already set or if different
                    current_value = ifc_param.AsString() if ifc_param.HasValue else ""
                    
                    if current_value != ifc_class:
                        ifc_param.Set(ifc_class)
                        updated_count += 1
                        try:
                            type_name = query.get_name(element_type)
                        except:
                            try:
                                type_name = element_type.Name
                            except:
                                type_name = "Unknown Type"
                        logger.debug("Set {} -> {}".format(type_name, ifc_class))
                else:
                    try:
                        type_name = query.get_name(element_type)
                    except:
                        try:
                            type_name = element_type.Name
                        except:
                            type_name = "Unknown Type"
                    logger.warning("Cannot set IFC parameter for: {}".format(type_name))
                    
            except Exception as e:
                logger.error("Error applying classification: {}".format(str(e)))
                continue
    
    return updated_count

def apply_classifications_with_progress(classification_results, progress_bar):
    """Apply the ML classifications to element types with progress tracking"""
    if not classification_results:
        return 0
    
    updated_count = 0
    total_count = len(classification_results)
    
    with revit.Transaction("Apply ML IFC Classifications"):
        for i, result in enumerate(classification_results):
            if not result:
                progress_bar.update_progress(i + 1, total_count)
                continue
                
            try:
                element_type = result['element_type']
                # Only append "Type" if it's not already there
                ifc_class = result['ifc_class']
                if not ifc_class.endswith("Type"):
                    ifc_class += "Type"
                predefined_type = result['predefined_type']
                
                # Update progress bar with current element
                try:
                    type_name = query.get_name(element_type)
                except:
                    try:
                        type_name = element_type.Name
                    except:
                        type_name = "Unknown Type"
                
                progress_bar.update_progress(i + 1, total_count)
                
                # Get the IFC export parameter
                ifc_param = element_type.LookupParameter("Export Type to IFC As")
                
                if ifc_param and not ifc_param.IsReadOnly:
                    # Only set if not already set or if different
                    current_value = ifc_param.AsString() if ifc_param.HasValue else ""
                    
                    if current_value != ifc_class:
                        ifc_param.Set(ifc_class)
                        updated_count += 1
                        logger.debug("Set {} -> {}".format(type_name, ifc_class))
                        if predefined_type:
                            predefined_param = element_type.LookupParameter("Type IFC Predefined Type")
                            if predefined_param and not predefined_param.IsReadOnly:
                                predefined_param.Set(predefined_type)
                                logger.debug("Type IFC Predefined Type: {}".format(predefined_type))
                    

                else:
                    logger.warning("Cannot set IFC parameter for: {}".format(type_name))
                    
            except Exception as e:
                logger.error("Error applying classification: {}".format(str(e)))
                progress_bar.update_progress(i + 1, total_count)
                continue
    
    return updated_count

def display_results(classification_results):
    """Display classification results"""
    if not classification_results:
        return
        
    output = script.get_output()
    output.print_md("# ML IFC Classification Results")
    output.print_md("---")
    
    api_count = sum(1 for r in classification_results if r and r['method'] == 'catboost_api')
    rules_count = sum(1 for r in classification_results if r and 'rules' in r['method'])
    
    output.print_md("**Classification Methods:**")
    output.print_md("- CatBoost API: {} types".format(api_count))
    output.print_md("- Rule-based: {} types".format(rules_count))
    output.print_md("---")
    
    for i, result in enumerate([r for r in classification_results if r], 1):
        method_icon = "🤖" if result['method'] == 'catboost_api' else "📋"
        confidence_text = ""
        
        if 'confidence' in result and result['confidence'] > 0:
            confidence_text = " (confidence: {:.1%})".format(result['confidence'])
        
        output.print_md("## {}. {} {}{}".format(i, method_icon, result['type_name'], confidence_text))
        output.print_md("- **Category:** {}".format(result['category']))
        output.print_md("- **Family:** {}".format(result['family']))
        output.print_md("- **IFC Class:** {}".format(result['ifc_class']))
        
        # Show predefined type if available
        if result.get('predefined_type'):
            output.print_md("- **Predefined Type:** {}".format(result['predefined_type']))
            
        output.print_md("- **Method:** {}".format(result['method']))
        output.print_md("---")

def main():
    """Main function"""
    try:
        # Initialize classifier
        classifier = MLIFCClassifier()
        
        # Check API availability
        classifier.api_available = classifier.check_api_availability()
                
        # Initialize fallback rules regardless
        classifier.initialize_fallback_rules()
        
        # Log method being used (no alert)
        if classifier.api_available:
            logger.info("Using CatBoost API with rule-based fallback")
        else:
            logger.info("CatBoost API not available - using rule-based classification only")
        
        # Get element types from active view
        element_types = get_element_types_in_view()
        
        if not element_types:
            logger.warning("No element types found in the active view.")
            return
        
        # Show selection dialog with enhanced debugging
        type_names = []
        for i, et in enumerate(element_types):
            try:
                # Debug element type properties
                logger.debug("Element Type {}: ID={}, Type={}, IsElementType={}".format(
                    i, et.Id, type(et).__name__, hasattr(et, 'Name')))
                
                # Check if this is actually an ElementType
                if not isinstance(et, ElementType):
                    logger.warning("Element {} is not an ElementType: {}".format(i, type(et).__name__))
                    type_names.append("Not an ElementType - {}".format(type(et).__name__))
                    continue
                
                # Try to get category safely
                category_name = "Unknown"
                try:
                    if hasattr(et, 'Category') and et.Category:
                        category_name = et.Category.Name
                        logger.debug("  Category: {}".format(category_name))
                    else:
                        logger.debug("  No Category found")
                except Exception as cat_e:
                    logger.debug("  Category error: {}".format(str(cat_e)))
                
                # Try to get type name using PyRevit's safe method
                type_name = "Unknown Type"
                try:
                    # Use PyRevit's get_name function which handles different Revit versions
                    type_name = query.get_name(et)
                    logger.debug("  Name (PyRevit): {}".format(type_name))
                except Exception as name_e:
                    logger.debug("  Name error: {}".format(str(name_e)))
                    # Fallback to direct access if PyRevit method fails
                    try:
                        type_name = et.Name
                        logger.debug("  Name (direct fallback): {}".format(type_name))
                    except:
                        logger.debug("  All name methods failed")
                
                type_names.append("{} - {}".format(category_name, type_name))
                
            except Exception as e:
                logger.warning("Error processing element type {}: {}".format(i, str(e)))
                type_names.append("Error - {}".format(str(e)))
        
        selected_indices = forms.SelectFromList.show(
            type_names,
            title="Select Element Types to Classify",
            width=600,
            height=400,
            multiselect=True,
            button_name="Classify Selected"
        )
        
        if not selected_indices:
            return
        
        selected_types = [element_types[i] for i in range(len(element_types)) 
                         if type_names[i] in selected_indices]
        
        logger.info("User selected {} types for ML classification".format(len(selected_types)))
        
        # Classify selected types (no progress bar for API request)
        logger.info("Sending batch request to CatBoost API...")
        classification_results = classifier.classify_batch(selected_types)
        
        # Display results
        display_results(classification_results)
        
        # Ask if user wants to apply classifications
        apply_changes = forms.alert(
            "Apply these classifications to the element types?",
            title="Apply Classifications",
            ok=False,
            yes=True,
            no=True
        )
        
        if apply_changes:
            # Use progress bar only for applying parameters
            with forms.ProgressBar(title="Applying IFC Classifications to Element Types...") as pb:
                pb.update_progress(0, len(classification_results))
                updated_count = apply_classifications_with_progress(classification_results, pb)
            
            logger.info("Applied classifications to {} element types.".format(updated_count))
            
    except Exception as e:
        logger.error("Error in main function: {}".format(str(e)))
        forms.alert("An error occurred: {}".format(str(e)), title="Error")

if __name__ == '__main__':
    main() 