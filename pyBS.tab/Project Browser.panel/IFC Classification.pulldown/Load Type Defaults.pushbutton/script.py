# -*- coding: utf-8 -*-
"""
Load Type Defaults
Loads IFC type defaults from Excel file based on Revit categories
and applies them to element types in the project.
"""

__title__ = "Load Type Defaults"
__author__ = "pyByggstyrning"
__doc__ = """Load IFC type defaults from Excel file based on Revit categories"""

import clr
import os
import sys

# PyRevit Excel support - use built-in modules instead of Office Interop
try:
    import xlrd
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# Revit imports
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# PyRevit imports
import pyrevit
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms

# Get current document and logger
doc = revit.doc
logger = script.get_logger()

class IFCTypeDefaultsLoader(object):
    """Loads IFC type defaults from Excel and applies them to Revit types"""
    
    def __init__(self):
        self.script_dir = os.path.dirname(__file__)
        self.excel_file = os.path.join(self.script_dir, "OOTB_revit-ifc_classes.xlsx")
        self.type_mappings = {}
        
    def load_excel_mappings(self):
        """Load IFC mappings from Excel file using PyRevit's built-in Excel support"""
        if not EXCEL_AVAILABLE:
            forms.alert("Excel reading support not available. Please ensure xlrd is installed.", 
                       title="Excel Support Missing")
            return False
            
        if not os.path.exists(self.excel_file):
            forms.alert("Excel file not found: {}".format(self.excel_file), 
                       title="File Not Found")
            return False
        
        try:
            # Open Excel file using PyRevit's xlrd
            workbook = xlrd.open_workbook(self.excel_file)
            
            # Get first worksheet
            worksheet = workbook.sheet_by_index(0)
            
            # Read data (assuming columns A=Revit Category, B=IFC Target)
            # Start from row 1 (xlrd is 0-indexed, assuming row 0 has headers)
            for row_idx in range(1, worksheet.nrows):
                try:
                    # Get cell values
                    revit_category_cell = worksheet.cell(row_idx, 0)  # Column A
                    ifc_target_cell = worksheet.cell(row_idx, 1)      # Column B
                    
                    # Extract string values
                    revit_category = revit_category_cell.value if revit_category_cell.value else ""
                    ifc_target = ifc_target_cell.value if ifc_target_cell.value else ""
                    
                    # Skip empty rows
                    if not revit_category or not ifc_target:
                        continue
                        
                    # Store mapping
                    self.type_mappings[str(revit_category).strip()] = str(ifc_target).strip()
                    logger.debug("Mapped: {} -> {}".format(revit_category, ifc_target))
                    
                except Exception as e:
                    logger.warning("Error reading row {}: {}".format(row_idx + 1, str(e)))
                    continue
            
            logger.info("Loaded {} category mappings from Excel".format(len(self.type_mappings)))
            return True
            
        except Exception as e:
            logger.error("Error reading Excel file: {}".format(str(e)))
            forms.alert("Error reading Excel file: {}".format(str(e)), title="Excel Error")
            return False

    def get_element_types_by_category(self):
        """Get all element types grouped by category"""
        collector = FilteredElementCollector(doc)
        element_types = collector.WhereElementIsElementType().ToElements()
        
        types_by_category = {}
        
        for element_type in element_types:
            try:
                if element_type.Category:
                    category_name = element_type.Category.Name
                    
                    if category_name not in types_by_category:
                        types_by_category[category_name] = []
                    
                    types_by_category[category_name].append(element_type)
                    
            except Exception as e:
                try:
                    element_id = element_type.Id
                except:
                    element_id = "Unknown ID"
                logger.warning("Error processing type {}: {}".format(element_id, str(e)))
                continue
        
        return types_by_category

    def apply_ifc_defaults(self, selected_categories=None):
        """Apply IFC defaults to element types"""
        types_by_category = self.get_element_types_by_category()
        
        if selected_categories is None:
            # Show category selection if none specified
            available_categories = [cat for cat in types_by_category.keys() 
                                  if cat in self.type_mappings]
            
            if not available_categories:
                forms.alert("No categories found that match the Excel mappings.", 
                           title="No Matches")
                return
            
            selected_categories = forms.SelectFromList.show(
                available_categories,
                title="Select Categories to Apply IFC Defaults",
                width=400,
                height=300,
                multiselect=True,
                button_name="Apply Defaults"
            )
            
            if not selected_categories:
                return
        
        # Apply defaults
        updated_count = 0
        total_count = 0
        
        with revit.Transaction("Apply IFC Type Defaults"):
            for category_name in selected_categories:
                if category_name not in types_by_category:
                    continue
                    
                if category_name not in self.type_mappings:
                    logger.warning("No IFC mapping found for category: {}".format(category_name))
                    continue
                
                ifc_class = self.type_mappings[category_name]
                element_types = types_by_category[category_name]
                
                logger.info("Applying IFC class '{}' to {} types in category '{}'".format(
                    ifc_class, len(element_types), category_name))
                
                for element_type in element_types:
                    try:
                        total_count += 1
                        
                        # Get the IFC export parameter
                        ifc_param = element_type.LookupParameter("Export Type to IFC As")
                        
                        if ifc_param and not ifc_param.IsReadOnly:
                            # Only set if not already set or if different
                            current_value = ifc_param.AsString() if ifc_param.HasValue else ""
                            
                            if current_value != ifc_class:
                                ifc_param.Set(ifc_class)
                                updated_count += 1
                                try:
                                    type_name = element_type.Name
                                except:
                                    type_name = "Unknown Type"
                                logger.debug("Set {} -> {}".format(type_name, ifc_class))
                        else:
                            try:
                                type_name = element_type.Name
                            except:
                                type_name = "Unknown Type"
                            logger.warning("Cannot set IFC parameter for: {}".format(type_name))
                            
                    except Exception as e:
                        try:
                            type_name = element_type.Name
                        except:
                            type_name = "Unknown Type"
                        logger.error("Error setting IFC class for {}: {}".format(
                            type_name, str(e)))
                        continue
        
        # Show results
        message = "IFC Defaults Applied:\n\n"
        message += "Updated: {} types\n".format(updated_count)
        message += "Total processed: {} types\n".format(total_count)
        message += "Categories: {}".format(len(selected_categories))
        
        logger.info("Applied IFC defaults: {}/{} types updated".format(updated_count, total_count))

def main():
    """Main function"""
    try:
        # Initialize loader
        loader = IFCTypeDefaultsLoader()
        
        # Load Excel mappings
        if not loader.load_excel_mappings():
            return
        
        if not loader.type_mappings:
            forms.alert("No category mappings found in Excel file.", 
                       title="No Data")
            return
        
        # Show available mappings
        output = script.get_output()
        output.print_md("# Available IFC Category Mappings")
        output.print_md("---")
        
        for category, ifc_class in sorted(loader.type_mappings.items()):
            output.print_md("**{}** -> {}".format(category, ifc_class))
        
        output.print_md("---")
        output.print_md("Total mappings: {}".format(len(loader.type_mappings)))
        
        # Apply defaults
        loader.apply_ifc_defaults()
        
    except Exception as e:
        logger.error("Error in main function: {}".format(str(e)))
        forms.alert("An error occurred: {}".format(str(e)), title="Error")

if __name__ == '__main__':
    main() 