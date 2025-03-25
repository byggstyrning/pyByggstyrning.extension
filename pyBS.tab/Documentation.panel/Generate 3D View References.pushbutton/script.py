# -*- coding: utf-8 -*-
"""Place 3D View Reference families at location and extent of selected views."""

__title__ = "Generate 3D View References"
__author__ = "PyRevit Extensions"
__doc__ = """This tool places workplane-based families at the location and extent of selected views.
The family instance parameters "View Height" and "View Width" are set based on the view's bounding box.
"""

# Import .NET libraries
import clr
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Collections")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Xaml")
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Windows import Application, Window
from System.Windows.Controls import CheckBox
import System.Windows.Media as Media
import System.Windows.Interop as Interop

# Import Revit API
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI import *

# Import pyRevit libraries
import os
import sys
from pyrevit import revit, DB, UI
from pyrevit import forms, script

# Get the current script directory
script_dir = script.get_script_path()
logger = script.get_logger()

# Get Revit document, app and UIDocument
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument

# Set up enhanced logging
logger.info("Starting Generate 3D View References tool")
logger.debug("Script directory: {}".format(script_dir))
logger.debug("Active document: {}".format(doc.Title))


class ViewItemData(forms.Reactive):
    """Class for view data binding with WPF UI."""
    
    def __init__(self, view):
        """Initialize with a Revit view."""
        super(ViewItemData, self).__init__()
        self.view = view
        self.view_id = view.Id
        self._is_selected = True
        
        # Set view properties
        self.view_name = view.Name
        self.view_category = self._get_view_category_name(view)
        self.view_scale = self._get_view_scale(view)
        self.sheet_reference = self._get_sheet_reference(view)
        
        logger.debug("View loaded: {}, Type: {}, Scale: {}, Sheet: {}".format(
            self.view_name, self.view_category, self.view_scale, self.sheet_reference))
    
    def _get_view_category_name(self, view):
        """Get the view category name."""
        try:
            # Try the safer approach to get category name
            if view.Category:
                return view.Category.Name
            return "Unknown"
        except:
            logger.debug("Could not get category name for view: {}".format(view.Name))
            # Try with the ViewType if category fails
            try:
                return view.ViewType.ToString()
            except:
                return "Unknown"
    
    def _get_view_scale(self, view):
        """Get the view scale."""
        try:
            scale = view.Scale
            if scale:
                return "1:{}".format(scale)
            return "Unknown"
        except:
            logger.debug("Could not get scale for view: {}".format(view.Name))
            return "Unknown"
    
    def _get_sheet_reference(self, view):
        """Get the sheet number and name if view is placed on a sheet."""
        sheet_id = None
        sheet_info = "Not on sheet"
        
        # Get sheet ID if view is on a sheet
        for sheet_id in FilteredElementCollector(doc).OfClass(ViewSheet):
            view_ids = sheet_id.GetAllPlacedViews()
            if view_ids and self.view_id in view_ids:
                sheet_num = sheet_id.SheetNumber
                sheet_name = sheet_id.Name
                sheet_info = "{} - {}".format(sheet_num, sheet_name)
                logger.debug("View {} found on sheet: {}".format(view.Name, sheet_info))
                break
        
        return sheet_info
    
    @property
    def ViewName(self):
        return self.view_name
    
    @property
    def ViewCategory(self):
        return self.view_category
    
    @property
    def ViewScale(self):
        return self.view_scale
    
    @property
    def SheetReference(self):
        return self.sheet_reference
    
    @property
    def IsSelected(self):
        return self._is_selected
    
    @IsSelected.setter
    def IsSelected(self, value):
        self._is_selected = value
        logger.debug("View selection changed: {} -> {}".format(self.view_name, value))
        self.OnPropertyChanged("IsSelected")


class Generate3DViewReferencesWindow(forms.WPFWindow):
    """WPF window for selecting and generating 3D view references."""
    
    def __init__(self):
        """Initialize the window."""
        logger.info("Initializing Generate 3D View References window")
        xaml_file = os.path.join(script_dir, "Generate3DViewReferencesWindow.xaml")
        forms.WPFWindow.__init__(self, xaml_file)
        
        # Store created elements for isolation
        self.created_elements = []
        
        # Initialize view data
        self.views_data = ObservableCollection[ViewItemData]()
        
        # Check for required family
        logger.info("Checking for 3D View Reference family on initialization")
        try:
            self.family_symbol = self._get_family_symbol()
            
            # Only try to access properties if family_symbol is not None
            if self.family_symbol:
                # Safely access Family and Name properties
                family_name = "Unknown"
                symbol_name = "Unknown"
                
                try:
                    if hasattr(self.family_symbol, "Family") and self.family_symbol.Family is not None:
                        family_name = self.family_symbol.Family.Name
                except Exception as ex:
                    logger.debug("Error accessing Family.Name: {}".format(ex))
                
                try:
                    if hasattr(self.family_symbol, "Name"):
                        symbol_name = self.family_symbol.Name
                except Exception as ex:
                    logger.debug("Error accessing Name: {}".format(ex))
                    
                logger.info("Found required family: {} - {}".format(family_name, symbol_name))
            else:
                logger.warning("Could not find 3D View Reference family with Standard Reference type during initialization")
                forms.alert(
                    "This tool requires the '3D View Reference' family with 'Standard Reference' type.\n\n"
                    "Please load this family into your project before using this tool.",
                    title="Required Family Not Found"
                )
        except Exception as ex:
            logger.error("Error initializing family symbol: {}".format(ex))
            self.family_symbol = None
        
        # Set up UI
        logger.debug("Setting up view categories")
        self._setup_view_categories()
        logger.debug("Populating views")
        self._populate_views()
        
        # Bind views to DataGrid
        self.viewsDataGrid.ItemsSource = self.views_data
        logger.info("UI setup complete. Found {} views.".format(self.views_data.Count))
        
    def _setup_view_categories(self):
        """Set up view category checkboxes."""
        # Define relevant view categories
        view_categories = [
            ("Sections", ViewType.Section),
            ("Elevations", ViewType.Elevation),
            ("Callouts", ViewType.Detail),
            ("Floor Plans", ViewType.FloorPlan),
            ("Ceiling Plans", ViewType.CeilingPlan),
            ("3D Views", ViewType.ThreeD)
        ]
        
        # Create a dictionary to store view types
        self.view_types = {}
        
        # Create checkboxes for each category
        for category_name, view_type in view_categories:
            checkbox = CheckBox()
            checkbox.Content = category_name
            checkbox.IsChecked = True
            checkbox.Tag = view_type
            checkbox.Checked += self.ViewCategory_CheckedChanged
            checkbox.Unchecked += self.ViewCategory_CheckedChanged
            
            self.viewCategoriesPanel.Children.Add(checkbox)
            self.view_types[view_type] = checkbox
            logger.debug("Added view category: {}".format(category_name))
    
    def _populate_views(self):
        """Populate the views data grid based on selected categories."""
        # Clear current views
        self.views_data.Clear()
        
        # Get all checked view types
        checked_view_types = [vt for vt, cb in self.view_types.items() if cb.IsChecked == True]
        logger.debug("Selected view types: {}".format([vt.ToString() for vt in checked_view_types]))
        
        # Filter views by checked types
        all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
        logger.debug("Total views in document: {}".format(len(all_views)))
        
        view_count = 0
        for view in all_views:
            # Filter out template views and schedules
            if view.IsTemplate or view.ViewType == ViewType.Schedule:
                continue
                
            # Check if view type is in checked types
            if view.ViewType in checked_view_types:
                # Add view to list
                self.views_data.Add(ViewItemData(view))
                view_count += 1
                
        logger.info("Added {} views to data grid".format(view_count))
    
    def ViewCategory_CheckedChanged(self, sender, args):
        """Handle view category checkbox changes."""
        category_name = sender.Content
        is_checked = sender.IsChecked
        logger.debug("View category changed: {} -> {}".format(category_name, is_checked))
        
        self._populate_views()
        
        # Update select all checkbox state
        if self.views_data.Count > 0:
            all_selected = all(view.IsSelected for view in self.views_data)
            self.selectAllCheckbox.IsChecked = all_selected
    
    def SelectAll_Checked(self, sender, args):
        """Handle select all checkbox checked."""
        logger.debug("Select All checkbox checked")
        for view_data in self.views_data:
            view_data.IsSelected = True
    
    def SelectAll_Unchecked(self, sender, args):
        """Handle select all checkbox unchecked."""
        logger.debug("Select All checkbox unchecked")
        for view_data in self.views_data:
            view_data.IsSelected = False
    
    def CreateViewReferences_Click(self, sender, args):
        """Handle create button click."""
        logger.info("Create 3D View References button clicked")
        
        # Get selected views
        selected_views = [view_data for view_data in self.views_data if view_data.IsSelected]
        logger.debug("Selected views count: {}".format(len(selected_views)))
        
        if not selected_views:
            logger.warning("No views selected")
            forms.alert("No views selected. Please select at least one view.", title="Warning")
            return
        
        # Check if family was found during initialization
        if not self.family_symbol:
            # Try one more time to find it
            logger.info("Trying again to find 3D View Reference family")
            self.family_symbol = self._get_family_symbol()
            
            if not self.family_symbol:
                logger.error("Could not find 3D View Reference family with Standard Reference type")
                forms.alert(
                    "Could not find '3D View Reference' family with 'Standard Reference' type.\n\n"
                    "Please load this family into your project before using this tool.",
                    title="Family Not Found"
                )
                return
        
        # Create 3D view references for selected views
        self._create_view_references(selected_views)
    
    def _create_view_references(self, selected_views):
        """Create 3D view references for selected views."""
        # Get family symbol from stored property 
        family_symbol = self.family_symbol
        
        # Safe way to access family symbol properties
        try:
            family_name = "Unknown Family"
            symbol_name = "Unknown Type"
            
            # Safe access to Family.Name
            if hasattr(family_symbol, "Family") and family_symbol.Family is not None:
                try:
                    family_name = family_symbol.Family.Name
                except Exception as ex:
                    logger.debug("Cannot access Family.Name: {}".format(ex))
            
            # Safe access to Name
            try:
                if hasattr(family_symbol, "Name"):
                    symbol_name = family_symbol.Name
            except Exception as ex:
                logger.debug("Cannot access Name: {}".format(ex))
                
            logger.debug("Using family symbol: {} - {}".format(family_name, symbol_name))
        except Exception as ex:
            logger.debug("Error getting family symbol details: {}".format(ex))
        
        # Clear previous created elements
        self.created_elements = []
        
        # Start transaction
        logger.info("Starting transaction to create view references")
        with revit.Transaction("Create 3D View References"):
            # Ensure family symbol is active - with safety check
            try:
                if hasattr(family_symbol, "IsActive"):
                    # Check if we need to activate it
                    try:
                        if not family_symbol.IsActive:
                            logger.debug("Activating family symbol")
                            family_symbol.Activate()
                    except Exception as ex:
                        logger.debug("Error checking/activating symbol: {}".format(ex))
            except Exception as ex:
                logger.debug("Error accessing IsActive: {}".format(ex))
            
            # Create family instances for each selected view
            success_count = 0
            for view_data in selected_views:
                view = view_data.view
                logger.debug("Processing view: {}".format(view.Name))
                
                try:
                    # Get view bounding box
                    bbox = self._get_view_bounding_box(view)
                    
                    if not bbox:
                        logger.warning("Could not get bounding box for view: {}".format(view.Name))
                        continue
                    
                    # Calculate width and height
                    width = bbox.Max.X - bbox.Min.X
                    height = bbox.Max.Y - bbox.Min.Y
                    logger.debug("View dimensions - Width: {:.2f}, Height: {:.2f}".format(width, height))
                    
                    # Get center point
                    center_point = XYZ((bbox.Min.X + bbox.Max.X) / 2, 
                                       (bbox.Min.Y + bbox.Max.Y) / 2, 
                                       0)
                    logger.debug("Center point: ({:.2f}, {:.2f}, {:.2f})".format(
                        center_point.X, center_point.Y, center_point.Z))
                    
                    # Get view direction safely
                    view_dir = XYZ(0, 0, 1)  # Default direction if we can't get the real one
                    try:
                        if hasattr(view, "ViewDirection") and view.ViewDirection is not None:
                            view_dir = view.ViewDirection
                    except Exception as ex:
                        logger.debug("Error getting view direction: {}".format(ex))
                        
                    try:
                        # Create instance safely
                        logger.debug("Creating family instance for view: {}".format(view.Name))
                        
                        # Fallback for StructuralType in case import doesn't resolve the issue
                        try:
                            structural_type = StructuralType.NonStructural
                        except NameError:
                            # If StructuralType is not defined, try alternative approach
                            logger.debug("Using alternative approach for StructuralType")
                            try:
                                from Autodesk.Revit.DB.Structure import StructuralType
                                structural_type = StructuralType.NonStructural
                            except:
                                # Final fallback - use the integer value that corresponds to NonStructural
                                logger.debug("Using integer value for StructuralType.NonStructural")
                                structural_type = 0  # 0 is the integer value for NonStructural
                        
                        # Create the family instance
                        new_instance = doc.Create.NewFamilyInstance(
                            center_point, family_symbol, view_dir, 
                            doc.ActiveView, structural_type
                        )
                        
                        # Set parameters
                        self._set_instance_parameters(new_instance, width, height)
                        
                        # Add to created elements list
                        self.created_elements.append(new_instance.Id)
                        success_count += 1
                        logger.debug("Successfully created reference for view: {}".format(view.Name))
                    except Exception as ex:
                        logger.error("Error creating family instance: {}".format(ex))
                    
                except Exception as ex:
                    logger.error("Error creating reference for view {}: {}".format(view.Name, ex))
        
        # Update isolate button text and enable it
        isolate_text = "Isolate {} created elements".format(len(self.created_elements))
        self.isolateButton.Content = isolate_text
        self.isolateButton.IsEnabled = True
        logger.debug("Updated isolate button: {}".format(isolate_text))
        
        # Show results
        result_message = "Created {} 3D View References successfully.".format(len(self.created_elements))
        logger.info(result_message)
        forms.alert(result_message, title="Success")
    
    def _get_family_symbol(self):
        """Get the family symbol for 3D View Reference."""
        logger.debug("Searching for 3D View Reference family")
        collectors = FilteredElementCollector(doc).OfClass(FamilySymbol)
        
        # Simple approach: find the 3D View Reference family
        for symbol in collectors:
            try:
                # Get family name
                if not hasattr(symbol, "Family") or symbol.Family is None:
                    continue
                
                family_name = symbol.Family.Name
                
                # Get type name safely
                type_name = "Unknown"
                try:
                    if hasattr(symbol, "Name"):
                        type_name = symbol.Name
                except:
                    pass
                
                # Look for exact family named "3D View Reference"
                if family_name == "3D View Reference":
                    logger.info("Found family named '3D View Reference'")
                    
                    # Return the first symbol from this family
                    logger.info("Found exact match for 3D View Reference family")
                    return symbol
                
            except Exception as ex:
                logger.debug("Error processing family symbol: {}".format(ex))
                continue
        
        # If we get here, we didn't find the family
        logger.warning("3D View Reference family not found")
        return None
    
    def _get_view_bounding_box(self, view):
        """Get the bounding box of the view."""
        logger.debug("Getting bounding box for view: {}".format(view.Name))
        
        # Get the crop box of the view if available
        if hasattr(view, "CropBox") and view.CropBoxActive:
            logger.debug("Using crop box for view: {}".format(view.Name))
            bbox = view.CropBox
            logger.debug("Crop box: Min({:.2f}, {:.2f}, {:.2f}), Max({:.2f}, {:.2f}, {:.2f})".format(
                bbox.Min.X, bbox.Min.Y, bbox.Min.Z, 
                bbox.Max.X, bbox.Max.Y, bbox.Max.Z))
            return bbox
        
        # If no crop box, try to get bounding box
        try:
            logger.debug("Trying to get bounding box for view: {}".format(view.Name))
            bbox = view.get_BoundingBox(view)
            if bbox:
                logger.debug("Got bounding box: Min({:.2f}, {:.2f}, {:.2f}), Max({:.2f}, {:.2f}, {:.2f})".format(
                    bbox.Min.X, bbox.Min.Y, bbox.Min.Z, 
                    bbox.Max.X, bbox.Max.Y, bbox.Max.Z))
                return bbox
        except Exception as ex:
            logger.debug("Could not get bounding box: {}".format(ex))
        
        # If no bounding box, create a default one
        try:
            # Use active view's outline if possible
            logger.debug("Trying to use view outline")
            outline = view.Outline
            min_point = XYZ(outline.Min.U, outline.Min.V, 0)
            max_point = XYZ(outline.Max.U, outline.Max.V, 0)
            logger.debug("Using outline: Min({:.2f}, {:.2f}), Max({:.2f}, {:.2f})".format(
                outline.Min.U, outline.Min.V, outline.Max.U, outline.Max.V))
            
            # Create bounding box from outline points
            bbox = BoundingBoxXYZ()
            bbox.Min = min_point
            bbox.Max = max_point
            return bbox
        except Exception as ex:
            logger.debug("Could not use view outline: {}".format(ex))
        
        # Use a default size if all else fails
        logger.warning("Using default bounding box for view: {}".format(view.Name))
        bbox = BoundingBoxXYZ()
        bbox.Min = XYZ(-5, -5, 0)
        bbox.Max = XYZ(5, 5, 0)
        return bbox
    
    def _set_instance_parameters(self, instance, width, height):
        """Set instance parameters for width and height."""
        logger.debug("Setting parameters for instance {}: Width={:.2f}, Height={:.2f}".format(
            instance.Id, width, height))
        
        try:
            # Set View Width parameter
            width_param = instance.LookupParameter("View Width")
            if width_param:
                width_param.Set(width)
                logger.debug("Set View Width parameter to {:.2f}".format(width))
            else:
                logger.warning("View Width parameter not found on instance")
            
            # Set View Height parameter
            height_param = instance.LookupParameter("View Height")
            if height_param:
                height_param.Set(height)
                logger.debug("Set View Height parameter to {:.2f}".format(height))
            else:
                logger.warning("View Height parameter not found on instance")
                
        except Exception as ex:
            logger.error("Error setting parameters: {}".format(ex))
    
    def IsolateElements_Click(self, sender, args):
        """Isolate created elements in current view."""
        logger.info("Isolate elements button clicked")
        
        if not self.created_elements:
            logger.warning("No created elements to isolate")
            return
        
        try:
            # Get current view
            current_view = doc.ActiveView
            logger.debug("Isolating {} elements in view: {}".format(
                len(self.created_elements), current_view.Name))
            
            # Create ICollection of element ids
            element_ids = List[ElementId]()
            for element_id in self.created_elements:
                element_ids.Add(element_id)
            
            # Isolate elements
            with revit.Transaction("Isolate View References"):
                current_view.IsolateElementsTemporary(element_ids)
                logger.info("Successfully isolated {} elements".format(len(element_ids)))
                
        except Exception as ex:
            logger.error("Error isolating elements: {}".format(ex))
    
    def Cancel_Click(self, sender, args):
        """Handle cancel button click."""
        logger.info("Cancel button clicked, closing window")
        self.Close()


# Run the window
if __name__ == '__main__':
    logger.info("Launching Generate 3D View References window")
    window = Generate3DViewReferencesWindow()
    window.ShowDialog()
    logger.info("Generate 3D View References tool completed")