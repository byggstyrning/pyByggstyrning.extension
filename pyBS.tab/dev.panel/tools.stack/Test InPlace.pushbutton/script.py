# -*- coding: utf-8 -*-
"""Test script for in-place element zone containment.

This script tests the improved in-place family handling for 3D zone containment.
It reports on:
1. Detection of in-place families
2. Representative point extraction from in-place families
3. Multi-solid geometry handling

Run this with a model containing in-place elements and zones (Mass/Generic Model).
"""

from __future__ import print_function
from pyrevit import script, forms

# Revit API imports
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, FamilyInstance,
    Options, XYZ
)

# Import the containment module
from lib.zone3d.containment import (
    is_inplace_family,
    get_element_representative_point,
    get_element_test_points,
    is_point_in_element,
    _collect_all_solids_from_geometry
)

__title__ = "Test\nInPlace"
__doc__ = "Test in-place element zone containment detection."
__author__ = "pyByggstyrning"

# Get document
doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

def test_inplace_detection():
    """Test 1: Detect all in-place families in the model."""
    output.print_md("## Test 1: In-Place Family Detection")
    
    # Collect all family instances
    all_instances = FilteredElementCollector(doc)\
        .OfClass(FamilyInstance)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    inplace_count = 0
    inplace_elements = []
    
    for inst in all_instances:
        if is_inplace_family(inst):
            inplace_count += 1
            inplace_elements.append(inst)
            
            # Get family info
            try:
                family_name = inst.Symbol.Family.Name
                category = inst.Category.Name if inst.Category else "Unknown"
                output.print_md("- **{}** (ID: {}) - Category: {}".format(
                    family_name, inst.Id.IntegerValue, category))
            except:
                output.print_md("- Element ID: {} (could not get name)".format(inst.Id.IntegerValue))
    
    output.print_md("\n**Total in-place families found: {}** out of {} family instances\n".format(
        inplace_count, len(list(all_instances))))
    
    return inplace_elements

def test_representative_points(inplace_elements):
    """Test 2: Get representative points for in-place elements."""
    output.print_md("## Test 2: Representative Point Extraction")
    
    if not inplace_elements:
        output.print_md("No in-place elements to test.")
        return
    
    for elem in inplace_elements[:5]:  # Test first 5
        try:
            family_name = elem.Symbol.Family.Name
            
            # Get representative point
            point = get_element_representative_point(elem, doc)
            
            # Get Location for comparison
            loc = elem.Location
            loc_point = None
            if loc and hasattr(loc, "Point"):
                loc_point = loc.Point
            
            # Report
            if point:
                output.print_md("- **{}** (ID: {})".format(family_name, elem.Id.IntegerValue))
                output.print_md("  - Representative point: ({:.2f}, {:.2f}, {:.2f})".format(
                    point.X, point.Y, point.Z))
                if loc_point:
                    output.print_md("  - Location.Point: ({:.2f}, {:.2f}, {:.2f})".format(
                        loc_point.X, loc_point.Y, loc_point.Z))
                    # Check if Location.Point is at origin (potentially unreliable)
                    if abs(loc_point.X) < 0.01 and abs(loc_point.Y) < 0.01 and abs(loc_point.Z) < 0.01:
                        output.print_md("  - **WARNING**: Location.Point is at origin (0,0,0) - unreliable!")
                else:
                    output.print_md("  - Location.Point: None")
            else:
                output.print_md("- **{}**: Could not determine representative point!".format(family_name))
        except Exception as e:
            output.print_md("- Error processing element {}: {}".format(elem.Id.IntegerValue, str(e)))
    
    output.print_md("")

def test_multi_solid_geometry(inplace_elements):
    """Test 3: Check multi-solid geometry for in-place elements."""
    output.print_md("## Test 3: Multi-Solid Geometry Analysis")
    
    if not inplace_elements:
        output.print_md("No in-place elements to test.")
        return
    
    options = Options()
    options.ComputeReferences = False
    
    multi_solid_elements = []
    
    for elem in inplace_elements:
        try:
            family_name = elem.Symbol.Family.Name
            
            # Get geometry
            geometry = elem.get_Geometry(options)
            if not geometry:
                continue
            
            # Collect all solids
            solids = _collect_all_solids_from_geometry(geometry)
            
            if len(solids) > 1:
                multi_solid_elements.append((elem, len(solids)))
                output.print_md("- **{}** (ID: {}) has **{} solids**".format(
                    family_name, elem.Id.IntegerValue, len(solids)))
                
                # Report solid volumes
                for i, solid in enumerate(solids):
                    if hasattr(solid, "Volume"):
                        output.print_md("  - Solid {}: Volume = {:.2f} cu.ft".format(i+1, solid.Volume))
            elif len(solids) == 1:
                output.print_md("- {} (ID: {}) has 1 solid".format(family_name, elem.Id.IntegerValue))
            else:
                output.print_md("- {} (ID: {}) has NO solids".format(family_name, elem.Id.IntegerValue))
                
        except Exception as e:
            output.print_md("- Error processing element {}: {}".format(elem.Id.IntegerValue, str(e)))
    
    output.print_md("\n**Elements with multiple solids: {}**\n".format(len(multi_solid_elements)))

def test_containment_with_zones():
    """Test 4: Test containment detection with zones (Mass/Generic Model)."""
    output.print_md("## Test 4: Zone Containment Detection")
    
    # Find zone elements (Mass or Generic Model)
    zones = []
    
    # Collect Mass elements
    mass_elements = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_Mass)\
        .WhereElementIsNotElementType()\
        .ToElements()
    zones.extend(mass_elements)
    
    # Collect Generic Model elements
    gm_elements = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_GenericModel)\
        .WhereElementIsNotElementType()\
        .ToElements()
    zones.extend(gm_elements)
    
    if not zones:
        output.print_md("No Mass or Generic Model elements found to use as zones.")
        return
    
    output.print_md("Found {} potential zone elements (Mass + Generic Model)".format(len(zones)))
    
    # Check for in-place zones
    inplace_zones = [z for z in zones if is_inplace_family(z)]
    output.print_md("- {} are in-place families".format(len(inplace_zones)))
    
    # Test containment for in-place zones
    for zone in inplace_zones[:3]:  # Test first 3 in-place zones
        try:
            family_name = zone.Symbol.Family.Name if hasattr(zone, 'Symbol') else "Unknown"
            output.print_md("\n### Zone: {} (ID: {})".format(family_name, zone.Id.IntegerValue))
            
            # Get zone geometry info
            options = Options()
            geometry = zone.get_Geometry(options)
            solids = _collect_all_solids_from_geometry(geometry)
            output.print_md("- Has {} solid(s)".format(len(solids)))
            
            # Get test points
            test_points = get_element_test_points(zone, doc)
            output.print_md("- Generated {} test points".format(len(test_points)))
            
            # Test a point inside one of the solids (use centroid)
            if solids:
                for i, solid in enumerate(solids):
                    if hasattr(solid, "ComputeCentroid"):
                        centroid = solid.ComputeCentroid()
                        if centroid:
                            is_inside = is_point_in_element(zone, centroid, doc)
                            output.print_md("- Solid {} centroid containment test: {}".format(
                                i+1, "PASS" if is_inside else "FAIL"))
                
        except Exception as e:
            output.print_md("- Error testing zone {}: {}".format(zone.Id.IntegerValue, str(e)))

# Run all tests
output.print_md("# In-Place Element Zone Containment Test\n")
output.print_md("Testing improved in-place family handling...\n")

inplace_elements = test_inplace_detection()
test_representative_points(inplace_elements)
test_multi_solid_geometry(inplace_elements)
test_containment_with_zones()

output.print_md("\n---\n**Test complete.**")
