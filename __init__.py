# pyByggstyrning extension for PyRevit
"""
Main package for pyByggstyrning extension.
"""

import os
import sys

# Add extension lib folder to path
extension_dir = os.path.dirname(__file__)
lib_path = os.path.join(extension_dir, 'lib')
if lib_path not in sys.path:
    sys.path.append(lib_path) 