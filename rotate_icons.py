#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Rotate icon images counter-clockwise 90 degrees"""

from PIL import Image
import os

# Base directory
base_dir = r'C:\code\pyRevit Extensions\pyByggstyrning.extension'

# Image paths
images = [
    os.path.join(base_dir, r'pyBS.tab\3D Zone.panel\col2.stack\3D Zones from.splitpushbutton\3D Zones from Rooms.pushbutton\icon.png'),
    os.path.join(base_dir, r'pyBS.tab\3D Zone.panel\col2.stack\3D Zones from.splitpushbutton\3D Zones from Areas.pushbutton\icon.png')
]

for img_path in images:
    if os.path.exists(img_path):
        print("Rotating: {}".format(img_path))
        img = Image.open(img_path)
        # Rotate counter-clockwise 90 degrees (negative value)
        rotated = img.rotate(-90, expand=True)
        # Save back to the same file
        rotated.save(img_path)
        print("  Successfully rotated {}".format(os.path.basename(img_path)))
    else:
        print("File not found: {}".format(img_path))

print("Done!")

