# -*- coding: utf-8 -*-
"""
Created on Thu Dec 11 05:45:14 2025

@author: Nobel
"""

import pandas as pd 
import shutil
import zipfile
import os

os.chdir(r"E:\SAS-Excel-Fix")


df = pd.read_excel('testExcel.xlsx')

# File names relative to working directory
original_file = "testExcel.xlsx"
copy_file = "testExcel-copy.xlsx"
temp_zip = "temp.zip"
temp_dir = "unzipped"

# Step 1: Copy the Excel file
shutil.copy(original_file, copy_file)

# Step 2: Rename to zip
shutil.copy(copy_file, temp_zip)

# Step 3: Extract the zip
with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
    zip_ref.extractall(temp_dir)

# Step 4: Replace core.xml
core_xml_path = os.path.join(temp_dir, "docProps", "core.xml")

new_core_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Vikas .</dc:creator>
  <cp:lastModifiedBy>Vikas .</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2015-06-05T18:17:20Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2025-12-11T01:48:06Z</dcterms:modified>
</cp:coreProperties>
"""

with open(core_xml_path, "w", encoding="utf-8") as f:
    f.write(new_core_xml)

# Step 5: Re-zip the folder
new_zip = "report_modified.zip"
with zipfile.ZipFile(new_zip, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
    for root, dirs, files in os.walk(temp_dir):
        for file in files:
            file_path = os.path.join(root, file)
            arcname = os.path.relpath(file_path, temp_dir)
            zip_ref.write(file_path, arcname)

# Step 6: Rename back to .xlsx
final_excel = "report_modified.xlsx"
shutil.move(new_zip, final_excel)


df = pd.read_excel(final_excel)
