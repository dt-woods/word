#!/usr/bin/env python3
#
# img_locator.py
#
# Tyler W. Davis
#
# Searches a document.xml for associated paragraphs that contain an image.
#
##############################################################################
# IMPORT NECESSARY MODULES
##############################################################################
import os
import re
from zipfile import ZipFile
import xml.etree.ElementTree as ElementTree

# MAIN
docx_file = "example-2.docx"
docx_path = os.path.join("examples", docx_file)

# Find the XML document in the zip file:
my_zip = ZipFile(docx_path, 'r')
my_doc = ""
for zc in my_zip.namelist():
    if "document.xml" in zc and not zc.endswith('rels'):
        my_doc = zc

# Read the XML to a string
if my_doc != "":
    my_xml = ""
    with my_zip.open(my_doc) as f:
        for d in f.readlines():
            my_xml = "".join([my_xml, d.decode("utf-8")])

# Parse the XML using BS
my_root = ElementTree.fromstring(my_xml)

# Look through paragraphs:
# TODO: consider adding a namespace dictionary and using it for find / find all
num_paragraphs = 0
num_runs = 0
for child in my_root[0]:
    # Check for drawings as paragraphs
    for cchild in child[0]:
        # use sub to remove namespace
        cc_tag = re.sub("{.*}", "", cchild.tag)
        print(cc_tag)

    # Look through runs
    for cchild in child:
        cc_tag = re.sub("{.*}", "", cchild.tag)
        if cc_tag == 'r':
            num_runs += 1

    my_tag = re.sub("{.*}", "", child.tag)
    if my_tag == 'p':
        num_paragraphs += 1
print(num_paragraphs)
print(num_runs)
