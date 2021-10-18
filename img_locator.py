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


##############################################################################
# MAIN
##############################################################################
docx_file = "example-2.docx"
docx_path = os.path.join("examples", docx_file)

# Find the XML document in the zip file:
my_zip = ZipFile(docx_path, 'r')
my_doc = ""
my_rel = ""
for zc in my_zip.namelist():
    if "document.xml" in zc and not zc.endswith('rels'):
        my_doc = zc
    if "document.xml" in zc and zc.endswith('rels'):
        my_rel = zc

# Read the XML to a string
if my_doc != "":
    my_xml = ""
    with my_zip.open(my_doc) as f:
        for d in f.readlines():
            my_xml = "".join([my_xml, d.decode("utf-8")])

if my_rel != "":
    rel_xml = ""
    with my_zip.open(my_rel) as f:
        for d in f:
            rel_xml = "".join([rel_xml, d.decode("utf-8")])

# Parse the XML
my_root = ElementTree.fromstring(my_xml)
rel_root = ElementTree.fromstring(rel_xml)

# Get namespaces
my_ns = re.findall(r'xmlns:(\S+)="(\S+)"', my_xml)
ns_dict = {}
for ns in my_ns:
    k, v = ns
    ns_dict[k] = v

# Find the image reference IDs:
my_blips = re.findall('a:blip r:embed="(\S+)"', my_xml)

# Find the images associated with the IDs
# >>> TODO: match 'v' for 'Id' if in my_blips
for my_rel in rel_root:
    for k, v in my_rel.attrib.items():
        if k == "Id":
            print(v)

# Find and count paragraphs:
my_paras = my_root[0].findall("w:p", ns_dict)
num_paras = len(my_paras)

# Find and count runs, save them based on their paragraph ID:
num_runs = 0
my_runs = {}
for para in my_paras:
    # Find and count all runs in paragraph
    runs = para.findall("w:r", ns_dict)
    num_runs += len(runs)

    # Get paragraph ID:
    para_id = ''
    for k, v in para.attrib.items():
        if 'paraId' in k:
            para_id = v
    if para_id in my_runs.keys():
        my_runs[para_id] += runs
    else:
        my_runs[para_id] = runs

# See if there are any drawing elements:
# >>> Finds 3 drawings in example-2.docx
my_draws = {}
for para_id, runs in my_runs.items():
    for run in runs:
        draws = run.findall("w:drawing", ns_dict)
        num_draws = len(draws)
        if num_draws > 0:
            if para_id in my_draws.keys():
                my_draws[para_id] += draws
            else:
                my_draws[para_id] = draws
