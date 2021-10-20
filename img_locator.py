#!/usr/bin/env python3
#
# img_locator.py
#
# Tyler W. Davis
#
# Searches a document.xml and document.xml.rel within a .docx file for the
# associated paragraphs that contain an image and the related image path
# within the .docx (e.g., /media/image1.png)
#
# TODO:
# - create a function that returns the paragraph number (index from 0),
#   paragraph ID (from paragraph attributes), run number (index from 0),
#   relationship ID, and image path for each run that contains an image.
# - also return the total number of paragraphs and total number of runs per
#   paragraph
#
#  DICT = {
#    'num_paras': int,
#    'num_images': int,
#    'paralist': [int, int, int],
#    'paraIdList': ["Id1", "Id2", "Id3"],
#    'imgList': ["/media/img1.png", "/media/img2.png", "/media/img3.png"],
#    'imgIdList': ['Id1', "Id2", "Id3"],
#    'paras': {
#       0: {
#         'paraId': 'ID',
#         'num_run': int,
#         'num_images': int,
#         'runs': {
#            0: {
#              'imgID': "ID",
#              'imgPath': "Path"
#            }
#         }
#       }
#    }
#  }
#
#
##############################################################################
# IMPORT NECESSARY MODULES
##############################################################################
import os
import re
from zipfile import ZipFile
import xml.etree.ElementTree as ElementTree

##############################################################################
# FUNCTIONS
##############################################################################
def open_docxml(doc_path, isrel=False):
    """
    Name:     open_docxml
    Inputs:   - str, path to a valid .docx file
              - bool, whether to search for the document.xml.rel
    Returns:  str, XML of the file's contents
    Features: Returns the XML from the document.xml within a .docx as a string
    """
    my_doc = ""
    my_xml = ""
    my_zip = None
    if os.path.isfile(doc_path):
        my_zip = ZipFile(doc_path, 'r')
        for zc in my_zip.namelist():
            if isrel:
                if "document.xml" in zc and zc.endswith('rels'):
                    my_doc = zc
            else:
                if "document.xml" in zc and not zc.endswith('rels'):
                    my_doc = zc
    if my_doc != "" and my_zip:
        with my_zip.open(my_doc) as f:
            for d in f:
                my_xml = "".join([my_xml, d.decode("utf-8")])
    return my_xml


##############################################################################
# MAIN
##############################################################################
docx_file = "example-2.docx"
docx_path = os.path.join("examples", docx_file)

# Get xml content for document.xml
doc_xml = open_docxml(docx_path)
rel_xml = open_docxml(docx_path, True)

# Parse the XML
doc_root = ElementTree.fromstring(doc_xml)
rel_root = ElementTree.fromstring(rel_xml)

# Get namespaces
ns_dict = {}
my_ns = re.findall(r'xmlns:(\S+)="(\S+)"', doc_xml)
for ns in my_ns:
    k, v = ns
    ns_dict[k] = v

# Find the image reference IDs:
# >>> Finds three in example-2.docx
my_blips = re.findall('a:blip r:embed="(\S+)"', doc_xml)

# Find the images associated with the IDs
# >>> Finds the same three IDs in document.xml (example-2.docx)
for my_rel in rel_root:
    for k, v in my_rel.attrib.items():
        if k == "Id" and v in my_blips:
            print(v)

# Find and count paragraphs:
my_paras = doc_root[0].findall("w:p", ns_dict)
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
