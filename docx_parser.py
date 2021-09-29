#!/usr/bin/env python3
#
# docx_parser.py
#
# VERSION: 0.0.3
# UPDATED: 2021-09-29
#
##############################################################################
# PUBLIC DOMAIN NOTICE                                                       #
##############################################################################
# This software is freely available to the public for use.                   #
#                                                                            #
# Although all reasonable efforts have been taken to ensure the accuracy and #
# reliability of the software, the author does not and cannot warrant the    #
# performance or results that may be obtained by using this software.        #
# The author disclaims all warranties, express or implied, including         #
# warranties of performance, merchantability or fitness for any particular   #
# purpose.                                                                   #
#                                                                            #
# Please cite the author in any work or product based on this material.      #
#    Tyler W. Davis, PhD                                                     #
#    https://github.com/dt-woods/                                            #
##############################################################################
#
##############################################################################
# REQUIRED MODULES
##############################################################################
import os
import glob
import re

import docx

##############################################################################
# FUNCTIONS
##############################################################################
def find_word_files(d, k=""):
    """
    Name:     find_word_files
    Inputs:   - str, file path (d)
              - str, keyword(s) in the file to search (k)
    Outputs:  List
    Features: Searches the given directory for word files
    """
    my_search = "*%s*.docx" % (k)
    my_files = glob.glob(os.path.join(d, my_search))
    return my_files


def list_paragraph_styles(d):
    """
    Name:     list_paragraph_styles
    Inputs:   docx.document.Document, open word document
    Output:   dict, style_id (keys) with name and counts (keys) found
    Features: Returns a list of all the paragraph styles found in given doc
    """
    style_dict = {}
    para_num = len(my_doc.paragraphs)
    for i in range(para_num):
        para = my_doc.paragraphs[i]
        if para.style.style_id not in style_dict:
            style_dict[para.style.style_id] = {
                'name': para.style.name,
                'count': 1
            }
        else:
            style_dict[para.style.style_id]['count'] += 1
    return style_dict


def parse_file(d, styleid):
    """
    Name:     parse_file
    Inputs:   - docx.document.Document, an open Word document (d)
              - str, the Word paragraph style ID to break on (styleid)
    Features: Finds paragraphs of the given style and breaks it into a
              separate document
    TODO:     - include a user-defined output folder option
              - include a user-defined output file namning scheme
    """
    # Initialize output document (i.e., the chapter in a book to be written)
    my_out = None
    para_num = len(my_doc.paragraphs)
    j = 1
    for i in range(para_num):
        para = my_doc.paragraphs[i]
        # Split document on given style
        if para.style.style_id == styleid:
            if my_out:
                my_out.save(my_name)
            my_name = "DOCUMENT_%d.docx" % (j)
            j += 1
            my_out = docx.Document()
        if my_out:
            my_out.add_paragraph(para.text, para.style.name)

##############################################################################
# MAIN
##############################################################################
# Step 1: find the input word files
# TODO: allow for user-defined folders
my_dir = "examples"
my_files = find_word_files(my_dir, "example")
if len(my_files) == 1:
    my_file = my_files[0]
elif len(my_files) > 1:
    print("Found several word files; "
          "please use keywords to specify the one you want.")
else:
    print("Failed to find docx. Please check and try again.")
    my_file = None

# Step 2 - Open the document, define break style, and parse
if my_file:
    my_doc = docx.Document(my_file)
    my_styles = list_paragraph_styles(my_doc)
    bstyle = "Heading1"
    if bstyle in my_styles.keys():
        parse_file(my_doc, bstyle)
