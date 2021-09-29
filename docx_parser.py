#!/usr/bin/env python3
#
# docx_parser.py
#
# VERSION: 0.0.1
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
from docx.enum.section import WD_SECTION

##############################################################################
# MAIN
##############################################################################
# Step 1: find the input word files
# TODO: allow for user-defined folders and search strings for file names
my_dir = "examples"
my_files = glob.glob(os.path.join(my_dir, "*.docx"))
if len(my_files) == 1:
    my_file = my_files[0]
else:
    print("Failed to find docx. Please check and try again.")
    my_file = None

# Step 2 - Opening the document, find sections, and parse the original file
if my_file:
    my_doc = docx.Document(my_file)
    # TODO: move to function
    para_num = len(my_doc.paragraphs)
    my_out = None
    for i in range(para_num):
        para = my_doc.paragraphs[i]
        # Split document on style "Heading1"
        # TODO: allow for user-defined style
        # TODO: create a list of built-in style names and/or style IDs
        if para.style.style_id == "Heading1":
            if my_out:
                # TODO: include a user-defined output folder
                my_out.save(my_name)
            # TODO: regular expressions for searching chapter numbers ?
            my_name = "DOCUMENT_%d.docx" % (i+1)
            my_out = docx.Document()
        if my_out:
            my_out.add_paragraph(para.text, para.style.name)
