#!/usr/bin/env python3
#
# docx_parser.py
#
# VERSION: 1.3.1
# UPDATED: 2021-10-03
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
import docx

from docx_utils import find_word_files
from docx_utils import list_paragraph_styles
from docx_utils import match_char_style
from docx_utils import match_sect_properties


##############################################################################
# FUNCTIONS
##############################################################################
def parse_file(d, styleid):
    """
    Name:     parse_file
    Inputs:   - docx.document.Document, an open Word document (d)
              - str, the Word paragraph style ID to break on (styleid)
    Features: Finds paragraphs of the given style and breaks it into a
              separate document, preserving character formats (e.g., bold)
              and document properties (e.g., margins). There is no known way
              to match a paragraph to its section, which a known issue.
    TODO:     - include a user-defined output folder option
              - include a user-defined output file naming scheme
    """
    # Initialize output document (i.e., the chapter in a book to be written)
    my_out = None
    para_num = len(d.paragraphs)
    j = 0
    for i in range(para_num):
        para = d.paragraphs[i]
        # Split document on given style
        if i == 0:
            my_name = "DOCUMENT_%d.docx" % (j)
            j += 1
            my_out = docx.Document()
            out_s = my_out.sections[0]
            match_sect_properties(d.sections[0], out_s)
        if para.style.style_id == styleid:
            if my_out:
                my_out.save(my_name)
            my_name = "DOCUMENT_%d.docx" % (j)
            j += 1
            my_out = docx.Document()
            out_s = my_out.sections[0]
            # Assumes no changes in section properties
            # (i.e., paragraph i section properties are all the same)
            match_sect_properties(d.sections[0], out_s)
        if my_out:
            # Create a new empty paragraph, then iterate over paragraph runs
            # NOTE: every paragraph has at least one run
            out_p = my_out.add_paragraph(text="", style=para.style.name)
            for p_run in para.runs:
                out_r = out_p.add_run(
                    text = p_run.text, style = p_run.style.name)
                match_char_style(p_run, out_r)
    # Save the last chapter
    my_out.save(my_name)


##############################################################################
# MAIN
##############################################################################
# User inputs:
my_dir = "examples"     # where to look for the input document
my_key = "example-1"    # keyword for finding the right input document
br_style = "Heading1" # the paragraph style used to parse the input document

# Step 1: find the input word files
my_files = find_word_files(my_dir, my_key)
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
    if br_style in my_styles.keys():
        parse_file(my_doc, br_style)
