#!/usr/bin/env python3
#
# docx_utils.py
#
# VERSION: 0.1.0
# UPDATED: 2021-10-01
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
    return sorted(my_files)


def list_paragraph_styles(d):
    """
    Name:     list_paragraph_styles
    Inputs:   docx.document.Document, open word document
    Output:   dict, style_id (keys) with name and counts (keys) found
    Features: Returns a list of all the paragraph styles found in given doc
    """
    style_dict = {}
    para_num = len(d.paragraphs)
    for i in range(para_num):
        para = d.paragraphs[i]
        if para.style.style_id not in style_dict:
            style_dict[para.style.style_id] = {
                'name': para.style.name,
                'count': 1
            }
        else:
            style_dict[para.style.style_id]['count'] += 1
    return style_dict


def match_char_style(a, b):
    """
    Name:     match_char_style
    Inputs:   - docx.text.run.Run, run from original document (a)
              - docx.text.run.Run, run for new document (b)
    Outputs:  None
    Features: Edits a run object with the same font styles
    TODO:     - go into the font setting of the paragraph run and match at
                this level (see references)

    References:
    - https://python-docx.readthedocs.io/en/latest/api/text.html#font-objects
    """
    # First pass, just match "bold," "italic," and "underline" at run-level
    if a.bold:
        b.bold = True
    if a.italic:
        b.italic = True
    if a.underline:
        b.underline = True


def match_sect_properties(a, b):
    """
    Name:     match_sect_properties
    Inputs:   - docx.section.Section, given section to copy from (a)
              - docx.section.Section, section to apply properties to (b)
    Outputs:  None
    Features: Edits a section object to match the properties of given section

    Reference:
    https://python-docx.readthedocs.io/en/latest/user/sections.html
    """
    b.page_width = a.page_width
    b.page_height = a.page_height
    b.orientation = a.orientation
    b.left_margin = a.left_margin
    b.top_margin = a.top_margin
    b.right_margin = a.right_margin
    b.bottom_margin = a.bottom_margin
