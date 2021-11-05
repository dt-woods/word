#!/usr/bin/env python3
#
# docx_utils.py
#
# VERSION: 1.2.1
# UPDATED: 2021-11-05
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
import zipfile


##############################################################################
# FUNCTIONS
##############################################################################
def find_files(d, k=""):
    """
    Name:     find_files
    Inputs:   - str, file path (d)
              - str, keyword(s) in the file to search (k)
    Outputs:  List
    Features: Searches the given directory for word files
    """
    my_files = glob.glob(os.path.join(d, k))
    return sorted(my_files)


def find_images(doc_path, doc_obj):
    """
    Input: - str, document file path (d)
           - str, the whole-document blob decoded for utf-8

    Attempt to read the document.part.blob to find where the in-line images
    are located. For example, search for:
        * <w:drawing>.*</w:drawing>
        * <wp:inline>.*</wp:inline>
        * <wp:docPr id="1" name="Picture 1"/>
        * <pic:cNvPr id="0" name="Picture 3"/>
        * <a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"
    Find the text immediately before and after and attempt to find the
    paragraph that has this text (or, if it's between paragraphs). Then,
    try to unzip the .docx to find the embedded images (either by name or
    by URI; see above). Then go back to the paragraph and add the image to
    the run (if in-line) or add the image after the paragraph (if stand alone).
    IN-PROGRESS
    """
    # Open docx and find images:
    z = zipfile.ZipFile(d, 'r')
    for zfile in z.filelist:
        if 'image' in zfile.filename:
            print(zfile.filename)
    # From example-2.docx, we see that the three images are stored in the
    # word/media sub-directory (i.e., image1.png, image2.png, image3.png).
    # TODO: extract these images to binary and add them to the document.
    #
    # Find where to put them
    my_blob = doc_obj.part.blob.decode('utf-8')
    # TODO: find images within the blob (maybe xml tree?)

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
    Features: Edits a run object with the same font styles, including:
              - bold, italic, underline
              - all_caps, small_caps, strike, double_strike, outline,
              - superscript, subscript
    TODO:     - does not handle color, highlight color, or size

    References:
    - https://python-docx.readthedocs.io/en/latest/api/text.html#font-objects
    """
    # Second pass, go into the font styles
    # grab all settable properties from font class,
    # extract the value (allow AttributeError to pass)
    # use execute to set a's parameter value to b
    properties = [i for i in dir(a.font) if not i.startswith("_")]
    for p in properties:
        try:
            v = eval("a.font.{}".format(p))
        except AttributeError:
            pass
        else:
            if isinstance(v, bool):
                exec("b.font.{} = {}".format(p, v))
    #

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
