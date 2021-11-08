#!/usr/bin/env python3
#
# docx_utils.py
#
# VERSION: 2.0.0
# UPDATED: 2021-11-07
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
from zipfile import ZipFile
import xml.etree.ElementTree as ElementTree


##############################################################################
# CLASSES
##############################################################################
class DocxPics(object):
    """
    Name:     DocxPics
    Features: Class for organizing references to images found within a .docx
    History:  Version 1
    """
    # \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    # Class Initialization
    # ////////////////////////////////////////////////////////////////////////
    def __init__(self, doc_path):
        """
        Name:     DocxPics.__init__
        Inputs:   str, path to a docx document
        Features: Initializes the DocxPics class
        """
        # Initialize the class parameters
        self.num_paras = 0     # number of paragraphs in .docx
        self.num_images = 0    # number of images in .docx
        self.paralist = []     # list of paragraph indices containing an image
        self.paraIdList = []   # the paraIds associated with paragraphs above
        self.paras = {}        # dictionary containing paragraph information
        self.xml = None        # documant.xml as string
        self.xmlet = None      # ElementTree of document.xml
        self.rel = None        # document.xml.rel as string
        self.relet = None      # ElementTree of document.xml.rel
        self.namespace = {}    # namespace dictionary
        self.imagemap = {}     # map between relationship IDs and image paths
        self.imID = None       # temporary image ID

        # Check that input document is valid
        if os.path.isfile(doc_path):
            self.docx = doc_path
            self.open_docxml()
            self.get_docx_namespace()
            self.find_images()
        else:
            self.docx = None
            raise OSError("File %s does not exist!" % (doc_path))

    # \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    # Class Function Definitions
    # ////////////////////////////////////////////////////////////////////////
    def find_images(self):
        """
        Name:     DocxPics.count_paragraphs
        Inputs:   None
        Outputs:  None
        Features: Finds and counts paragraphs, runs, and images in .docx
        Depends:  - map_images
                  - search_for_attr
        """
        self.map_images()
        self.paralist = []
        self.paraIdList = []
        self.paras = {}
        self.num_images = 0
        if self.xmlet:
            my_paras = self.xmlet[0].findall("w:p", self.namespace)
            self.num_paras = len(my_paras)
            for i in range(self.num_paras):
                para = my_paras[i]
                # Get paragraph ID:
                para_id = ''
                for k, v in para.attrib.items():
                    if 'paraId' in k:
                        para_id = v

                # Find all runs in paragraph
                runs = para.findall("w:r", self.namespace)
                num_runs = len(runs)
                for j in range(num_runs):
                    run = runs[j]
                    draws = run.findall("w:drawing", self.namespace)
                    num_draws = len(draws)
                    self.num_images += num_draws
                    if num_draws > 0:
                        self.paralist.append(i)
                        self.paraIdList.append(para_id)
                        for draw in draws:
                            self.imID = None
                            self.search_for_attr(draw, 'embed')
                            draw_path = ""
                            if self.imID in self.imagemap.keys():
                                draw_path = self.imagemap[self.imID]
                            if i in self.paras.keys():
                                self.paras[i]['num_images'] += 1
                                self.paras[i]['runs'][j] = {
                                    'imgID': str(self.imID),
                                    "imgPath": draw_path
                                }
                            else:
                                self.paras[i] = {
                                    'paraId': para_id,
                                    'num_run': num_runs,
                                    'num_images': 1,
                                    'runs': {
                                        j: {
                                            'imgID': str(self.imID),
                                            'imgPath': draw_path
                                        }
                                    }
                                }

    def get_docx_namespace(self):
        """
        Name:     DocxPics.get_docx_namespace
        Inputs:   None
        Returns:  None
        Features: Creates a dictionary of namespaces from an .docx XML
        """
        if self.xml:
            self.namespace = {}
            my_ns = re.findall(r'xmlns:(\S+)="(\S+)"', self.xml)
            for ns in my_ns:
                k, v = ns
                self.namespace[k] = v

    def map_images(self):
        """
        Name:     DocxPics.map_images
        Inputs:   None
        Outputs:  None
        Features: Maps image relationship IDs to their paths within .docx
        Depends:  re
        """
        self.imagemap = {}
        if self.xml:
            my_blips = re.findall('a:blip r:embed="(\S+)"', self.xml)
            for my_rel in self.relet:
                is_img = False
                img_id = ""
                img_path = ""
                for k, v in my_rel.attrib.items():
                    if k == "Id" and v in my_blips:
                        is_img = True
                        img_id = v
                if is_img:
                    img_path = my_rel.attrib['Target']
                    self.imagemap[img_id] = img_path

    def open_docxml(self):
        """
        Name:     DocxPics.open_docxml
        Inputs:   None
        Returns:  None
        Features: Reads the XML from document.xml within a .docx and creates
                  ElementTree objects from string
        Depends:  - ElementTree
                  - zipfile.ZipFile
        """
        self.xml = ""
        self.rel = ""
        if self.docx:
            my_doc = ""
            my_rel = ""
            my_zip = ZipFile(self.docx, 'r')
            for zc in my_zip.namelist():
                if "document.xml" in zc and zc.endswith('rels'):
                    my_rel = zc
                if "document.xml" in zc and not zc.endswith('rels'):
                    my_doc = zc

            if my_doc != "":
                with my_zip.open(my_doc) as f:
                    for d in f:
                        self.xml = "".join([self.xml, d.decode("utf-8")])
                self.xmlet = ElementTree.fromstring(self.xml)

            if my_rel != "":
                with my_zip.open(my_rel) as f:
                    for d in f:
                        self.rel = "".join([self.rel, d.decode("utf-8")])
                self.relet = ElementTree.fromstring(self.rel)

    def search_for_attr(self, my_et, my_attr, is_found=False):
        """
        Name:     DocxPics.search_for_attr
        Inputs:   - xml.etree.ElementTree.Element
                  - str, attribute name (my_attr)
                  - bool, if attribute has been found
        Outputs:  str, value of the search attribute
        Features: Recursive search of ET until element is found with attribute
        """
        if not is_found:
            for child in my_et:
                for k, v in child.attrib.items():
                    if my_attr in k:
                        is_found = True
                        self.imID = v
                self.search_for_attr(child, my_attr, is_found)


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
