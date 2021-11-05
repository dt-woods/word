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
#
# TODO: create a "get_image" function that unzips an image from the .docx
#       and puts it in the temporary directory, and returns the path to the
#       extracted image; to be used in Python's docx run.add_picture()
#
##############################################################################
# IMPORT NECESSARY MODULES
##############################################################################
import os
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
    History:  Version 0
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
def get_docx_namespace(doc_path):
    """
    Name:     get_docx_namespace
    Inputs:   str, path to .docx
    Returns:  dict, namespaces and their URIs
    Features: Returns a dictionary of namespaces from an .docx XML
    Depends:  open_docxml
    """
    xml_str = open_docxml(doc_path, False)
    ns_dict = {}
    my_ns = re.findall(r'xmlns:(\S+)="(\S+)"', xml_str)
    for ns in my_ns:
        k, v = ns
        ns_dict[k] = v
    return ns_dict


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


def search_for_attr(my_et, my_attr, is_found=False):
    """
    Features: Recursive search of ET until element is found with attribute
    """
    if not is_found:
        for child in my_et:
            for k, v in child.attrib.items():
                if my_attr in k:
                    is_found = True
                    print(v)
            search_for_attr(child, my_attr, is_found)


##############################################################################
# MAIN
##############################################################################
if __name__ == '__main__':
    import json
    docx_file = "example-2.docx"
    docx_path = os.path.join("examples", docx_file)
    dp = DocxPics(docx_path)
    #print(json.dumps(dp.paras, sort_keys=True, indent=2))
    print(json.dumps(dp.imagemap, indent=2))
