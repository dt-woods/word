#!/usr/bin/env python3
#
# docx_merger.py
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
import docx

from docx_utils import find_word_files


##############################################################################
# MAIN
##############################################################################
# User inputs:
my_dir = "."           # where to look for the input document
my_key = "DOCUMENT"    # keyword for finding the right input document

# Step 1: find the input word files
my_files = find_word_files(my_dir, my_key)
num_files = len(my_files)
if num_files == 0:
    print("Failed to find any files. "
          "Please update path and keywords and try again.")
else:
    for my_file in my_files:
        print(my_file)
