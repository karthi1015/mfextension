# -*- coding: utf-8 -*-
__title__ = 'MF Lab Test 2'
__doc__ = """MF Lab Test 2
"""

__helpurl__ = ""
import os
import sys
## Add Path to MF Functions folder 	

# path_to_this_script = os.path.realpath(__file__)

# current_path = path_to_this_script.split("MFTools.extension")[0]

# print current_path + " MFTools.extension\MFTools.tab\MF_functions"

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_HeaderStuff import *

#print "all ok"

import listview

import treeview

import combobox

import colour_dialog

import itertools
import xaml_test

