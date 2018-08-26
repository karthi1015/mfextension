import clr
import os
import os.path as op
import pickle as pl

import sys
import subprocess
import time

import struct

import rpw
from rpw import doc, uidoc, DB, UI

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

import System

from System import Array

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

try:
    from pyrevit.versionmgr import PYREVIT_VERSION
except:
    from pyrevit import versionmgr
    PYREVIT_VERSION = versionmgr.get_pyrevit_version()

pyRevitNewer44 = PYREVIT_VERSION.major >= 4 and PYREVIT_VERSION.minor >= 5

if pyRevitNewer44:
    from pyrevit import script, revit, forms
    from pyrevit.forms import *
    output = script.get_output()
    logger = script.get_logger()
    linkify = output.linkify
    from pyrevit.revit import doc, uidoc, selection
    selection = selection.get_selection()

else:
    from scriptutils import logger
    from scriptutils.userinput import SelectFromList, SelectFromCheckBoxes
    from revitutils import doc, uidoc, selection
	


def MF_GetParameterValueByName(el, paramName):
	for param in el.Parameters:
		if param.IsShared and param.Definition.Name == paramName:
		#if param.Definition.Name == paramName:
			paramValue = el.get_Parameter(param.GUID)
			return paramValue.AsString()
		elif param.Definition.Name == paramName: #handle project parameters?
			#paramValue = el.get_Parameter(paramName)
			return param.AsValueString()	
			
def MF_GetTypeParameterValueByName(elType, paramName):
	
	
	
	for param in elType.Parameters:
		if param.IsShared and param.Definition.Name == paramName:
		#if param.Definition.Name == paramName:
			paramValue = el.get_Parameter(param.GUID)
			return paramValue.AsString()
		elif param.Definition.Name == paramName: #handle project parameters?
			#paramValue = el.get_Parameter(paramName)
			return param.AsValueString()				

    			