# -*- coding: utf-8 -*-
__title__ = 'Import Legend Line Styles'
__doc__ = """Import Legend Line Styles
"""

__helpurl__ = ""

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
    from pyrevit import script, revit
    from pyrevit.forms import SelectFromList, SelectFromCheckBoxes
    output = script.get_output()
    logger = script.get_logger()
    linkify = output.linkify
    from pyrevit.revit import doc, uidoc, selection
    selection = selection.get_selection()

else:
    from scriptutils import logger
    from scriptutils.userinput import SelectFromList, SelectFromCheckBoxes
    from revitutils import doc, uidoc, selection


	


allViews = FilteredElementCollector(doc).OfClass(View).ToElements()

viewTemplates = []
viewTemplateData = []

viewTemplateData.append(["View Template", "Element ID", "View Template Name", "View Type"])



log = []

def MF_GetParameterValueByName(el, paramName):
	for param in el.Parameters:
		if param.IsShared and param.Definition.Name == paramName:
			paramValue = el.get_Parameter(param.GUID)
			return paramValue.AsString()
	        
def MF_SetParameterByName(el, paramName, value):
	for param in el.Parameters:
		#if param.IsShared and param.Definition.Name == paramName:
		if param.Definition.Name == paramName:
			param.Set(value)
					


t = Transaction(doc, 'Import Legend Line Styles From Excel')
 
t.Start()
 


excel = Excel.ApplicationClass()   

from System.Runtime.InteropServices import Marshal

excel = Marshal.GetActiveObject("Excel.Application")

excel.Visible = True
excel.DisplayAlerts = False   

##### This needs to be user independent... #############
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

filename = desktop + '\LegendLineStyles.xlsx'
#Workbooks

# creating a new one
#workbook = excel.Workbooks.Add()

# opening a workbook
#workbook = excel.Workbooks.Open(filename)

# finding a workbook that's already open
#System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")

workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename]
if workbooks:
	workbook = workbooks[0]
else:
	workbook = excel.Workbooks.Open(filename)


ws = workbook.Worksheets[1]

#lastRow = len(viewTemplateData)
#lastColumn = len(viewTemplateData[0])

#lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row



categories = doc.Settings.Categories;

lineCat = categories.get_Item(BuiltInCategory.OST_Lines )


##### Temporary Fix #############


lastRow = 38
lastColumn = 8

def rgb2hex(rgb):
    return struct.pack('BBB',*rgb).encode('hex')

importData = []

i=1
while i <= lastRow:
	importRow = []
	j=1
	while j <= lastColumn:
		importRow.append(ws.Cells(i,j).Text)
		j += 1
	importData.append(importRow)		
	i += 1

for item in importData[1:]:

	name = item[0]
	
	lw = 1
	
	if int(item[4]) > 0 :
		lw = int(item[4])
	
	
	lineRGB= item[6]
	
	lrgb = map(str.strip, lineRGB.split('_'))
	
	#lCol = rgb2hex( (lrgb[0], lrgb[1], lrgb[2] ) )
	
	
	
	
	newLineStyleCat = categories.NewSubcategory( lineCat, "MXF_Legend__" + name )
	  
	try:
		newLineStyleCat.SetLineWeight( lw, GraphicsStyleType.Projection )
	except Exception as e:
		print str(e)
	
	# r = int(lrgb[0]).to_bytes(1, byteorder='big', signed=True) 
	# g = int(lrgb[1]).to_bytes(1, byteorder='big', signed=True) 
	# b = int(lrgb[2]).to_bytes(1, byteorder='big', signed=True) 
	
	r = (int(lrgb[0]))
	g = (int(lrgb[1]))
	b = (int(lrgb[2]))
	
	newLineStyleCat.LineColor = Color( r, g, b )
	
	#print name + '\t\t\t\t\t\t\t\t' + lineWeight + '\t\t\t' + lineColor
	
	

print("Done")		


 
t.Commit()
 
#__window__.Close()


