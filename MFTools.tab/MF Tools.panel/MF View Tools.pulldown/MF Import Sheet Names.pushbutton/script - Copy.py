# -*- coding: utf-8 -*-
__title__ = 'Import Sheet Names and Numbers from Excel'
__doc__ = """Import Sheet Names and Numbers from Excel List
"""

__helpurl__ = ""

import clr
import os
import os.path as op
import pickle as pl

import sys
import subprocess
import time

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
					
#print("Exporting View Templates")		

for v in allViews:
	if v.IsTemplate:
		vt = doc.GetElement(v.Id)
		
		viewType = MF_GetParameterValueByName(vt, "View type")
		

		
		
		#viewType = ' - '
		viewTemplateData.append([v, v.Id, v.Name, viewType])
		viewTemplates.append(v)

# for item in viewTemplateData:
        # print ' \t\t\t\t\t '.join(str(x) for x in item[1:])
# print("Done")	

t = Transaction(doc, 'Import Sheet Names from Excel')
 
t.Start()
 


excel = Excel.ApplicationClass()   

from System.Runtime.InteropServices import Marshal

excel = Marshal.GetActiveObject("Excel.Application")

excel.Visible = True
excel.DisplayAlerts = False   

#filename = 'C:\Users\e.green\Desktop\SheetListDataExport.xlsx'

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

filename = desktop + '\SheetListDataExport.xlsx'
#Workbooks


# finding a workbook that's already open
System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")

workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename]
if workbooks:
	workbook = workbooks[0]
else:
	workbook = excel.Workbooks.Open(filename)



ws = workbook.ActiveSheet

#lastRow = len(viewTemplateData)
#lastColumn = len(viewTemplateData[0])

#lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastRow = 77 ######################################################
lastColumn = 7 ######################################################



importData = []

i=1
while i < lastRow:
	importRow = []
	j=1
	while j < lastColumn:
		importRow.append(ws.Cells(i,j).Text)
		j += 1
	importData.append(importRow)		
	i += 1

for item in importData[1:]:

	#elemId = Autodesk.Revit.DB.ElementId(int(item[0]))
	
	if item[2]:
		el = doc.GetElement(ElementId(int(item[2])))
		#el = doc.GetElement(ElementId(item[2]))
	
		oldName = el.Name
		
		if oldName != item[1]:
		
			try:
				el.Name = item[1]
				print oldName + ' renamed to ' + item[1]
			except Exception as e:
				print str(e)
				
		
		#oldNumber = el.SheetNumber
		
		# this should work.. 
		#oldNumber = el.get_Parameter(BuiltInParameter.SHEET_NUMBER).AsString()
		
		#if true:
		#if oldNumber != item[0]:
		try:
			el.get_Parameter(BuiltInParameter.SHEET_NUMBER).Set(item[0])
			#print oldNumber + ' renumbered to ' + item[0]
		except Exception as e:
			print str(e)
		
		
		
		#MF_SetParameterByName(el, "Sheet Number", item[0])
		MF_SetParameterByName(el, "Sub-Discipline", item[3])
		MF_SetParameterByName(el, "View type", item[4])
		
	
	#print '\t\t\t\t\t\t\t\t\t'.join(str(x) for x in item)

print("Done")		


 
t.Commit()
 
#__window__.Close()


