# -*- coding: utf-8 -*-
__title__ = 'Export View Templates'
__doc__ = """Export View Templates
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

viewTemplateData.append(["View Template", "Element ID", "View Template Name", "View Type", "Sub-Discipline"])



log = []

def MF_GetParameterValueByName(el, paramName):
	for param in el.Parameters:
		if param.IsShared and param.Definition.Name == paramName:
		#if param.Definition.Name == paramName:
			paramValue = el.get_Parameter(param.GUID)
			return paramValue.AsString()
		elif param.Definition.Name == paramName: #handle project parameters?
			#paramValue = el.get_Parameter(paramName)
			return param.AsValueString()

		
def MF_SetParameterByName(el, paramName, value):
	for param in el.Parameters:
		if param.IsShared and param.Definition.Name == paramName:
			param.Set(value)
					
#print("Exporting View Templates")		

for v in allViews:
	if v.IsTemplate:
		vt = doc.GetElement(v.Id)
		
		viewType = MF_GetParameterValueByName(vt, "View type")
		
		#subDiscipline = MF_GetParameterValueByName(vt, "Sub-Discipline")
		
		subDiscipline = vt.LookupParameter("Sub-Discipline").AsString()
		
		
		#viewType = ' - '
		viewTemplateData.append([v, v.Id, v.Name, viewType, subDiscipline])
		viewTemplates.append(v)

# for item in viewTemplateData:
        # print ' \t\t\t\t\t '.join(str(x) for x in item[1:])
# print("Done")	

t = Transaction(doc, 'Write Excel.')
 
t.Start()
 
excel = Excel.ApplicationClass()   

from System.Runtime.InteropServices import Marshal

excel = Marshal.GetActiveObject("Excel.Application")

excel.Visible = True
excel.DisplayAlerts = False   

filename = r'C:\Users\e.green\Desktop\VTNamesExport.xlsx'

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

filename = desktop + '\VTNamesExport.xlsx'

#Workbooks
#if workbook exists, try to open it
try:
	workbook = excel.Workbooks.Open(filename)
except:
	# if not, create a new one
	workbook = excel.Workbooks.Add()
	#save it with the desired name
	workbook.SaveAs(filename)

	# oopen it
	workbook = excel.Workbooks.Open(filename)

# finding a workbook that's already open

workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename]
if workbooks:
    workbook = workbooks[0]


ws = workbook.Worksheets[1]


def ColIdxToXlName(idx):
    if idx < 1:
        raise ValueError("Index is too small")
    result = ""
    while True:
        if idx > 26:
            idx, r = divmod(idx - 1, 26)
            result = chr(r + ord('A')) + result
        else:
            return chr(idx + ord('A') - 1) + result



exportData = viewTemplateData


lastRow = len(exportData)
lastColumn = len(exportData[0])

totalColumns = len(max(exportData,key=len))

#totalColumns = 11 # temporary hack

lastColumn = totalColumns

lastColumnName = ColIdxToXlName(totalColumns)

xlrange = ws.Range["A1", lastColumnName+str(lastRow)]

a = Array.CreateInstance(object, len(exportData),totalColumns)

exportData[1:] = sorted(exportData[1:],key=lambda x: x[1]) 

i = 0


while i < lastRow:
	j = 0
	while j < totalColumns:
	
		a[i,j] = exportData[i][j]
		j += 1
	
	
	i += 1

xlrange.Value2 = a 

ws.Range(ws.Cells(1,1), ws.Cells(1,lastColumn)).Font.Bold = True
ws.Range(ws.Cells(1,1), ws.Cells(lastRow,lastColumn)).Columns.AutoFit()
ws.Range(ws.Cells(1,1), ws.Cells(lastRow,lastColumn)).AutoFilter()

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
#print '<table>'
for item in importData:
	#print '<tr><td>'	
	#print '</td><td>'.join(str(x) for x in item[1:])
	print '\t\t\t\t\t\t\t\t\t'.join(str(x) for x in item[1:])
	#print '</td></tr>'
#print '</table>'
print("Done")		
 
t.Commit()
 
#__window__.Close()


