# -*- coding: utf-8 -*-
__title__ = 'Export Printing View List'
__doc__ = """Export Printing View List
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


printingViews = []
#does this keep re collecting?
printingViews.append(["View Name", "View Id", "Level", "Scope Box", "Sub Discipline",  "Uniclass Group", "Uniclass - Sub group", "Uniclass - Section" , "Uniclass - Object" ])

allViews = FilteredElementCollector(doc).OfClass(View).ToElements()
for v in allViews:
	if "500" in v.Name and not v.IsTemplate:
		subDisicipline = " - "
		uniclassGroup = " - "
		uniclassSubGroup = " - "
		uniclassSection= " - "
		uniclassObject = " - "
		scopeBox = " - "
		try:
			subDisicipline = v.LookupParameter("Sub-Discipline").AsString()
			uniclassGroup = v.LookupParameter("Uniclass Ss - Group").AsString()
			uniclassSubGroup = v.LookupParameter("Uniclass Ss - Sub group").AsString()
			uniclassSection= v.LookupParameter("Uniclass Ss - Section").AsString()
			uniclassObject = v.LookupParameter("Uniclass Ss - Object").AsString()
			scopeBox = v.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsValueString()
		except Exception as e:
			print str(e)
			pass
			
		
		printingViews.append([v.Name, v.Id, v.GenLevel.Name, scopeBox, subDisicipline, uniclassGroup, uniclassSubGroup, uniclassSection, uniclassObject ])
		
		


# vpLocation = IN[4]

#vpX = vpLocation.X
#vpY = vpLocation.Y

log = []

viewports = []


t = Transaction(doc, 'Write Excel.')
 
t.Start()
 


excel = Excel.ApplicationClass()   

from System.Runtime.InteropServices import Marshal

excel = Marshal.GetActiveObject("Excel.Application")

excel.Visible = True
excel.DisplayAlerts = False   

filename = 'C:\Users\e.green\Desktop\ViewListDataExport.xlsx'
#Workbooks

# creating a new one
workbook = excel.Workbooks.Add()

# opening a workbook
#workbook = excel.Workbooks.Open(filename)

System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo("en-US")
#workbook = excel.Workbooks.Open(r"C:\Users\e.green\Desktop\VTExport.xlsx")
#workbook = excel.Workbooks.Open(r"Master View Template Settings.xlsm")

# finding a workbook that's already open

# workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename]
# if workbooks:
    # workbook = workbooks[0]


#ws = workbook.Worksheets[1]


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


ws = workbook.Worksheets.Add()

#############################################################################
### View List

ws.Name = "View List"

exportData = printingViews


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

#ws.Range(ws.Cells(1,1), ws.Cells(1,lastColumn)).Font.Bold = True
ws.Range(ws.Cells(1,1), ws.Cells(lastRow,4)).Columns.AutoFit()
ws.Range(ws.Cells(1,6), ws.Cells(lastRow,lastColumn)).Columns.AutoFit()
ws.Range(ws.Cells(1,1), ws.Cells(lastRow,lastColumn)).AutoFilter()

ws.Activate

t.Commit()
 
#__window__.Close()