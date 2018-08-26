# -*- coding: utf-8 -*-
__title__ = 'Export Printing Sheet List'
__doc__ = """Export Printing Sheet List
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


import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")
	
from MF_ExcelOutput import *	

sheet_fields = [
	"01. Project Code",
	"02. Originator",
	"03. Volume",
	"04. Level",
	"05. Document Type",
	"06. Role / Discipline",
	"07. Classification",
	
	"08. Number"

]	
	

printingViews = []
printingViews.append([
	"01. Project Code",
	"02. Originator",
	"03. Volume",
	"04. Level",
	"05. Document Type",
	"06. Role / Disicipline",
	"07. Classification",
	
	"08. Number",
	
	("Sheet Number"), 
	("Sheet Name"), 
	
	
	
	("Sheet ID"), 
	('Sub-Discipline - Sheet'), 
	('View Type - Sheet'), 
	("View Name"), 
	('View ID'),
	('Sub-Disicpline - View'), 
	('View Type'), 
	('View Level'), 
	('View Uniclass Group'), 
	('View Uniclass Sub Group'),
	('Location on Sheet'), 
	('Sheet Outline')
])

# selSheets = forms.select_sheets(title='Select Target Sheets',
                                # button_name='Select Sheets')
								
								


       								


allSheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()



#select multiple
selected = []
return_options = SelectFromCheckBoxes.show(
			sorted([SheetOption(x) for x in allSheets],
				   key=lambda x: x.number),
			title="Select Sheets",
			button_name="Select Sheets",
			width=800)
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]


sheetList = list(selected)

for s in sheetList:
	
	sheetSubDiscipline = s.LookupParameter("Sub-Discipline").AsString()
	sheetViewType = s.LookupParameter("View type").AsString()
	
	if s.GetAllPlacedViews():
		placedViews = s.GetAllPlacedViews()
		
		pvIds = []
		for pvID in placedViews: ## edit this to get distinguish between floorplanviews, legends etc
			pvIds.append(pvID)
			firstPlacedView = doc.GetElement(pvIds[0])
			viewSubDiscipline = firstPlacedView.LookupParameter("Sub-Discipline").AsString()
			# get all views placed on sheet
		#	printingViews.append([s.SheetNumber, s.Name, s.Id, sheetSubDiscipline, sheetViewType, doc.GetElement(pvID).Name,    viewSubDiscipline, str(doc.GetElement(pvID).ViewType)])
		# gets first View placed on the sheet
	#	printingViews.append([s.SheetNumber, s.Name, s.Id, sheetSubDiscipline, sheetViewType, firstPlacedView.Name,  viewSubDiscipline])
		
	if s.GetAllViewports():
		viewports = s.GetAllViewports()
		
		sheetOutline = "TBC"
		viewLevel = " - " 
		
		
		sheet_field_values = []
		for field in sheet_fields:
			field_value = ' - '
			try:
				field_value = s.LookupParameter(field).AsString()
			except Exception as e:
				print "Error: " + field + " - " + str(e)
				pass
			sheet_field_values.append(field_value)
		
		#speed up
		#sheetOutline = str(s.Outline.Max) + ',' + str(s.Outline.Min)
		for vp in viewports:
			vPort = doc.GetElement(vp)
			view = doc.GetElement(vPort.ViewId)
			viewSubDiscipline = view.LookupParameter("Sub-Discipline").AsString()
			viewUniclassGroup = ' - '
			viewUniclassSubGroup  = ' - '
			try:
				viewUniclassGroup =  view.LookupParameter("Uniclass Ss - Group").AsString()
				viewUniclassSubGroup = view.LookupParameter("Uniclass Ss - Sub group").AsString()
			except:
				pass
			try:
				viewLevel =  view.GenLevel.Name
				
			except:
				pass
			
			location = "TBC"
			# speed up
			#location = str(vPort.GetBoxCenter() )
			
			viewType = str(view.ViewType)
			
			printingViewRow = sheet_field_values
			
			
			
			printingViewRow.extend([s.SheetNumber, s.Name, s.Id, sheetSubDiscipline, sheetViewType, view.Name,  view.Id,  viewSubDiscipline, str(doc.GetElement(vPort.ViewId).ViewType), viewLevel , viewUniclassGroup, viewUniclassSubGroup,  location, sheetOutline ])
			
			printingViews.append(printingViewRow)
			
			# printingViews.append([s.SheetNumber, s.Name, s.Id, sheetSubDiscipline, sheetViewType, view.Name,  view.Id,  viewSubDiscipline, str(doc.GetElement(vPort.ViewId).ViewType), viewLevel , viewUniclassGroup, viewUniclassSubGroup,  location, sheetOutline ])
			
	else:
		
		printingViewRow = sheet_field_values
		printingViewRow.extend([s.SheetNumber, s.Name, s.Id, sheetSubDiscipline, sheetViewType, '-', '-', '-', '-', '-', '-', '-', '-', '-' ])
		
		printingViews.append(printingViewRow)
		
		# printingViews.append([s.SheetNumber, s.Name, s.Id, sheetSubDiscipline, sheetViewType, '-', '-', '-', '-', '-', '-', '-', '-', '-' ])
		

# vpLocation = IN[4]

#vpX = vpLocation.X
#vpY = vpLocation.Y

log = []

viewports = []


# t = Transaction(doc, 'Export Printing Sheet List')
 
# t.Start()
 


excel = Excel.ApplicationClass()   

from System.Runtime.InteropServices import Marshal

excel = Marshal.GetActiveObject("Excel.Application")

excel.Visible = True
excel.DisplayAlerts = False   

#filename = 'C:\Users\e.green\Desktop\SheetListDataExport.xlsx'

#desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

filename =  doc.Title +  "_ SheetListDataExport.xlsx"


# #Workbooks
# #if workbook exists, try to open it
# try:
	# workbook = excel.Workbooks.Open(filename)
# except:
	# # if not, create a new one
	# workbook = excel.Workbooks.Add()
	# #save it with the desired name
	# workbook.SaveAs(filename)

	# # oopen it
	# workbook = excel.Workbooks.Open(filename)

# # finding a workbook that's already open

# workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename]
# if workbooks:
    # workbook = workbooks[0]


# ws = workbook.Worksheets[1]


# def ColIdxToXlName(idx):
    # if idx < 1:
        # raise ValueError("Index is too small")
    # result = ""
    # while True:
        # if idx > 26:
            # idx, r = divmod(idx - 1, 26)
            # result = chr(r + ord('A')) + result
        # else:
            # return chr(idx + ord('A') - 1) + result


# ws = workbook.Worksheets.Add()


#############################################################################
### View List



#exportData = printingViews

ws = MF_WriteToExcel(filename, "Filter Overrides", printingViews)


# lastRow = len(exportData)
# #lastColumn = len(exportData[1])

# totalColumns = len(max(exportData,key=len))

# #totalColumns = 11 # temporary hack

# lastColumn = totalColumns

# lastColumnName = ColIdxToXlName(totalColumns)

# xlrange = ws.Range["A1", lastColumnName+str(lastRow)]

# a = Array.CreateInstance(object, len(exportData),totalColumns)

# exportData[1:] = sorted(exportData[1:],key=lambda x: x[1]) 

# i = 0 # ignore header row


# while i < lastRow:
	# j = 0
	# while j < totalColumns:
	
		# a[i,j] = exportData[i][j]
		# j += 1
	
	
	# i += 1

# xlrange.Value2 = a 

# ws.Range(ws.Cells(1,1), ws.Cells(1,lastColumn)).Font.Bold = True
# ws.Range(ws.Cells(1,1), ws.Cells(lastRow,4)).Columns.AutoFit()
# ws.Range(ws.Cells(1,6), ws.Cells(lastRow,lastColumn)).Columns.AutoFit()
# ws.Range(ws.Cells(1,1), ws.Cells(lastRow,lastColumn)).AutoFilter()

# ws.Activate

# t.Commit()
 
#__window__.Close()