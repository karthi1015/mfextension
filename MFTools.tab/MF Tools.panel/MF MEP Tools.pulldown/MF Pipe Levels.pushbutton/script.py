# -*- coding: utf-8 -*-
__title__ = 'MF Pipe Levels'
__doc__ = """Pipe Magic
"""

__helpurl__ = ""

import clr
import os
import os.path as op
import pickle as pl

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

import System

from System import Array
from System import Enum

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

try:
    from pyrevit.versionmgr import PYREVIT_VERSION
except:
    from pyrevit import versionmgr
    PYREVIT_VERSION = versionmgr.get_pyrevit_version()

from pyrevit import script
#from pyrevit import scriptutils 
from pyrevit import framework
from pyrevit import revit, DB, UI
from pyrevit import forms

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

output = script.get_output()
output.set_width(1100)



def MF_SetParameterByName(el, paramName, value):
	for param in el.Parameters:
		if param.Definition.Name == paramName:
			param.Set(value)	


			
from Autodesk.Revit.UI.Selection import * 


#selFilter = 
#picked = uidoc.Selection.PickElementsByRectangle()

			
horizontalPipes = []

horizontalPipes.append(["Name", "Reference Level", "System Type", "System Name", "delta Z",  "Offset From Reference Level", "Pipe Level Description" , "Reference Level Elevation", "Pipe Offset" ]  )

	
verticalPipes = []

filter = ElementCategoryFilter(BuiltInCategory.OST_PipeCurves)

pipes = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()

#pipes = list(pipes)[:500]

filter = ElementCategoryFilter(BuiltInCategory.OST_Levels)

levels = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()



curves = []
linePoints = []

def fuzzyMatch(a,b, precision):
	return abs(a - b) <= precision
	
options = [2000, 3000 , 4000, 5000, 6000]

########
## need to handle sloping pipes!!!

# shows the form and returns the selected options
selected_options = SelectFromList.show(options,
			title='Select Height of High level Threshold for Pipes (mm from floor)',
			width=800,
			height=500,
			multiselect=False)
	
highLevelThreshold = selected_options[0]

	
t = Transaction(doc, 'Pipe Levels')

t.Start()

for pipe in pipes:
	
	line = pipe.Location.Curve
	
	curves.append(line)
	
	startPoint = line.GetEndPoint(0)
	endPoint = line.GetEndPoint(1)
	
	pipeSystemType = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
	pipeSystemName = pipe.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
	
	linePoints.append([startPoint, endPoint])
	
	p = 0.5
	
	pipeTop = max(startPoint.Z, endPoint.Z)
	pipeBottom = min(startPoint.Z, endPoint.Z)
	
	pipeRefLevelElevation = pipe.ReferenceLevel.Elevation * 304.8  #convert feet to mm
	
	pipeOffset = float(pipe.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).AsValueString())
	
	pipeOffsetFromReferenceLevel = (pipeOffset - abs(pipeRefLevelElevation)  )
	## try this.. 
	pipeOffsetFromReferenceLevel = pipeOffset
	
	pipeLevelDescription = ""
	
	if pipeOffsetFromReferenceLevel >= highLevelThreshold:
		pipeLevelDescription = "High Level"
	else :
		pipeLevelDescription = ""
	##########################	
	#pipeLevel = pipeOffset - pipeRefLevelElevation
	
	crossesLevel = "Crosses Level(s) :"
	
	for level in levels:
		levelZ = level.ProjectElevation
		
		if pipeTop > levelZ and pipeBottom < levelZ:
			crossesLevel +=  level.Name + " ,  " 
		
	
	if not fuzzyMatch(startPoint.Z, endPoint.Z, p):
		verticalPipes.append(pipe)
		MF_SetParameterByName(pipe,"MF_Offset_Description", crossesLevel)
		#print "Vertical:" + pipe.Name + " Z-offset: " + str(round(abs(endPoint.Z - startPoint.Z), 2)) + " --- Top: " + str(max(startPoint.Z, endPoint.Z)) + " --- Bottom: " +str(min(startPoint.Z, endPoint.Z)) + " --- " + crossesLevel 
	else:
		horizontalPipes.append([pipe.Name, pipe.ReferenceLevel.Name, pipeSystemType, pipeSystemName, str(round(abs(endPoint.Z - startPoint.Z),2)),  str(pipeOffsetFromReferenceLevel ), str(pipeLevelDescription) , str(pipeRefLevelElevation) , str(pipeOffset)]  )
		#print "Horizontal:" + pipe.Name + " Z-offset: " + str(round(abs(endPoint.Z - startPoint.Z),2)) + " Pipe Offset from Level: " + str(pipeOffsetFromReferenceLevel ) + "mm  -  " + str(pipeLevelDescription) +" mm -  Ref Level: " + str(pipeRefLevelElevation)  +" mm -  " 
		MF_SetParameterByName(pipe,"MF_Offset_Description", pipeLevelDescription)
	

t.Commit()
	
for level in levels:
	print level.Name + " ---- Elevation: " + str(level.ProjectElevation)
	
	
OUT = pipes, curves, linePoints, verticalPipes

excel = Excel.ApplicationClass()   

from System.Runtime.InteropServices import Marshal

excel = Marshal.GetActiveObject("Excel.Application")

excel.Visible = True
excel.DisplayAlerts = False   

filename = 'C:\Users\e.green\Desktop\SheetListDataExport.xlsx'
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
### Objects Visibile in View



exportData = horizontalPipes


lastRow = len(exportData)
#lastColumn = len(exportData[1])

totalColumns = len(max(exportData,key=len))

#totalColumns = 1# temporary hack

lastColumn = totalColumns

lastColumnName = ColIdxToXlName(totalColumns)

xlrange = ws.Range["A1", lastColumnName+str(lastRow)]

a = Array.CreateInstance(object, len(exportData),totalColumns)

exportData[1:] = sorted(exportData[1:],key=lambda x: x[1]) 

i = 0 # ignore header row


while i < lastRow:
	j = 0
	while j < totalColumns:
	
		a[i,j] = exportData[i][j]
		j += 1
	
	
	i += 1

xlrange.Value2 = a 

ws.Range(ws.Cells(1,1), ws.Cells(1,lastColumn)).Font.Bold = True
ws.Range(ws.Cells(1,1), ws.Cells(lastRow,4)).Columns.AutoFit()
ws.Range(ws.Cells(1,6), ws.Cells(lastRow,lastColumn)).Columns.AutoFit()
ws.Range(ws.Cells(1,1), ws.Cells(lastRow,lastColumn)).AutoFilter()

ws.Activate