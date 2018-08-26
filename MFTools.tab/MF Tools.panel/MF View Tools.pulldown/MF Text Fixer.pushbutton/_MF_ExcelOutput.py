#############################################
# WRITE TO EXCEL 

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

def MF_WriteToExcel(outfilename, sheet, data):

	

	excel = Excel.ApplicationClass()   
	from System.Runtime.InteropServices import Marshal

	excel = Marshal.GetActiveObject("Excel.Application")

	excel.Visible = True
	excel.DisplayAlerts = False   

	###################################



	desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

	filename = desktop + "\\" + outfilename

	# finding a workbook that's already open

	workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename]
	if workbooks:
		workbook = workbooks[0]
	else:
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




	try:
		ws = workbook.Sheets.Item[sheet]
	except:
		ws = workbook.Worksheets.Add()

		ws.Name = sheet

	######################################################

	

	
	exportData = data
	


	lastRow = len(exportData)
	

	totalColumns = len(max(exportData,key=len))

	

	lastColumn = totalColumns

	lastColumnName = ColIdxToXlName(totalColumns)

	xlrange = ws.Range["A1", lastColumnName+str(lastRow)]

	a = Array.CreateInstance(object, len(exportData),totalColumns)

	

	i = 0 


	while i < lastRow:
		j = 0
		while j < totalColumns:
		
			a[i,j] = exportData[i][j]
			j += 1
		
		
		i += 1

	xlrange.Value2 = a 

	ws.Range(ws.Cells(1,1), ws.Cells(1,lastColumn)).Font.Bold = True

	ws.Range(ws.Cells(1,2), ws.Cells(lastRow,lastColumn)).Columns.AutoFit()
	ws.Range(ws.Cells(1,1), ws.Cells(lastRow,lastColumn)).AutoFilter()