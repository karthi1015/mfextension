#############################################
# READ FROM EXCEL 

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

from MF_HeaderStuff import *

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

def MF_ReadFromExcel(infilename, sheet):

	

	excel = Excel.ApplicationClass()   
	from System.Runtime.InteropServices import Marshal

	excel = Marshal.GetActiveObject("Excel.Application")

	excel.Visible = True
	excel.DisplayAlerts = False   

	###################################

	

	desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

	filename = desktop + "\\" + infilename

	# finding a workbook that's already open

	workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename]
	if workbooks:
		workbook = workbooks[0]
	else:
	#Workbooks
	#if workbook exists, try to open it
		try:
			workbook = excel.Workbooks.Open(infilename)
		except Exception as e:
			print "Error opening workbook: " + str(e)


	# choose sheet
	
	if sheet:

		try:
			ws = workbook.Sheets.Item[sheet]
		except Exception as e:
			print "Error opening sheet: " + str(e)
			
	else:

		print "Sheets in Workbook: " + workbook.Sheets.Count
		

	######################################################

	#ws.Activate

	lastRow = 5
	
	
	rowCount = ws.UsedRange.Rows.Count
	
	columnCount = ws.UsedRange.Columns.Count
	
	lastRow = rowCount
	lastColumn = columnCount
	
	print "Rowcount: " + str(rowCount)
	
	importData = []
	
	i=1
	while i <= lastRow:
		
		print "Reading Data: " + str(100* i/rowCount) + " % complete" 
		importRow = []
		j=1
		while j <= lastColumn:
			importRow.append(ws.Cells(i,j).Text)
			j += 1
		importData.append(importRow)		
		i += 1

	return importData
	
	
def MF_OpenExcelAndRead(infilename, sheet=None, rowLimit = None):

	

		excel = Excel.ApplicationClass()   
		from System.Runtime.InteropServices import Marshal

		excel = Marshal.GetActiveObject("Excel.Application")

		excel.Visible = True
		excel.DisplayAlerts = False   

		###################################

		

		#desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

		filename = infilename

		# finding a workbook that's already open

		workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename]
		if workbooks:
			workbook = workbooks[0]
		else:
		#Workbooks
		#if workbook exists, try to open it
			try:
				workbook = excel.Workbooks.Open(infilename)
			except Exception as e:
				print "Error opening workbook: " + str(e)




		# choose sheet
		
		if sheet:

			try:
				ws = workbook.Sheets.Item[sheet]
			except Exception as e:
				print "Error opening sheet: " + str(e)
				
		else:
			
			# choose sheet
			
			sheetOptions = [s.Name for s in workbook.Sheets]
			options = sheetOptions
			selected = forms.SelectFromList.show(options,
			title='Choose Sheet to Import From',
			width=800,
			height=800,
													 multiselect=False)	
			
			#print "Sheets in Workbook: " + str(workbook.Sheets(1).Name)
			sheet = selected[0]
			
			try:
				ws = workbook.Sheets.Item[sheet]
			except Exception as e:
				print "Error opening sheet: " + str(e)

		######################################################

		#ws.Activate

		lastRow = 5
		
		
		rowCount = ws.UsedRange.Rows.Count
		
		columnCount = ws.UsedRange.Columns.Count
		
		lastRow = rowCount
		lastColumn = columnCount
		
		if rowLimit and rowLimit > 0:
			lastRow = rowLimit
			
		
		print "Rowcount: " + str(rowCount)
		
		importData = []
		
		i=1
		while i <= lastRow:
			
			print "Reading Data: " + str(100* i/rowCount) + " % complete" 
			importRow = []
			j=1
			while j <= lastColumn:
				importRow.append(ws.Cells(i,j).Text)
				j += 1
			importData.append(importRow)		
			i += 1

		return importData