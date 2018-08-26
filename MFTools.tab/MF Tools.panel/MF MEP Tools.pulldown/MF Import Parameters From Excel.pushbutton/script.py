# -*- coding: utf-8 -*-
__title__ = 'Import Parameters from Excel'
__doc__ = """Import Parameters from Excel
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

sys.path.append("\lib")

from MF_HeaderStuff import *

	
from MF_CustomForms import *

from MF_MultiMapParameters import *



# options = ["option 1", "option 2", "option 3"]


# test = SelectFromDoubleList.show(options,
			# title='Choose Parameter to Import',
			# width=800,
			# height=800,
													 # multiselect=False)	



# log = []

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
					



# from MF_ExcelOutput import *

# MF_WriteToExcel("TextData.xlsx", "Tags", tagData)
# MF_WriteToExcel("TextData.xlsx", "TextNotes", textNoteData)

#################################

# Read from Excel

from MF_ExcelInput import *





from rpw.ui.forms import select_file

filepath = select_file('Excel File (*.xlsx)|*.xlsx')

file = filepath

inputData = 	MF_OpenExcelAndRead(file, None, 20 )  # limit to import 20 rows of data




pairs = MF_MultiMapParameters(inputData)
####################################


t = Transaction(doc, 'Import Parameter Values from Excel')
 
t.Start()
 

inputData = 	MF_OpenExcelAndRead(file, None )  # read in all of the data.. 

importData = inputData

headerRow = importData[0]

paramPairs = pairs

idColumnIndex = pairs[0][2] ## index of column containing element ids

for item in importData[1:]:

	#[fieldToImport, selectedParamName]
	
	
	#idColumnIndex = header.index(selected[0])
	
	if item[idColumnIndex]:
		el = doc.GetElement(ElementId(int(item[idColumnIndex])))
		
		existingValue = ' - '
		newValue = ' - '
		
		for p in paramPairs:
			fieldToImportColIndex = headerRow.index(p[0])
			selectedParamName = p[1]
			try:
				existingValue = MF_GetParameterValueByName(el, selectedParamName)
			
			
				newValue = item[fieldToImportColIndex]
			
				print str(el.Id) + " : " + str( selectedParamName) + " : " + str(existingValue) + " --->  " + str(newValue)
			
			except Exception as e:
				
				print "ERROR : " + str(e)
				pass
			
			
			try:
				MF_SetParameterByName(el, selectedParamName, newValue)
			except Exception as e:
				print "ERROR : " + str(e)
				pass
	

print("Done")		


 
t.Commit()
 
#__window__.Close()


