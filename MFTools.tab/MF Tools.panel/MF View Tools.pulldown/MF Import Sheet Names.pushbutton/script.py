# -*- coding: utf-8 -*-
__title__ = 'Import Sheet Names and Numbers from Excel'
__doc__ = """Import Sheet Names and Numbers from Excel List
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_HeaderStuff import *

	


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
					



# from MF_ExcelOutput import *

# MF_WriteToExcel("TextData.xlsx", "Tags", tagData)
# MF_WriteToExcel("TextData.xlsx", "TextNotes", textNoteData)

#################################

# Read from Excel

from MF_ExcelInput import MF_ReadFromExcel

inputData = 	MF_ReadFromExcel("SheetListDataExport.xlsx", "Update")

for row in inputData:
	print str(row)





t = Transaction(doc, 'Import Sheet Names from Excel')
 
t.Start()
 


importData = inputData

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


