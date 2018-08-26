# -*- coding: utf-8 -*-
__title__ = 'MF Family Fixer'
__doc__ = """Fix Family Issues
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

sys.path.append("\lib")

from MF_HeaderStuff import *


	
from MF_CustomForms import *

from MF_MultiMapParameters import *

import itertools
from itertools import *

from MF_ExcelOutput import *



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

#file = select_file('Excel File (*.xlsx)|*.xlsx', 'Excel File (*.xlsm)|*.xlsm' )

#file = "C:\Users\e.green\Desktop\j6276 - Master View Template Settings2.xlsm"

#inputData = 	MF_OpenExcelAndRead(file, None, 20 )  # limit to import 20 rows of data - to see what we are dealing with


# user selects stuff - need to ask user to select column containing view template ids

#pairs = MF_MultiMapParameters(inputData)

from time import *





#inputData = 	MF_OpenExcelAndRead(file, "Filter Update" )  # now read in all of the data.. 

#importData = inputData

#headerRow = importData[0]

#paramPairs = pairs

#idColumnIndex = pairs[0][2] ## index of column containing element ids



def str2bool(v):
  return str(v).lower() in ("yes", "true", "t", "1")
  
 

def getParameterValue(fp, ft):
	if fp.StorageType == StorageType.Double:
		val = ft.AsDouble(fp)
		return val	
	elif fp.StorageType == StorageType.Integer: 	
		val = ft.AsInteger(fp)
		return val	
	elif fp.StorageType == StorageType.String: 	
		val = ft.AsString(fp)
		return val	
	elif fp.StorageType == StorageType.ElementId: 	
		val = ft.AsElementId(fp)
		return val	
	else: 
		val = ft.AsValueString(fp)
   	
   	return val	 

files = select_file('Revit Family File (*.rfa)|*.rfa', multiple = True )
	
docPaths = files

#doc = DocumentManager.Instance.CurrentDBDocument
elementlist = list()

familyList = []
paramList = []
paramLongList = []

paramList.append([

		("Family Index" ),
		
		("Family Name" ), 
		("Parameter Name" ), 
		("Parameter Group" ), 
		("GUID" ), 
		("Parameter Value" ), 
		("DataType" ), 
		("IsShared" ), 
		("Type Name" ), 
		("Type Index" ), 
		("Total Types" ), 
		("Parameter Type" ), 
		("Type / Instance" ), 
		("Formula"), 
		("Timestamp"), 
		("Category"),
		("Family File Path" )
		])





paramLongList.append([
		("Family Index" ),
		
		("Family Name" ), 
		("Parameter Name" ), 
		("Parameter Group" ), 
		("GUID" ), 
		("Parameter Value" ), 
		("DataType" ), 
		("IsShared" ), 
		("Type Name" ), 
		("Type Index" ), 
		("Total Types" ), 
		("Parameter Type" ), 
		("Type / Instance" ), 
		("Formula"), 
		("Timestamp"), 
		("Category"),
		("Family File Path" )
		])


d = 0




doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
#uiapp = UIApplication(app)


for f in files:

	doc = app.OpenDocumentFile(f)

	paramSubList = []
	
	types = doc.FamilyManager.Types
	ntypes = types.Size
	familyList.append(doc.Title)
	
	typeNames = []
	tid = 0
	
	fam = doc.OwnerFamily
	cat = fam.FamilyCategory
	
	try:
		for t in types:
			tid = tid + 1
			for param in doc.FamilyManager.Parameters:
				
				pdfn = param.Definition
				paramType = "Built In Parameter: " + str(pdfn.BuiltInParameter)
				#if parameter is a shared parameter get the GUID
				guid = " - "
				value = " - "
				
				pTypeOrInstance = "Type"
				
				if param.IsInstance:
					pTypeOrInstance = "Instance"
				
				if t.HasValue(param):
					value = getParameterValue(param, t)
			
				isShared = doc.FamilyManager.get_Parameter(BuiltInParameter.FAMILY_SHARED)
			
				
					
				
				if str(pdfn.BuiltInParameter) == "INVALID":
					paramType = "Family Parameter"
				
				if param.IsShared:
					paramType = "Shared Parameter"
					guid = str(param.GUID)
				
			
				time = strftime("%Y-%m-%d %H%M%S", localtime())
					
				
				
				paramDataRow = [
					d, 
					 
					doc.Title, 
					pdfn.Name, 
					pdfn.ParameterGroup, 
					guid, 
					str(value),  
					param.StorageType, 
					isShared,  
					t.Name, 
					tid, 
					ntypes, 
					paramType, 
					pTypeOrInstance, 
					param.Formula,  
					time, 
					cat.Name,
					f,
					]
				
				
				
				#print paramDataRow
				paramList.append(paramDataRow)
			
	
	except Exception as e:
		print 'Error: ' + str(e)
	
	
	doc.Close()
	
 	d += 1
	#familyList.append(paramSubList)
			


#print paramList
  
# t = Transaction(doc, __title__)

MF_WriteToExcel("Family Data.xlsx", "FamilyData", paramList)

 
# t.Start()

# #updateActions = importData



# t.Commit()
