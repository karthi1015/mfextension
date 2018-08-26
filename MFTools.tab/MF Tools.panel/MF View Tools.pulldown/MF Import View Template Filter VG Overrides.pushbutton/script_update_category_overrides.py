# -*- coding: utf-8 -*-
__title__ = 'Import View Template Settings'
__doc__ = """Import View Template Settings
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

filepath = select_file('Excel File (*.xlsx)|*.xlsx', 'Excel File (*.xlsm)|*.xlsm' )

file = filepath

#inputData = 	MF_OpenExcelAndRead(file, None, 20 )  # limit to import 20 rows of data - to see what we are dealing with


# user selects stuff - need to ask user to select column containing view template ids

#pairs = MF_MultiMapParameters(inputData)


inputData = 	MF_OpenExcelAndRead(file, None )  # now read in all of the data.. 

importData = inputData

#headerRow = importData[0]

#paramPairs = pairs

#idColumnIndex = pairs[0][2] ## index of column containing element ids


	
sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

def str2bool(v):
  return str(v).lower() in ("yes", "true", "t", "1")

  
t = Transaction(doc, __title__)
 
t.Start()

updateCategoryActions = importData

modifiedVTs = log = []

for a in updateCategoryActions[1:]:
# for eact 'action' a in the input spreadsheet list

	#first column has view template Id	
	
	vtIdint = int(a[0])   # choose column from sheet
	
	vtId = ElementId(vtIdint)
	
	vt	= doc.GetElement(vtId)	
	
	modifiedVTs.append(vt)	

	#for cat in cats:
	# for each Category populated above		
		
	catName = a[2]
	catID = a[3]
	
	try:
		category = doc.Settings.Categories.get_Item(catName)
		
		c = System.Enum.ToObject(BuiltInCategory, int(catID) )
		
		categoryList.append(UnwrapElement(category) )
	
		# get values from Update sheet columns... 
		visibility = str2bool(a[4])  
		
		halftone = str2bool(a[6])
		lineweight = int(a[7])
		
		ogs = OverrideGraphicSettings()
				
		ogs.SetHalftone(halftone)
		ogs.SetProjectionLineWeight(lineweight)
		
		time = strftime("%Y-%m-%d %H%M%S", localtime())
		
		#set Visibility
		
		try:  #2017 api
			vt.SetVisibility(category, visibility)
		except: #2018 api
			vt.SetCategoryHidden(category.Id, not(visibility) )
		
		# set Overrides
		try:
			###################################
			vt.SetCategoryOverrides(category.Id, ogs)
			#vt.SetCategoryHidden(cat.Id, True)
			###############################################
			
			log.append( ( ( time),  ( "Success: "), (vt.Name ), (catName), ("updated"), ("visibility:  "), (a[4]), ("Halftone: "), ( str2bool(a[6]) ), ("Lineweight"), ( int(a[7]) )  ) )
		except Exception as e: 
			log.append(" Error setting Category Overrides: " + str(e) )
	except Exception as e: 
		log.append(str(e))
					
				#try:
				#	vt.SetVisibility(cat.Id, visibility )
				#except:
				#	vt.SetCategoryHidden(cat.Id, not(visibility) )  # API changes in Revit 2018			
		
t.Commit()		


print log