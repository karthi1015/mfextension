# -*- coding: utf-8 -*-
__title__ = 'Clean Family Paths'
__doc__ = """Clean Family Paths
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

sys.path.append("\lib")

from MF_HeaderStuff import *

	
from MF_CustomForms import *

from MF_MultiMapParameters import *

from MF_ExcelOutput import *


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
	

families = FilteredElementCollector(doc).OfClass(Family)

centralModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())

# ﻿

 # \\Maxfordham.com\Jobs\J6082\Cad\Revit\MF Model\J6082 MEP Model.rvt
 
basePath = centralModelPath.split("\MF Model",1)[0] 

savePath = basePath + "\Project Families\@_From Template"

print savePath
#sys.exit()
output = []
output.append(["Family Name", "FilePath"])


for family in families:
	try:
		
		
		famDoc = doc.EditFamily(family)
		
		fName = family.Name
		if famDoc.PathName is not '':
			print savePath + "\\" + fName + ".rfa"
		
			famDoc.SaveAs( savePath + "\\" + fName + ".rfa" )
		
		famDoc.Close()
		
		#output.append([ family.Name , famDoc.PathName ])
		
		## unfortunately = PathName is read only
		# t = Transaction(famDoc, 'Clean Family Paths')
		
		# t.Start()
		# famDoc.PathName = ""
		# t.Commit()
		
		# BUT - if we get the project job folder name from the central file path, and then save as to the Project Families folder
		# and then reload into the project.. 
		
		
	
	
	except Exception as e:
		print str(e) + " " + family.Name
	
				

# from MF_ExcelOutput import *

MF_WriteToExcel("Project Family Paths.xlsx", "FamilyPaths", output)
# MF_WriteToExcel("TextData.xlsx", "TextNotes", textNoteData)


sys.exit()	
#################################

# Read from Excel







from rpw.ui.forms import select_file

filepath = select_file('Excel File (*.xlsx)|*.xlsx')

file = filepath

inputData = 	MF_OpenExcelAndRead(file, None, 20 )  # limit to import 20 rows of data




pairs = MF_MultiMapParameters(inputData)
####################################
#sys.exit()

# for row in inputData[0]:
	# print str(row)


# ### choose parameter (s) to import from sheet by column heading	

# headerRow = inputData[0]
# options = []
# #
# options.extend(headerRow)
# selected = forms.SelectFromList.show(options,
			# title='Choose Parameter to Import',
			# width=800,
			# height=800,
													 # multiselect=False)	

	
	

# #find index of selected item	
# #print "Selected Field: " + str(selected[0])

# fieldToImport =  str(selected[0])

# fieldToImportColIndex = options.index(selected[0])

# print "Selected Field: " + str(selected[0]) + " --- Column Index: " + str( options.index(selected[0]))

# ### choose column containing element ids

# headerRow = inputData[0]
# options = []
# #
# options.extend(headerRow)
# selected = forms.SelectFromList.show(options,
			# title='Choose Column Containing Element IDs',
			# width=800,
			# height=800,
													 # multiselect=False)	

	
	

# #find index of selected item	
# #print "Selected Field: " + str(selected[0])



# idColumnIndex = options.index(selected[0])



# # check what elements we have in the input sheet 

# sampleDataRow = inputData[1] # look at first row of data

# ## choose column containing element Ids.. 

# elementIdstring = sampleDataRow[idColumnIndex]  ## temporary

# sampleElement = doc.GetElement(ElementId(int(elementIdstring)))

# sampleElementParams = sampleElement.Parameters

# options = [p.Definition.Name for p in sampleElementParams]
# #
# #options.extend(headerRow)
# selected = forms.SelectFromList.show(options,
			# title='Choose Element Parameter to Update',
			# width=800,
			# height=800,
													 # multiselect=False)	

	

	
# selectedParamName = selected[0]


# ## build a list of pairs of (importField, selectedParamname)

# # test = forms.SelectFromDoubleList.show(options,
			# # title='Choose Parameter to Import',
			# # width=800,
			# # height=800,
													 # # multiselect=False)	
	
# ####################
# #sys.exit()
# #######################

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


