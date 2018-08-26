# -*- coding: utf-8 -*-
__title__ = 'MF Import View Template Filter VG Overrides'
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

import itertools
from itertools import *



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

file = select_file('Excel File (*.xlsx)|*.xlsx', 'Excel File (*.xlsm)|*.xlsm' )

#file = "C:\Users\e.green\Desktop\j6276 - Master View Template Settings2.xlsm"

#inputData = 	MF_OpenExcelAndRead(file, None, 20 )  # limit to import 20 rows of data - to see what we are dealing with


# user selects stuff - need to ask user to select column containing view template ids

#pairs = MF_MultiMapParameters(inputData)






inputData = 	MF_OpenExcelAndRead(file, "Filter Update" )  # now read in all of the data.. 

importData = inputData

#headerRow = importData[0]

#paramPairs = pairs

#idColumnIndex = pairs[0][2] ## index of column containing element ids



def str2bool(v):
  return str(v).lower() in ("yes", "true", "t", "1")

  
t = Transaction(doc, __title__)
 
t.Start()

updateActions = importData

modifiedVTs = log = []

allRules = updates = []

cLists = []

filterMatches = []

bicats = []

for a in updateActions[1:]:
	vtId = a[0]
	
	view = doc.GetElement(ElementId(int(vtId)))  #Error - multiple targets could match.. ?
	
	fId = 0
	
	if a[2] != "NEW":
	
		fId = int( a[2] )
		
	
	fName = a[3]
	
	
	## get Graphics Settings from Excel update sheet
	
	visibility = str2bool(a[5])
		
	halftone = str2bool(a[6])
	lineweight = int(a[7])
		
	lineRGB = a[8]
	
	#lrgb = lineRGB.split(',')
	lrgb = map(str.strip, lineRGB.split('_'))
	
	
	
	r = int(lrgb[0])
	g = int(lrgb[1])
	b = int(lrgb[2])
	
	lcol = Color(r,g,b)
	
	fillPatternId = a[10]
	
	fillRGB = a[9]
	
	
	
	frgb = map(str.strip, fillRGB.split('_'))
	
	r = int( frgb[0])
	g = int( frgb[1])
	b = int( frgb[2])
	

	
	fcol = Color(r,g,b)
	
	
	
		
	transparency = int(a[11])
		
	ogs = OverrideGraphicSettings()
				
	ogs.SetHalftone(halftone)
	ogs.SetProjectionLineWeight(lineweight)
	
	ogs.SetSurfaceTransparency(transparency)
	
	ogs.SetProjectionLineColor(lcol)
	
	ogs.SetProjectionFillColor(fcol)
	
	ogs.SetProjectionFillPatternId(ElementId(int(fillPatternId)))
	
	
	## do we want to update filter categories from here??
	categories = a[4]  # one long comma separated string from excel cell
	catList = map(str.strip, categories.split(',')) # list of string category names
	cats = []
	cList = []
	for c in catList:
		try:
			#category = doc.Settings.Categories.get_Item(catName)
		
			#c = System.Enum.ToObject(BuiltInCategory, int(catID) )
		
			#categoryList.append(UnwrapElement(category) )
			
			
			
			cat = doc.Settings.Categories.get_Item(c)
			cats.append(cat)  # Autodesk DB. Category
		except: pass	
		
	
	## do we want to update filter rules from here??
	
	ruleData = a[12:] 
	
	
	
	#ruleStringList = list(group(ruleData, ' --- '))
	
	lst = ruleData
	w = ' --- '
	spl = [list(y) for x, y in itertools.groupby(lst, lambda z: z == w) if not x]
	
	ruleStringList = spl
	
	#allRules.append(ruleData)
	allRules.append(ruleStringList)
	
	updates.append([(vtId, fName, fId, categories, cats, ruleData) ])
#############
	doit = False
	
	typedCatList = ' - '
	
	try:
		for cat in cats:
			
			#cList.append(i)  # cannot cast System.String to Element Id
			#cList.append(ElementId(cat.Id))  # ERROR: Expected Built In Parameter - got 'ElementId'
			
			
			cList.append(cat.Id)
			
			typedCatList = List[ElementId](cList)    #unable to cast string to element id
			#typedCatList = cList  ## temporary debug
			#create rule list for filter
	except Exception as e:
		log.append("cat error: "+ str(e) )
	
	#typedCatList = cList  ## temporary debug
	
	
	## check if filter already exists
	existingFilters = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
			
	existingFilterIds = []
	existingFilterNames = []
			
	for ef in existingFilters:
		existingFilterIds.append(ef.Id)  # this is a list of ElementIds
		existingFilterNames.append(ef.Name) 
	
	
	
	
	
	#try:
		
			
	rules = []
	
	for r in ruleStringList:
		paramName = r[1]
		comparator = r[2]
		
		pValue = r[4]
			
		bip = GetBuiltInParam(paramName)
		#string contains
		if comparator == "contains":
			rules.append(ParameterFilterRuleFactory.CreateContainsRule(ElementId(bip), pValue, False) )
		# not contains
		if comparator == "does not contain":
			rules.append(ParameterFilterRuleFactory.CreateNotContainsRule(ElementId(bip), pValue, False) )
		
			
		
	try:	
			
			## ClearRules for existing filter
			
			## Update with new rules ? 
		
			#fId = int(filterIDs[0])
		
		
		match = "not matched"
		# update existing filters
		
		
		if ElementId(fId) in existingFilterIds: # pfe does not get defined?
			
			match = "matched"			
			pfe = doc.GetElement(ElementId(fId))
			
			pFilter = pfe
			# get previous name
			previousName = pfe.Name
			# update name
			pfe.Name = fName
			
			if doit:
				pfe.ClearRules()
				pfe.SetRules(rules)
				## ParameterFilterElement.SetRules( rules)
			
			
			
			ruleDataString ='(' + '),( '.join(ruleData) + ')'
			
			logRow = [ (time), ("Filter " ),( previousName), ( " updated to "),(  pfe.Name), ("Rules:") ]
			logRow.extend(ruleData) 
			
			####
			

			####
			
			
			log.append(logRow) 
		
		filterMatches.append(match)
		#add new filters?
		logRow = [ (time) ]
		
		
		#if fId == 0:
		try:
			## try creating filter
			
			## need to check if fName matches any existing filter name... 
			if fName not in existingFilterNames:
				pfe = ParameterFilterElement.Create(doc, fName, typedCatList, rules)
			
				pFilter = pfe			
			
				createResult = [("Filter Created :" ),( pfe.Name)]
			else:
				createResult = [("Filter name :" ),( fName), ("already in use")]
			
				i = existingFilterNames.index(fName)
				id = existingFilterIds[i]
				
				pfe = doc.GetElement(id)
				
				pFilter = pfe	
			
			logRow.extend(createResult)
			#view.AddFilter(pfe.Id)
		
		
		except Exception as e: 
			createResult = [( "Error Creating Filter : "),(  str(e) )]
			logRow.extend(createResult)
		pass
		log.append(logRow) 
		
		# Try Adding the filter pfe to the view - if it is already added it will throw an error
		try:
			viewFilters = view.GetFilters()
			
			if pFilter.Id not in viewFilters:
			
				view.AddFilter(pFilter.Id)
						
				addResult = [(fName),("added to " ),( view.Name)]
			else: 
				addResult = [(fName),("already applied to " ),( view.Name)]
			
			logRow.extend(addResult) 
		except Exception as e: 
			addResult = [ ("Error Adding Filter"),(fName),("to view"), (view.Name), (str(e)) ]
			logRow.extend(addResult) 
		pass
		
		#set the overrides for the filter now that it is in in the View
		try:
			view.SetFilterOverrides(pFilter.Id, ogs)
			
			view.SetFilterVisibility(pFilter.Id,visibility)
			setResult = [(fName),("overrides updated in " ),( view.Name)]
			logRow.extend(setResult)
		except Exception as e: 
			setResult = [ ("Error setting overrides for Filter"),(fName),("in view"), (view.Name), (str(e)) ]
			logRow.extend(setResult) 
		pass 	
		
		
		
		log.append(logRow) 
		
		
		
	except Exception as e:
		log.append(str(e))	
	
	
	cLists.append(cList)
t.Commit()

# Place your code below this line

# Assign your output to the OUT variable.
#print str([updates, allRules, log, cLists, existingFilterIds, filterMatches])

for entry in log:
	print str(entry)