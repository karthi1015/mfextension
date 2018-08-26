# -*- coding: utf-8 -*-
__title__ = 'Import View Filters from Excel'
__doc__ = """Import View Filters from Excel
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

# clr.AddReference('DSCoreNodes')
# import DSCore
# from DSCore import Color

import System

from System import Array

from System import Enum

from time import gmtime, strftime, localtime

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

try:
    from pyrevit.versionmgr import PYREVIT_VERSION
except:
    from pyrevit import versionmgr
    PYREVIT_VERSION = versionmgr.get_pyrevit_version()

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


	
sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_GetFilterRules import *



filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement)

allViews = FilteredElementCollector(doc).OfClass(View).ToElements()

viewTemplates = []

allBuiltInCategories = Enum.GetValues(clr.GetClrType (BuiltInCategory) )

allCategories = [i for i in doc.Settings.Categories]

modelCategories = []

modelCategoryNames = []

time = strftime("%Y-%m-%d %H:%M:%S", localtime())

projectInfo = [["Project File: ", doc.Title], ["Path", doc.PathName], ["Export Date:", time] ]

def str2bool(v):
  return str(v).lower() in ("yes", "true", "t", "1")
  
  ############################################

def GetBuiltInParam(paramName):
	builtInParams = System.Enum.GetValues(BuiltInParameter)
	
	test = []
	
	for i in builtInParams:
		if i.ToString() == paramName:
			test.append(i)
			break
		else:
			continue
	return test[0]
	
##############################
def group(seq, sep):
    g = []
    for el in seq:
        if el == sep:
            yield g
            g = []
	g.append(el)
    yield g
	

time = strftime("%Y-%m-%d %H:%M:%S", localtime())

bics = Enum.GetValues( clr.GetClrType( BuiltInCategory )  )
bicats = []
for bic in bics:

	try:
		if doc.Settings.Categories.get_Item(bic):
			c = doc.Settings.Categories.get_Item(bic)
			bicats.append(c)
			#catIDs.append(c.Id)
			
			realCat = Category.GetCategory(doc, bic)
			bicats.append("realCat" + str(realCat) )
	
	except Exception as e: 
		bicats.append(str(e))
		pass


# The inputs to this node will be stored as a list in the IN variables.
dataEnteringNode = IN

updateActions = IN[0]

updates = []

log = []

allRules = []

typedCatLists = []
cLists = []

filterMatches = []

colours = []

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
	
	colours.append(lrgb)
	
	r = Convert.ToByte( lrgb[0])
	g = Convert.ToByte( lrgb[1])
	b = Convert.ToByte( lrgb[2])
	

	
	lcol = Color(r,g,b)
	
	fillPatternId = a[10]
	
	fillRGB = a[9]
	
	
	
	frgb = map(str.strip, fillRGB.split('_'))
	
	r = Convert.ToByte( frgb[0])
	g = Convert.ToByte( frgb[1])
	b = Convert.ToByte( frgb[2])
	

	
	fcol = Color(r,g,b)
	
	
	
		
	transparency = int(a[11])
		
	ogs = OverrideGraphicSettings()
				
	ogs.SetHalftone(halftone)
	ogs.SetProjectionLineWeight(lineweight)
	
	ogs.SetSurfaceTransparency(transparency)
	
	ogs.SetProjectionLineColor(lcol)
	
	ogs.SetProjectionFillColor(fcol)
	
	ogs.SetProjectionFillPatternId(ElementId(int(fillPatternId)))
	
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
		
	ruleData = a[12:] 
	
	
	
	ruleStringList = list(group(ruleData, ' --- '))
	
	lst = ruleData
	w = ' --- '
	spl = [list(y) for x, y in itertools.groupby(lst, lambda z: z == w) if not x]
	
	ruleStringList = spl
	
	#allRules.append(ruleData)
	allRules.append(ruleStringList)
	
	updates.append([(vtId, fName, fId, categories, cats, ruleData) ])
#############
	doit = 1
	
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
	
	
	#if doit:
	
	
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
		
		t = Transaction(doc, __title__)
		
		t.Start()
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
		
		
		t.Commit()
	except Exception as e:
		log.append(str(e))	
		
	cLists.append(cList)


# Place your code below this line

# Assign your output to the OUT variable.
print str( [updates, allRules, log, bicats, cLists, existingFilterIds, filterMatches, colours] )
