# -*- coding: utf-8 -*-
__title__ = 'Export View Template Settings'
__doc__ = """Export View Template Settings, Category and Filter Overrides
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


	



import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

sys.path.append("\lib")

from MF_HeaderStuff import *

from MF_GetFilterRules import *
	
from MF_CustomForms import *

from MF_MultiMapParameters import *

from MF_ExcelOutput import *


filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement)

allViews = FilteredElementCollector(doc).OfClass(View).ToElements()

viewTemplates = []

allBuiltInCategories = Enum.GetValues(clr.GetClrType (BuiltInCategory) )

allCategories = [i for i in doc.Settings.Categories]

modelCategories = []

modelCategoryNames = []

time = strftime("%Y-%m-%d %H:%M:%S", localtime())

projectInfo = [["Project File: ", doc.Title], ["Path", doc.PathName], ["Export Date:", time] ]

for c in allCategories:
	
	#c = doc.GetCategory( bic)
	
	
	if c.CategoryType == CategoryType.Model:
		modelCategories.append(c)
		modelCategoryNames.append(c.Name)


for v in allViews:
	if v.IsTemplate:
		viewTemplates.append(v)
		
		fs = v.GetFilters()
		
		

#filters = IN[0]
#view = IN[1]

#filters = view.GetFilters()

bips = Enum.GetValues( clr.GetClrType( BuiltInParameter )  )
#bips = Enum.GetNames( clr.GetClrType( BuiltInParameter ) )
bipIDs = []

for bip in bips:
	bipIDs.append(int(bip))
	bipIDs.append( Enum.Parse( clr.GetClrType( BuiltInParameter ), str(bip) )) 

rules = []
filterRules = []

filterSheetHeadings = [("View Template ID"),("View Template Name"),("Filter ID"), ("Filter Name"),("Categories"),("Visibility"),("Halftone"),("Line Weight"),("Line Colour"), ("Fill Colour"), ("Fill Pattern"),("Transparency"), ("Filter Rules")]
filterRules.append(filterSheetHeadings)
filterHeadingRow = []
filterHeadingRow.append(filterSheetHeadings)


filterDetails = []

ruleValues = []


ruleName = " "
i=0
types = []

catOverrideList = []


categorySheetHeadings = [
	("View Template ID"), 
	("View Template Name"), 
	("Category"), 
	("Category ID"), 
	("Visible"), 
	("Settings"), 
	("Halftone"), 
	("Line Weight"), 
	("Line Colour") 
	]

catOverrideList.append(categorySheetHeadings)

catHeadingRow = []
catHeadingRow.append(categorySheetHeadings)

vtData = []
filterData = []
vtData.append([("View Template Name"),("View Template ID")])
filterData.append([ ("Filter Name"),("Filter ID"), "Filter Rules"])
for f in filters:
	
	
	fRules = MF_GetFilterRules(f)

	filterData.append([ (f.Name),(f.Id), str(fRules) ] )

for vtc in viewTemplates:
	
	
	for c in modelCategories:
		ogs = vtc.GetCategoryOverrides(c.Id)
	
		try:
			visibility = vtc.GetVisibility(c.Id)
		except:
			visibility = not( vtc.GetCategoryHidden(c.Id) )  # API changes in Revit 2018
		
		
		linecol = ogs.ProjectionLineColor
		
		if linecol.IsValid:
		
			linergb = (255, linecol.Red, linecol.Green, linecol.Blue)
		else:
			linergb = None	
		
		
		catOverrideList.append([ 
			(vtc.Id), 
			(vtc.Name), 
			(c.Name), 
			(c.Id), 
			(visibility), 
			(ogs) , 
			(ogs.Halftone), 
			(ogs.ProjectionLineWeight), 
			(linergb ) 
			])

for vt in viewTemplates:
	filters = vt.GetFilters()
	vtFilterRuleList = []
	
	vtData.append([ vt.Name, vt.Id])		
	
	
	for pfe in filters:
		
		f = doc.GetElement(pfe)
		
		
		
		ruleData = ""
		categories = ""
		
		visibility = ""
		
		ruleInfoList = []
		ruleList = []
		
		pfRuleList = []
		try:
			visibility = vt.GetFilterVisibility(f.Id)		
		except:
			pass
		
		filterOverrides = vt.GetFilterOverrides(f.Id)
		
		fillPatternId = filterOverrides.ProjectionFillPatternId
		
		fillcol = filterOverrides.ProjectionFillColor
		
		lineR = "-"
		lineG = "-"
		lineB = "-"
		fillR = "-"
		fillG = "-"
		fillB = "-"
		
		if fillcol.IsValid:
			#fillrgb = DSCore.Color.ByARGB(255, fillcol.Red, fillcol.Green, fillcol.Blue)
			fillR = str(fillcol.Red)
			fillG = str(fillcol.Green)
			fillB = str(fillcol.Blue)
			
			fRGB =  (fillR + "_" + fillG + "_" + fillB )
			
			
		
		else:
			#fillrgb = None
			fRGB = "-"
		
		linecol = filterOverrides.ProjectionLineColor
		
		if linecol.IsValid:
			#linergb = DSCore.Color.ByARGB(255, linecol.Red, linecol.Green, linecol.Blue)
			linergb = (255, linecol.Red, linecol.Green, linecol.Blue)
			lineR = str(linecol.Red)
			lineG = str(linecol.Green)
			lineB = str(linecol.Blue)
			lRGB = (lineR + "_" + lineG + "_" + lineB )
		else:
			linergb = None
			lRGB = "-"
		
		
		
		halftone = filterOverrides.Halftone
		
		transparency = filterOverrides.Transparency
		
		
		lineweight	= filterOverrides.ProjectionLineWeight
		
		
		for c in f.GetCategories():
			categories += Category.GetCategory(doc, c).Name + "  ,  "
		
		
	
			
		sublist = [ 
			(vt.Id), 
			(vt.Name), 
			(f.Id), 
			(f.Name   ), 
			(categories), 
			(visibility), 
			(halftone), 
			( lineweight), 
			(lRGB), 
			(fRGB ) , 
			(fillPatternId),(transparency) 
			]
			
		
		filterRules.append(sublist)




# Start with filterRules


 


excel = Excel.ApplicationClass()   

from System.Runtime.InteropServices import Marshal

excel = Marshal.GetActiveObject("Excel.Application")

excel.Visible = True
excel.DisplayAlerts = False   

# filename = 'C:\Users\e.green\Desktop\VTDataExport.xlsx'

# filename = sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF Tools.panel\MF View Tools.pulldown\MF Export View Template Settings.pushbutton\Master View Template Settings.xlsm"

# # copy template excel file to user's desktop

import shutil

#src = sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF Tools.panel\MF View Tools.pulldown\MF Export View Template Settings.pushbutton\Master View Template Settings - Template.xlsm"

# try  folder local to current script


import sys, os

#print('sys.argv[0] =', sys.argv[0])             
pathname = os.path.dirname(sys.argv[0])        
#print('path =', pathname)
#print('full path =', os.path.abspath(pathname)) 


# copy from template excel file to users desktop
# template excel file has macros for colourising and formatting 


src = os.path.abspath(pathname) + "\Master View Template Settings - Template.xlsm"

filename = doc.Title  + " - View Template Settings.xlsm"

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

shutil.copyfile(src, desktop +"/"+ filename)



#############################################################################
### Filter Overrides

ws = MF_WriteToExcel(filename, "Filter Overrides", filterRules)

ws.Activate

excel.Run("UpdateColours")

#############################################################################
### View Templates Data

ws = MF_WriteToExcel(filename, "View Templates Data", vtData)

#############################################################################
### Filters Data ##################################################

ws = MF_WriteToExcel(filename, "Filters Data", filterData)

#############################################################################
### Category Overrides Data ##################################################

#print catOverrideList
#ws = MF_WriteToExcel(filename, "Category Overrides", catOverrideList)

ws.Activate


excel.Run("Colourise") ## this causes the ctegory list to get truncated.. why??
 



