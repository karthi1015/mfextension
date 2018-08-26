import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference('DSCoreNodes')
import DSCore
from DSCore import Color

#Import ToDSType(bool) extension method
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

# Import geometry conversion extension methods
clr.ImportExtensions(Revit.GeometryConversion)

# Import DocumentManager and TransactionManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

from System import Enum

from time import gmtime, strftime, localtime

# Import RevitAPI
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
doc = DocumentManager.Instance.CurrentDBDocument

# Import RevitAPI
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

import System
from System import Array
from System.Collections.Generic import *


import sys

targetVT = UnwrapElement(IN[0])

log = []

targetVTs = [targetVT]

categoryIDs = IN[1]

cats = []

catIDs = []

categoryList = []

bics = Enum.GetValues( clr.GetClrType( BuiltInCategory )  )

for bic in bics:

	try:
		if doc.Settings.Categories.get_Item(bic):
			c = doc.Settings.Categories.get_Item(bic)
			cats.append(c)
			catIDs.append(c.Id)
			
			realCat = Category.GetCategory(doc, bic)
			cats.append("realCat" + str(realCat) )
	
	except Exception as e: 
		cats.append(str(e))
		pass

updateCategoryActions = IN[1]
		
TransactionManager.Instance.EnsureInTransaction(doc)

modifiedVTs = []

def str2bool(v):
  return str(v).lower() in ("yes", "true", "t", "1")


for a in updateCategoryActions[1:]:
# for eact 'action' a in the input spreadsheet list

	#first column has view template Id	
	
	vtIdint = int(a[0])
	
	vtId = Autodesk.Revit.DB.ElementId(vtIdint)
	
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
		
TransactionManager.Instance.ForceCloseTransaction()		

OUT = categoryIDs, cats, catIDs, log, categoryList, modifiedVTs
