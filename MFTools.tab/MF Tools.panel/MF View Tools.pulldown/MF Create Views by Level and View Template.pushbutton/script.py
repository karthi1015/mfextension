# -*- coding: utf-8 -*-
__title__ = "Create Views by Level and View Template"
__doc__ = """Create Views by duplicating a selected view, creating one View per level and View Template, names Combining Level Name with View Template name
"""

__helpurl__ = ""

import clr
import os
import os.path as op
import pickle as pl

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

try:
    from pyrevit.versionmgr import PYREVIT_VERSION
except:
    from pyrevit import versionmgr
    PYREVIT_VERSION = versionmgr.get_pyrevit_version()

pyRevitNewer44 = PYREVIT_VERSION.major >= 4 and PYREVIT_VERSION.minor >= 5

if pyRevitNewer44:
    from pyrevit import script, revit, forms
    from pyrevit.forms import *
    output = script.get_output()
    logger = script.get_logger()
    linkify = output.linkify
    from pyrevit.revit import doc, uidoc, selection
    selection = selection.get_selection()

else:
    from scriptutils import logger
    from scriptutils.userinput import SelectFromList, SelectFromCheckBoxes
    from revitutils import doc, uidoc, selection


	
#Create Floor Plan

#viewFamilyTypes = FilteredElementCollector(doc).OfClass(type(ViewFamilyType)).ToElements()

#viewFamilyTypeId = ViewFamily.FloorPlan.Id





# For each Level and View Template, Duplicate Source View with options (not as dependent etc), and set name of view

# Get View Templates

# Get Levels



#Collect View Templates from Project

viewTemplates = []
collector = FilteredElementCollector(doc).OfClass(View)
for i in collector:
	if i.IsTemplate == True:
		viewTemp = i
		viewTemplates.append(i)
		
		
class ViewOption(BaseCheckBoxItem):
    def __init__(self, view_element):
        super(ViewOption, self).__init__(view_element)

    @property
    def name(self):
		
        
		return '{} ({}) '.format(self.item.ViewName, self.item.ViewType)


#select multiple
selected = []
return_options = SelectFromCheckBoxes.show(
			sorted([ViewOption(x) for x in viewTemplates],
				   key=lambda x: x.name),
			title="Select View Templates to Create Views From",
			button_name="Create Views for Selected Templates for Each Level",
			width=800)
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]
		
viewTemplates = list(selected)

		
#select scope box
#options = ["None"]
scopeBoxes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).ToElements()

options = ["No Scope Box"]  ## how to implement this




selected = []
return_options = SelectFromList.show(
			[[x.Name, x] for x in scopeBoxes],
				 
			title="Select Scope Box to Apply to New Views",
			button_name="Select",
			width=800)
if return_options:
		selected = [x for x in return_options ]		

scopeBox = selected[0][1]

print str(scopeBox.Id) + " : " + scopeBox.Name


levels = FilteredElementCollector(doc).OfClass(Level).ToElements()

viewFamilyTypes = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()


	
#create Floor Plan Views for each Level and View Template
name = "Floor Plan"
id = -1
for viewType in viewFamilyTypes:
	typeName = viewType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
	if typeName == name:
		viewTypeId = viewType.Id
		break

t = Transaction(doc)
t.Start(__title__)		
		
for level in levels:
	for vt in viewTemplates:
		try:
			#Create Floor Plan
			v = ViewPlan.Create(doc, viewTypeId, level.Id) 
			
			#Set Name of new Floor Plan - based on level name and View Template name
			v.Name = vt.Name + " - " + level.Name
			
			#Apply View Template  
			v.ViewTemplateId = vt.Id
			
			#Apply Scope Box
			# this seems to take longer than it should... 
			v.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).Set(scopeBox.Id)
		except Exception as e:
			print(str(e))
			pass



t.Commit()


