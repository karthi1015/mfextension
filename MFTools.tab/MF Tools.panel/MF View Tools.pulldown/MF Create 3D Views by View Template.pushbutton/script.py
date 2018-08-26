# -*- coding: utf-8 -*-
__title__ = "Create 3D Views by View Template"
__doc__ = """Create 3D Views for each selected View Template, names Combining Level Name with View Template name
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

#viewFamilyTypeId = ViewFamily.ThreeD.Id





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
			button_name="Create 3D Views for Selected Templates",
			width=800)
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]
		
viewTemplates = list(selected)

		




levels = FilteredElementCollector(doc).OfClass(Level).ToElements()

viewFamilyTypes = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()


	
#create Floor Plan Views for each Level and View Template
name = "3D View"
id = -1
for viewType in viewFamilyTypes:
	typeName = viewType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
	if typeName == name:
		viewTypeId = viewType.Id
		break

		
#viewFamilyTypeId = ViewFamily.ThreeD.Id
		
t = Transaction(doc)
t.Start(__title__)		
		

for vt in viewTemplates:
	try:
		#Create Floor Plan
		v = View3D.CreateIsometric(doc, viewTypeId) 
		
		#Set Name of new Floor Plan - based on level name and View Template name
		v.Name = vt.Name + " - 3D"
		
		#Apply View Template  
		v.ViewTemplateId = vt.Id
	except Exception as e:
		print(str(e))
		pass



t.Commit()


