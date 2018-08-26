# -*- coding: utf-8 -*-
__title__ = "MF Batch Change View Template Includes"
__doc__ = """Batch Change View Template Includes
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
		
class ViewTemplateOption(BaseCheckBoxItem):
    def __init__(self, vt_param):
        super(ViewTemplateOption, self).__init__(vt_param)

    @property
    def name(self):
		
        
		return '{}  '.format(self.item.Definition.Name)		


#select multiple
selected = []
return_options = SelectFromCheckBoxes.show(
			sorted([ViewOption(x) for x in viewTemplates],
				   key=lambda x: x.name),
			title="Select View Templates to Modify",
			button_name="Select",
			width=800)
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]


viewTemplates = list(selected)

vtParams = viewTemplates[0].Parameters

selected = []
return_options = SelectFromCheckBoxes.show(
			sorted([ViewTemplateOption(x) for x in vtParams],
				   key=lambda x: x.name),
			title="Select View Template Parameters to set to UnControlled",
			button_name="Select",
			width=800)
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]
		
uncheck = [s.Id.IntegerValue for s in selected]		

#print str(uncheck)		

t = Transaction(doc)
t.Start(__title__)	
	    
for vt in viewTemplates:
	update = False
	
	allParams = [id.IntegerValue for id in vt.GetTemplateParameterIds()]
	
	existingUnControlled = [e.IntegerValue for e in vt.GetNonControlledTemplateParameterIds() ]
	setunchecked = uncheck 
	setunchecked.extend( existingUnControlled )
	
	#print vt.Name + ": " + str(vt.GetNonControlledTemplateParameterIds() )
	
	toSet = []
	for j in allParams:
		#if j not in exclude:
		if j in setunchecked:
			toSet.append(ElementId(j))
			update = True
	if update:
		#need to modify this to also retain existing non controlled items... 
		sysList = List[ElementId](toSet)
		
		print str(sysList)
		vt.SetNonControlledTemplateParameterIds(sysList)
	
t.Commit()
