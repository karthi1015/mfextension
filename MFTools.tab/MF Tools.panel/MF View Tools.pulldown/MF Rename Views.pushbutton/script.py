# -*- coding: utf-8 -*-
__title__ = 'Rename Views by Level and View Template'
__doc__ = """Rename Views by Combining Level Name with View Template name
"""

__helpurl__ = "https://apex-project.github.io/pyApex/help#copy-vg-filters"

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


floorPlanViews = []

newNames = []
collector = FilteredElementCollector(doc).OfClass(View)
views = []
for c in collector:
	if not c.IsTemplate:
		views.append(c)
		

class ViewOption(BaseCheckBoxItem):
    def __init__(self, view_element):
        super(ViewOption, self).__init__(view_element)

    @property
    def name(self):
		vtName = "<No Template>"
		
		genLevel = "<No Level>"
		if doc.GetElement(self.item.ViewTemplateId):
			vtName = doc.GetElement(self.item.ViewTemplateId).Name
		if self.item.GenLevel:
			genLevel = self.item.GenLevel.Name	
        
		return '{} ({})  -----  [{}] ---  [{}]'.format(self.item.ViewName, self.item.ViewType, str(vtName), genLevel)


#select multiple
selected = []
return_options = SelectFromCheckBoxes.show(
			sorted([ViewOption(x) for x in views],
				   key=lambda x: x.name),
			title="Select Views to Rename",
			button_name="Rename Selected Views",
			width=800)
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]


viewList = list(selected)


log = []

t = Transaction(doc)
t.Start(__title__)

for i in viewList:
	if i.GenLevel:
		floorPlanView = i
		floorPlanViews.append(i)
		vt = doc.GetElement(i.ViewTemplateId)
		
		try:
			newName = vt.Name + " - " + i.GenLevel.Name
			newNames.append(newName)

		
			i.Name = newName
			log.append("Success: " + newName )
		except Exception as e:
			log.append("Error:  : " + str(e) )
			pass
	if i.ViewType == ViewType.ThreeD:
		
		vt = doc.GetElement(i.ViewTemplateId)
		
		try:
			newName = vt.Name + " - " + "3D"
			newNames.append(newName)

		
			i.Name = newName
			log.append("Success: " + newName )
		except Exception as e:
			log.append("Error:  : " + str(e) )
			pass		
print log		
t.Commit()


