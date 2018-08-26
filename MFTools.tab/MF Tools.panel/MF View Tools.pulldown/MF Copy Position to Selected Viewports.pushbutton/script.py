# -*- coding: utf-8 -*-
__title__ = 'MF Copy Position to Selected Viewports'
__doc__ = """MF Copy Position to Selected Viewports
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_HeaderStuff import *


## START HERE #############





## select source viewport


sheets = forms.select_sheets(title='Select Source Sheet', button_name='Select',
                  width=800, multiple=True,
                  filterfunc=None, doc=None)


all_viewports = []

sheets = list(sheets)
for sheet in sheets:

	viewports = sheet.GetAllViewports()
	all_viewports.extend(viewports)

options = []

# if revit.ActiveView.ViewType is ViewType.ViewSheet:
	# options.extend(revit.ActiveView)


vp_dict = {'{} : {}: {}'.format(doc.GetElement(doc.GetElement(vp).ViewId).ViewType, doc.GetElement(vp).Name,
								   doc.GetElement(doc.GetElement(vp).ViewId).Name): vp
				   for vp in all_viewports}
options.extend(vp_dict.keys())
selected_vp = forms.SelectFromList.show(options,
			title='Choose Source Viewport to Copy Location',
			width=800,
			height=800,
													 multiselect=False)	
													 
source_vpId = vp_dict[selected_vp[0]]
source_vp = doc.GetElement(vp_dict[selected_vp[0]])

source_vp_location = 	source_vp.GetBoxCenter()
	




## select target sheets
sheets = forms.select_sheets(title='Select Tagret Sheets', button_name='Select',
                  width=800, multiple=True,
                  filterfunc=None, doc=None)


all_viewports = []
for sheet in sheets:

	viewports = sheet.GetAllViewports()
	all_viewports.extend(viewports)





options = []
vp_dict = {'{} : {}: {} {}'.format(doc.GetElement(doc.GetElement(vp).ViewId).ViewType,  doc.GetElement(vp).Name,
								doc.GetElement(doc.GetElement(vp).SheetId).Name,
								   doc.GetElement(doc.GetElement(vp).ViewId).Name): vp
				   for vp in all_viewports}
options.extend(vp_dict.keys())
selected = forms.SelectFromList.show(options,
			title='Choose Viewport Locations to Modify',
			width=800,
			height=800,
													 multiselect=True)	
													 
selected = [ vp_dict[s] for s in selected]

#forms.alert(str(selected))													 

													 
t = Transaction(doc)
t.Start(__title__)		

for s in selected:
	vp = doc.GetElement(s)
	vp.SetBoxCenter(source_vp_location)

## choose Viewport to get positions from



		


t.Commit()		

