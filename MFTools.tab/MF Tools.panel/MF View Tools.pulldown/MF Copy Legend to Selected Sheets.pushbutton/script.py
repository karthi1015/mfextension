# -*- coding: utf-8 -*-
__title__ = 'MF Copy Legend to Selected Sheets'
__doc__ = """MF Copy Legend from the active sheet to other sheets selected from a list
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_HeaderStuff import *


## START HERE #############




## ask user to select levels





viewports = revit.activeview.GetAllViewports()





#forms.alert(str(viewports))
options = []
vp_dict = {'{}: {}'.format(doc.GetElement(vp).Name,
								   doc.GetElement(doc.GetElement(vp).ViewId).Name): vp
				   for vp in viewports}
options.extend(vp_dict.keys())
selected_vp = forms.SelectFromList.show(options,
													 multiselect=False)	
													 
source_vpId = vp_dict[selected_vp[0]]
source_vp = doc.GetElement(vp_dict[selected_vp[0]])

source_vp_location = 	source_vp.GetBoxCenter()
	


#vp_location = legend_vp.GetBoxCenter()

dest_sheets = forms.select_sheets()

## choose Viewport to get positions from











		
t = Transaction(doc)
t.Start(__title__)		

for s in dest_sheets:		

	
		
		## create sensible sheet number here... 
		
		vp_location = source_vp_location
		
		#add legend view (lvp) to sheet
		lvp = Viewport.Create(doc, s.Id, source_vp.ViewId,  vp_location )
		
		


t.Commit()		

# Create Drawing Views From Levels and View Template

# Choose views to create Working sheets for

# duplicate selected views and apply 150 template

# create corresponding bg view and apply bg template




# Create Working Sheets from Excel List

# Create Background View - Duplicate as Dependent so can add to sheet

# Add Bg Background View to Working Sheet

# Add Dr Drawing View to Working Sheet (in same location)
