# -*- coding: utf-8 -*-
__title__ = 'MF Export Space Data'
__doc__ = """Flat Magic
"""

__helpurl__ = ""



import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_HeaderStuff import *

from MF_ExcelOutput import *



areaFilter = ElementCategoryFilter(BuiltInCategory.OST_Areas)

spaceFilter = ElementCategoryFilter(BuiltInCategory.OST_MEPSpaces)



#Collect all the Elements that pass the ElementLevelFilter...
spaces =  FilteredElementCollector(doc).OfClass(SpatialElement).WherePasses(spaceFilter).ToElements()

spaces = list(spaces)

placedSpaces = []

spaceParameters = spaces[0].Parameters
sp_headings = ["Id"]

## choose parameters to export


class ParamOption(BaseCheckBoxItem):
    def __init__(self, view_element):
        super(ParamOption, self).__init__(view_element)

    @property
    def name(self):
		
	
        
		return '{} '.format(self.item.Definition.Name)


#select multiple
selected = []
return_options = SelectFromCheckBoxes.show(
			sorted([ParamOption(x) for x in spaceParameters],
				   key=lambda x: x.name),
			title="Select Parameters to Export",
			button_name="Select",
			width=800)
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]


selectedParamList = list(selected)

selectedParamListNames = [p.Definition.Name for p in selectedParamList]

#print str(selectedParamListNames)

log = []


sp_names = [sp.Definition.Name for sp in spaceParameters]
sp_headings.extend(sp_names)
#print sp_names
output = ["Space Id"]
output.extend( selectedParamListNames ) # headings
for space in spaces[:30]:
	dataRow = [space.Id.IntegerValue]
	try:
		sp = [space.Id.IntegerValue]
		p_values = []
		
		
		#paramList = [p for p in space.Parameters]
		paramList= list(space.Parameters)
		#print str(paramList)
		for param in selectedParamList:  ## need to compare names instead?
		
			if param.Definition.Name in selectedParamListNames:
				paramValue = " - "
				if param.AsValueString():
					paramValue = param.AsValueString()
				elif param.AsString():
					paramValue = param.AsString()	
					
				
				
				
				dataRow.extend(paramValue)
		output.append(dataRow)		
				
				
	
	except Exception as e:
		log.append(str(e))
		print str(e)
		pass
print str(output)
print str(log)

MF_WriteToExcel("SpaceDataTable.xlsx", "Space Data", output)