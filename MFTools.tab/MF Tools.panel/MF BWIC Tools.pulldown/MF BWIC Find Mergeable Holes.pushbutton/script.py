# -*- coding: utf-8 -*-
__title__ = 'MF BWIC Find Mergeable Holes'
__doc__ = """BWIC Magic
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

sys.path.append("\lib")

from MF_HeaderStuff import *

	
from MF_CustomForms import *

from MF_MultiMapParameters import *

from MF_ExcelOutput import *


from MF_BWICFunctions import *

from MF_CheckIntersections import *

# dt = Transaction(doc, "Duct Intersections")
# dt.Start()

# pipe_groups = MF_CheckIntersections("pipes", 'dummy', doc)

# duct_groups = MF_CheckIntersections("ducts", 'dummy', doc)

# cabletray_groups = MF_CheckIntersections("cabletrays", 'dummy', doc)

targetName = "MXF_Generic Models_BWIC_MEP_Element"

levels = FilteredElementCollector(doc).OfClass(Level).ToElements()

class LevelOption(BaseCheckBoxItem):
	def __init__(self, level_element):
		super(LevelOption, self).__init__(level_element)

	@property
	def name(self):
		
		
		return '{} '.format(self.item.Name)



options = []
seleted = []
return_options = forms.SelectFromCheckBoxes.show(
								[LevelOption(x) for x in levels],
										  
			title="Select Levels",
			button_name="Choose Levels",
			width=800)				
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]													 


levels = list(selected)		

selected_level = selected[0]
  
generic_models =  FilteredElementCollector( doc ).OfCategory(BuiltInCategory.OST_GenericModel).OfClass( FamilyInstance  )

all_bwics = []

for item in list(generic_models):
	if item.Symbol.Family.Name == targetName and item.LevelId == selected_level.Id:
		all_bwics.append(item)
	
		#print targetName


print len(all_bwics)

#print bwic_list
	
	# get bwic_list from project
	#sort bwic_list

bwic_list = []	
for b in all_bwics:
		
		
		
	bwic_list.append([
						b.Id.IntegerValue,
						b.LookupParameter("MF_BWIC Building Element Id").AsInteger(),
						"wall_name",
						b.LookupParameter("MF_BWIC MEP Element Id").AsInteger(),						
						b.LookupParameter("MF_Package").AsString(),
						"str(b.Location)",
						str(b.Location.Point.X),
						str(b.Location.Point.Y),
						str(b.Location.Point.Z),
						b.LookupParameter("MF_Width").AsDouble()*304.8,
						b.LookupParameter("MF_Depth").AsDouble()*304.8,
						str(b.FacingOrientation),
						b.LookupParameter("MF_BWIC - From Space - Id").AsInteger(),
						b.LookupParameter("MF_BWIC - To Space - Id").AsInteger(),
						b.LookupParameter("MF_BWIC - From Space - Name").AsString(),
						b.LookupParameter("MF_BWIC - To Space - Name").AsString()
						])  
	
	sorted_bwic_list = sorted(bwic_list, key = lambda x: int(x[1])) # sorted by wall 
	
	groups = []
	uniquekeys = []
	
from  itertools import groupby
for k, g in groupby(sorted_bwic_list, lambda x: (x[1], x[12], x[13])):
	groups.append(list(g))
	uniquekeys.append(k)
print "--- Groups ---------------------------------------"	
print str(groups)	
		
	
	
	
# # # bwic_list.append([
						# # # str(b.Id),
						# # # str(wall.Id),
						# # # wall.Name,
						# # # str(MEP_Element.Id),						
						# # # MEP_ElementSystemName,
						# # # str(location),
						# # # str(location.X),
						# # # str(location.Y),
						# # # str(location.Z),
						# # # BWIC_width*304.8,
						# # # BWIC_height*304.8,
						# # # str(wall.Orientation)])



group_x_dims = []


items_with_nearby_neighbours = []





merge_groups = find_mergeable_holes(groups)

############################################
for group in merge_groups:

	# this group needs to be sorted.. ?
	
	#group_sorted_by_z = sorted(group, key = lambda x: int(x[0].Location.Point.Z)) # sorted by z  
	
	nearby_groups_z = group_nearby_points_vertical(group)
	
	
	
	
	flattened_z = [x for x in list(nearby_groups_z)]
	
	print "nearby_groups - flattened z --------------------------"
	print flattened_z
	
	for fz in flattened_z:
		
		#nearby_groups_x = group_nearby_points_horizontal(fz)
		
		#flattened_x = [x for x in list(nearby_groups_x)]
		#print "nearby_groups - flattened x -----------------------"
		#print flattened_x
		#for fx in flattened_x:
		
			try:
				#merge_bw_holes(fx)
				merge_bw_holes(fz)
			except Exception as e:
				print str(e)
	
	# for fz in flattened_z:
		
		# nearby_groups_x = group_nearby_points_horizontal(fz)
		
		# flattened_x = [x for x in list(nearby_groups_x)]
		# print "nearby_groups - flattened x -----------------------"
		# print flattened_x
		# for fx in flattened_x:
		
			# try:
				# merge_bw_holes(fx)
			# except Exception as e:
				# print str(e)
	
	# for fz in list(nearby_groups_z):
		# nearby_groups_x = group_nearby_points_horizontal(fz)
		# print "nearby_groups - x"
		# print list(nearby_groups_x)
		# for fx in list(nearby_groups_x):
			# try:
				# merge_bw_holes(fx)
			# except Exception as e:
				# print str(e)

#print GetBWICInfo(doc)
