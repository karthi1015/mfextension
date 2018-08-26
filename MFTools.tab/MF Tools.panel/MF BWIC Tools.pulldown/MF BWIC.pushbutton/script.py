# -*- coding: utf-8 -*-
__title__ = 'MF BWIC'
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

pipe_groups = MF_CheckIntersections("pipes", 'dummy', doc)

duct_groups = MF_CheckIntersections("ducts", 'dummy', doc)

cabletray_groups = MF_CheckIntersections("cabletrays", 'dummy', doc)



# #print bwic_list
	
	## get bwic_list from project
	# #sort bwic_list
	
	# sorted_bwic_list = sorted(bwic_list[1:], key = lambda x: int(x[1]))
	
	# groups = []
	# uniquekeys = []
	
	# from  itertools import groupby
	# for k, g in groupby(sorted_bwic_list, lambda x: x[1]):
		# groups.append(list(g))
		# uniquekeys.append(k)
	# print "--- Groups ---------------------------------------"	
	# print str(groups)	
		
	
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
	
	
	
	# group_x_dims = []
	
	
	# items_with_nearby_neighbours = []





#find_mergeable_holes(duct_groups)


#print GetBWICInfo(doc)
