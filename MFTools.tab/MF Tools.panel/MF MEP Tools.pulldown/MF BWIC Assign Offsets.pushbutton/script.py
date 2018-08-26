# -*- coding: utf-8 -*-
__title__ = 'MF BWIC Assign Offsets'
__doc__ = """BWIC Magic
"""

__helpurl__ = ""

__context__ = 'Selection'

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

sys.path.append("\lib")

from MF_HeaderStuff import *

	
from MF_CustomForms import *

from MF_MultiMapParameters import *

from MF_ExcelOutput import *

clr.AddReference("System")
from System.Collections.Generic import List




	
	
from rpw import db

@db.Transaction.ensure('Assign Offsets to BWIC wall penetrations')
def assign_offsets_to_bw_holes():  #
	

	
	ids = uidoc.Selection.GetElementIds()
	
	
	

	
	for id in ids:
		#el = doc.GetElement(ElementId(int(id)))
		el = doc.GetElement(id)
		
		#get height above level
		bwic_offset = el.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble() # ft
		
		bwic_z = el.LookupParameter("MF_Depth").AsDouble() # ft
		
		bwic_height = bwic_offset - (0.5 * bwic_z)
		
		el.LookupParameter("MF_BWIC Elevation Above Level").Set(bwic_height)
		
		
		
		
	
		
		
assign_offsets_to_bw_holes()


