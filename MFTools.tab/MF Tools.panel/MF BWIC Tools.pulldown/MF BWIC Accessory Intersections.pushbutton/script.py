# -*- coding: utf-8 -*-
__title__ = 'MF BWIC Accessory Intersections'
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

MF_CheckIntersectionAccessory("blank" , "(walls)", doc)

# dt = Transaction(doc, "Duct Intersections")
# dt.Start()

# pipe_groups = MF_CheckIntersections("pipes", 'dummy', doc)

# duct_groups = MF_CheckIntersections("ducts", 'dummy', doc)

# cabletray_groups = MF_CheckIntersections("cabletrays", 'dummy', doc)

