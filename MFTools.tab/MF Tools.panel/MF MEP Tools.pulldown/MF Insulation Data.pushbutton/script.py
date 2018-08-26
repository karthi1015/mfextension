# -*- coding: utf-8 -*-
__title__ = 'MF Insulation Data'
__doc__ = """Insulation Magic
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



from MF_MEPCurveData import *

pipe_groups = MF_MEPCurveData("pipes", 'dummy', doc)

duct_groups = MF_MEPCurveData("ducts", 'dummy', doc)

cabletray_groups = MF_MEPCurveData("cabletrays", 'dummy', doc)






