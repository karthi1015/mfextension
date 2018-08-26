
import clr
import os
import os.path as op
import pickle as pl

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

import System

from System import Array
from System import Enum

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

sys.path.append("\lib")

from MF_HeaderStuff import *

	
from MF_CustomForms import *

from MF_MultiMapParameters import *



from MF_ExcelOutput import *


clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel



# #clr.AddReference('ProtoGeometry')
# #from Autodesk.DesignScript.Geometry import *

# clr.AddReference('DSCoreNodes')
# from DSCore import *

# clr.AddReference("RevitNodes")
# from Revit import *

try:
    from pyrevit.versionmgr import PYREVIT_VERSION
except:
    from pyrevit import versionmgr
    PYREVIT_VERSION = versionmgr.get_pyrevit_version()

from pyrevit import script
#from pyrevit import scriptutils 
from pyrevit import framework
from pyrevit import revit, DB, UI
from pyrevit import forms

from System import Enum 

pyRevitNewer44 = PYREVIT_VERSION.major >= 4 and PYREVIT_VERSION.minor >= 5

if pyRevitNewer44:
    from pyrevit import script, revit
    from pyrevit.forms import SelectFromList, SelectFromCheckBoxes
    output = script.get_output()
    logger = script.get_logger()
    linkify = output.linkify
    from pyrevit.revit import doc, uidoc, selection
    selection = selection.get_selection()

else:
    from scriptutils import logger
    from scriptutils.userinput import SelectFromList, SelectFromCheckBoxes
    from revitutils import doc, uidoc, selection

output = script.get_output()
output.set_width(1100)


# RBS_CONDUITRUN_OUTER_DIAM_PARAM	"Outside Diameter"
# RBS_CONDUITRUN_INNER_DIAM_PARAM	"Inside Diameter"
# RBS_CONDUITRUN_DIAMETER_PARAM	"Diameter(Trade Size)"
# RBS_CABLETRAYRUN_WIDTH_PARAM	"Width"
# RBS_CABLETRAYRUN_HEIGHT_PARAM	"Height"
# RBS_CABLETRAYCONDUITRUN_LENGTH_PARAM	"Length"
# RBS_LOAD_SUB_CLASSIFICATION_MOTOR	"Load Sub-Classification Motor"
# RBS_CABLETRAY_SHAPETYPE	"Shape"
# RBS_CABLETRAYCONDUIT_BENDORFITTING	"Bend or Fitting"
# RBS_CTC_SERVICE_TYPE	"Service Type"
# RBS_CONDUIT_OUTER_DIAM_PARAM	"Outside Diameter"
# RBS_CONDUIT_INNER_DIAM_PARAM	"Inside Diameter"
# RBS_CTC_BOTTOM_ELEVATION	"Bottom Elevation"
# RBS_CTC_TOP_ELEVATION	"Top Elevation"
# RBS_CONDUIT_DIAMETER_PARAM	"Diameter(Trade Size)"
# RBS_CABLETRAY_WIDTH_PARAM	"Width"
# RBS_CABLETRAY_HEIGHT_PARAM	"Height"

# RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM	"System Abbreviation"

# RBS_DUCT_SYSTEM_TYPE_PARAM	"System Type"

from MF_BWICFunctions import *

# global data 
# wall_keys = get_rooms_from_walls()[0]
	
# spaces_grouped_by_wall = get_rooms_from_walls()[1]	

# print "wall_keys"
# print wall_keys

# print "spaces_grouped_by_wall"
# print spaces_grouped_by_wall


from rpw import db
#@db.Transaction.ensure('Do Something')

import math



def roundup(x, interval):
	return int(math.ceil(x/interval)) * interval


def GetAngleFromMEPCurve(curve):

	for c in curve.ConnectorManager.Connectors:
     
		return math.asin(c.CoordinateSystem.BasisY.X);
     
	return 0

def pointOnLine(x1,y1,x2,y2,r):
	d = sqrt((x2-x1)^2 + (y2 - y1)^2) #distance
	r = n / d #segment ratio

	x3 = r * x2 + (1 - r) * x1 #find point that divides the segment
	y3 = r * y2 + (1 - r) * y1 #into the ratio (1-r):r
	
	return (x3,y3)

	
#########################################	
	
def point_inside_polygon(x, y, poly, include_edges=True):
    '''
    Test if point (x,y) is inside polygon poly.

    poly is N-vertices polygon defined as 
    [(x1,y1),...,(xN,yN)] or [(x1,y1),...,(xN,yN),(x1,y1)]
    (function works fine in both cases)

    Geometrical idea: point is inside polygon if horisontal beam
    to the right from point crosses polygon even number of times. 
    Works fine for non-convex polygons.
    '''
    n = len(poly)
    inside = False

    p1x, p1y = poly[0]
    for i in range(1, n + 1):
        p2x, p2y = poly[i % n]
        if p1y == p2y:
            if y == p1y:
                if min(p1x, p2x) <= x <= max(p1x, p2x):
                    # point is on horisontal edge
                    inside = include_edges
                    break
                elif x < min(p1x, p2x):  # point is to the left from current edge
                    inside = not inside
        else:  # p1y!= p2y
            if min(p1y, p2y) <= y <= max(p1y, p2y):
                xinters = (y - p1y) * (p2x - p1x) / float(p2y - p1y) + p1x

                if x == xinters:  # point is right on the edge
                    inside = include_edges
                    break

                if x < xinters:  # point is to the left from current edge
                    inside = not inside

        p1x, p1y = p2x, p2y

    return inside	

#####################################

def fuzzyMatch(a,b, precision):
	return abs(a - b) <= precision
	
def ccw(A,B,C):
    return (C.Y-A.Y) * (B.X-A.X) > (B.Y-A.Y) * (C.X-A.X)

# Return true if line segments AB and CD intersect
def intersect(A,B,C,D):
    return ccw(A,C,D) != ccw(B,C,D) and ccw(A,B,C) != ccw(A,B,D)	
	
def line_intersection(line1, line2):
    xdiff = (line1[0][0] - line1[1][0], line2[0][0] - line2[1][0])
    ydiff = (line1[0][1] - line1[1][1], line2[0][1] - line2[1][1]) #Typo was here

    def det(a, b):
        return a[0] * b[1] - a[1] * b[0]

    div = det(xdiff, ydiff)
    if div == 0:
       raise Exception('lines do not intersect')

    d = (det(*line1), det(*line2))
    x = det(d, xdiff) / div
    y = det(d, ydiff) / div
    return x, y	

@db.Transaction.ensure('PlaceBWICs')
def MF_PlaceBWIC(location, BWICFamilySymbol,level):
										
	BWICInstance = doc.Create.NewFamilyInstance( 
										location, 
										BWICFamilySymbol,
										level,
										Structure.StructuralType.NonStructural
										)
	b = BWICInstance
	return b

#@db.Transaction.ensure('Check Intersections')
def MF_MEPCurveData(MEP_ElementsToCheck, building_Element, doc):
	## convert this into a function for every MEP curve type? (pipes, duct, cabletray)
	
	# MEP_Elements include pipe, duct, cabletray
	
	#dt = Transaction(doc, "Duct Intersections")
	#dt.Start()
	
	
	# Levels #######################################################################
	## ask user to select levels

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
	
	
	
	#building element filter
	filter = ElementCategoryFilter(BuiltInCategory.OST_Levels)
	
	#levels = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()
	
	#levelIndex = 0
	
	#levelFilter = ElementLevelFilter(levels[levelIndex].Id)
	
	levelFilter = ElementLevelFilter(selected_level.Id)
	
	levelIndex = list(levels).index(selected_level)
	
	filter = ElementCategoryFilter(BuiltInCategory.OST_PipeCurves)
	
	# set up level filter for Reference Level

	pipes = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()

	filter = ElementCategoryFilter(BuiltInCategory.OST_DuctCurves)

	ducts = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()

	filter = ElementCategoryFilter(BuiltInCategory.OST_CableTray)

	cable_trays = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()
	
	pipes = [p for p in pipes  ]
	
	ducts = [d for d in ducts ]
	
	cable_trays = [c for c in cable_trays]
	
	print levels[levelIndex].Name 
	
	print levels[levelIndex].ProjectElevation
	
	
	
	
	# looking in linked document for wall elements... 
	links = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements() 
	linkDoc = links[0].GetLinkDocument() # assumes that the architects document is the first linked document! risky!
	
	filter = ElementCategoryFilter(BuiltInCategory.OST_Levels)
	
	linkedLevels = FilteredElementCollector(linkDoc).WhereElementIsNotElementType().WherePasses(filter).ToElements()
	
	#levelFilter = ElementLevelFilter(linkedLevels[levelIndex].Id)
	
	#print linkedLevels[levelIndex].Name
	#print linkedLevels[levelIndex].ProjectElevation
	
	## need to match levels by name!
	for l in linkedLevels:
		if levels[levelIndex].Name == l.Name:
		
			levelFilter = ElementLevelFilter(l.Id)
		
			print l.Name
		
	#sys.exit()

	
	filter = ElementCategoryFilter(BuiltInCategory.OST_Walls)
	
	curveDriven = ElementIsCurveDrivenFilter()
	
	## levels in liked file have different ids.... compare indexs instead
	walls = FilteredElementCollector(linkDoc).WhereElementIsNotElementType().WherePasses(filter).WherePasses(levelFilter).WherePasses(ElementIsCurveDrivenFilter()).ToElements()
	
	
	
	
	
	# if str(MEP_Element) is "ducts":
		# MEP_Elements = ducts
	# if str(MEP_Element) is "pipes":
		# MEP_Elements = pipes
	
	if MEP_ElementsToCheck == "ducts":
	
		MEP_Elements = ducts
		
	if MEP_ElementsToCheck == "pipes":
	
		MEP_Elements = pipes	
	
	if MEP_ElementsToCheck == "cabletrays":
	
		MEP_Elements = cable_trays	
	
	curves = []
	linePoints = []
	
	intersections = []
	slopingMEP_Elements = []
	all_MEP_Elements = []
	
	
	verticalMEP_ElementRuns = []

	headings = [					"level.Name", 
									"levelZ",
									"MEP_Element.Id", 
									"MEP_ElementSystem", 
									"MEP_ElementSystemName", 
									"length",
									"Width",
									"Height",
									"Diameter (OD)",
									"Insulation Thickness",
									
									"startPoint.X", 
									"startPoint.Y", 
									"startPoint.Z", 
									"endPoint.X", 
									"endPoint.Y", 
									"endPoint.Z",
									"deltaX",
									"deltaY",
									"deltaZ",
									"orientation"
									]
								




	

	verticalMEP_Elements = []
	horizontalMEP_Elements = []

	all_MEP_Elements.append(headings)


	
	print "Walls found: " + str(len(walls))
		
	print "Ducts found:" + str(len(ducts))
	
	print "Pipes found:" + str(len(pipes))
	
	print "Cable Trays found:" + str(len(cable_trays	))
	

	
	
	
	
	for MEP_Element in MEP_Elements:
	
		#set defaults
		MEP_ElementDiameter = 0
		MEP_ElementWidth  = 0
		MEP_ElementHeight  = 0
		
		line = MEP_Element.Location.Curve
		
		curves.append(line)
		
		startPoint = line.GetEndPoint(0)
		endPoint = line.GetEndPoint(1)
		
		startPointXY = ('%.2f' % startPoint.X, '%.2f' % startPoint.Y )
		
		
		linePoints.append([startPoint, endPoint])
		
		p = 0.1
		
		MEP_ElementTop = max(startPoint.Z, endPoint.Z)
		MEP_ElementBottom = min(startPoint.Z, endPoint.Z)
		
		if MEP_ElementsToCheck == "pipes":
			MEP_ElementSystem = MEP_Element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
			MEP_ElementSystemName = MEP_Element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
		if MEP_ElementsToCheck == "ducts":
			MEP_ElementSystem = MEP_Element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
			MEP_ElementSystemName = MEP_Element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
		if MEP_ElementsToCheck == "cabletrays":
			MEP_ElementSystem = "Cable Tray"
			MEP_ElementSystemName =  MEP_Element.get_Parameter(BuiltInParameter.RBS_CTC_SERVICE_TYPE).AsString()
		
		#MEP_ElementSystem = MEP_Element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
		
		
		
		
		deltaX = round(endPoint.X - startPoint.X, 2)
		deltaY = round(endPoint.Y - startPoint.Y, 2)
		deltaZ = round(endPoint.Z - startPoint.Z, 2)
		
		
		
		
		tolerance = 0.01
		
		orientation = ' - '
		
		if ( abs(deltaX) > tolerance or abs(deltaY) > tolerance ) and abs(deltaZ) > tolerance:
			orientation = "Sloped"
		
		
		
		if ( abs(deltaX) < tolerance and abs(deltaY) < tolerance ) and abs(deltaZ) > tolerance:
			orientation = "Vertical"
			#verticalMEP_Elements.append(MEP_Element)
			
		if ( abs(deltaX) > tolerance or abs(deltaY) > tolerance ) and abs(deltaZ) < tolerance:
			orientation = "Horizontal"
			#horizontalMEP_Elements.append(MEP_Element)
		
		
		# special treatment for cable tray!
		if MEP_ElementsToCheck == "cabletrays":
			insulation_thickness = 0  # mm
			margin = 100 # mm
			try:
				MEP_ElementWidth = MEP_Element.Width
				MEP_ElementHeight = MEP_Element.Height
			
				
				
				min_hole_width = MEP_Element.Width + 2*(insulation_thickness/304.8)  + (margin/304.8)# ft
				min_hole_height = MEP_Element.Height + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
				
				
				
			except:
				min_hole_width = MEP_Element.Diameter + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
				min_hole_height = MEP_Element.Diameter + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
				MEP_ElementDiameter = MEP_Element.Diameter
			rounded_hole_width = roundup((min_hole_width)*304.8, 50) / 304.8  # round up to nearest 50mm
			rounded_hole_height = roundup((min_hole_height)*304.8, 50) / 304.8  # round up to nearest 50mm
			
			BWIC_width = rounded_hole_width
			BWIC_height = rounded_hole_height
		
		
		# special treatment for ducts!
		if MEP_ElementsToCheck == "ducts":
			insulation_thickness = MEP_Element.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() # mm
			
			
			 # allow 100mm additional margin after accouting for insulation for pipework holes (50mm either side)
			margin = 100 # mm 
			try:
			
				MEP_ElementWidth = MEP_Element.Width
				MEP_ElementHeight = MEP_Element.Height
				
				min_hole_width = MEP_Element.Width + 2*(insulation_thickness/304.8)  + (margin/304.8)# ft
				min_hole_height = MEP_Element.Height + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
			
			except:
				min_hole_width = MEP_Element.Diameter + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
				min_hole_height = MEP_Element.Diameter + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
				
				MEP_ElementDiameter = MEP_Element.Diameter
			rounded_hole_width = roundup((min_hole_width)*304.8, 50) / 304.8  # round up to nearest 50mm
			rounded_hole_height = roundup((min_hole_height)*304.8, 50) / 304.8  # round up to nearest 50mm
			
			BWIC_width = rounded_hole_width
			BWIC_height = rounded_hole_height
		
		
		
		# special treatment for pipes!
		elif MEP_ElementsToCheck == "pipes":
			insulation_thickness = MEP_Element.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() # ft
			
			#insulation_thickness = 25  # mm
			# allow 50mm additional margin after accouting for insulation for pipework holes (25mm either side)
			margin = 50  
			
			outside_diameter = MEP_Element.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble()  #ft

			MEP_ElementWidth = outside_diameter
			MEP_ElementHeight = outside_diameter
			MEP_ElementDiameter = outside_diameter
			
			
			min_hole_size = outside_diameter + 2*insulation_thickness + (margin/304.8) # ft
			
			rounded_hole = roundup((min_hole_size)*304.8, 50) / 304.8  # round up to nearest 50mm
			
			BWIC_width = rounded_hole
			BWIC_height = rounded_hole 
						
		
		
		
		
		
		
		MEP_ElementData = [MEP_Element.ReferenceLevel.Name,
								MEP_Element.ReferenceLevel.ProjectElevation * 304.8,
								MEP_Element.Id, 
								MEP_ElementSystem, 
								MEP_ElementSystemName, 
								line.Length * 304.8,
								MEP_ElementWidth * 304.8,
								MEP_ElementHeight * 304.8 ,
								MEP_ElementDiameter * 304.8,
								insulation_thickness * 304.8,
							
								startPoint.X * 304.8, 
								startPoint.Y * 304.8, 
								startPoint.Z * 304.8, 
								endPoint.X * 304.8, 
								endPoint.Y * 304.8, 
								endPoint.Z * 304.8, 
								deltaX * 304.8, 
								deltaY * 304.8, 
								deltaZ * 304.8,
								orientation
								]

		all_MEP_Elements.append(MEP_ElementData)
		
				
		

	MF_WriteToExcel("MEP_Element Data.xlsx", "All " + MEP_ElementsToCheck, all_MEP_Elements)

	
	
	return TransactionStatus.Committed
	
def GetBWICInfo(doc):
		# Get the BWIC family to place - this should be selectable by the user and/or based on what type of penetration it will be
	BWICFamilySymbol = doc.GetElement(ElementId(1290415))
	
	
	familySymbolName = BWICFamilySymbol.Family.Name
  
	all_family_instances =    FilteredElementCollector( doc ).OfClass( FamilyInstance ).OfCategory(BuiltInCategory.OST_GenericModel)
	
	bw_penetrations = [f for f in all_family_instances if f.Name == BWICFamilySymbol.Family.Name]
	
	phase = list(doc.Phases)[-1] # get last phase of project
						
	
	for b in bw_penetrations:
		if b.Space[phase]:
			
			# # BuiltinParameters
						# # SPACE_ASSOC_ROOM_NUMBER	"Room Number"
						# # SPACE_ASSOC_ROOM_NAME	"Room Name"
			
			print b.Space[phase].get_Parameter(BuiltInParameter.SPACE_ASSOC_ROOM_NAME).AsString()
	
	return "BW Penetrations found: " + str(len(bw_penetrations))
	

