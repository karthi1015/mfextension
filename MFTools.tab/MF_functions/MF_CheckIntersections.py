
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
wall_keys = get_rooms_from_walls()[0]
	
spaces_grouped_by_wall = get_rooms_from_walls()[1]	

print "wall_keys"
print wall_keys

print "spaces_grouped_by_wall"
print spaces_grouped_by_wall


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

@db.Transaction.ensure('Check Intersections')
def MF_CheckIntersections(MEP_ElementsToCheck, building_Element, doc):
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

	headings = ["level.Name", 
									"levelZ",
									"MEP_Element.Id", 
									"MEP_ElementSystem", 
									"MEP_ElementSystemName", 
									"length",
									"startPointXY",
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
								
	wallHeadings = ["wall.Name", 
									
									"MEP_Element.Id", 
									"MEP_ElementSystem", 
									"MEP_ElementSystemName", 
									"length",
									"startPointXY",
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
	
	
	
	
	intersections.append(headings)




	slopingMEP_Elements.append(headings)

	verticalMEP_Elements = []
	horizontalMEP_Elements = []

	all_MEP_Elements.append(headings)

	# Get the BWIC family to place - this should be selectable by the user and/or based on what type of penetration it will be
	BWICFamilySymbol = doc.GetElement(ElementId(1290415))

	wallIntersections = []	
	wallIntersections.append(wallHeadings)
	
	print "Walls found: " + str(len(walls))
		
	print "Ducts found:" + str(len(ducts))
	
	print "Pipes found:" + str(len(pipes))
	
	print "Cable Trays found:" + str(len(cable_trays	))
	
	#sys.exit()
	# for wall in walls:
	
		# if wall.Location:
			
				# #print str(wall.Location.Curve)
			
				# #A = startPoint  # MEP Curve
				# #B = endPoint  # MEP Curve
				# C = wall.Location.Curve.Evaluate(0,True)
				# D = wall.Location.Curve.Evaluate(1,True)
			
			
				# print "( "+ str( C) + ", "+ str( D ) + ")"
	
	
	#sys.exit()
	
	# build list of BWIC penetrations
		
	bwic_list  = []
	penetration_list = []
	
	headings = [
							"b.Id",
							"wall.Id",
							"wall.Name",
							"MEPElement.Id",
							"MEP_ElementSystemName",
							"location",
							"location.X",
							"location.Y",
							"location.Z",
							"BWIC_width",
							"BWIC_height",
							"wall.Orientation"]
	
	bwic_list.append(headings)
	penetration_list.append(headings)
	
	for MEP_Element in MEP_Elements:
		
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
			verticalMEP_Elements.append(MEP_Element)
			
		if ( abs(deltaX) > tolerance or abs(deltaY) > tolerance ) and abs(deltaZ) < tolerance:
			orientation = "Horizontal"
			horizontalMEP_Elements.append(MEP_Element)
		
		MEP_ElementData = [MEP_Element.ReferenceLevel.Name,
								MEP_Element.ReferenceLevel.ProjectElevation,
								MEP_Element.Id, 
								MEP_ElementSystem, 
								MEP_ElementSystemName, 
								line.Length,
								str(startPointXY),
								startPoint.X, 
								startPoint.Y, 
								startPoint.Z, 
								endPoint.X, 
								endPoint.Y, 
								endPoint.Z, 
								deltaX, 
								deltaY, 
								deltaZ,
								orientation
								]

		all_MEP_Elements.append(MEP_ElementData)
		
		#walls = list(walls)[:20]
		
		#tg = TransactionGroup(doc, "Place Wall BWICs")
		#tg.Start()
		
		
		
		for wall in walls:
		
			if wall.Location:
			
				wallMinZ = wall.get_BoundingBox(None).Min.Z
				wallMaxZ = wall.get_BoundingBox(None).Max.Z
				
				#wallLevel = linkDoc.GetElement(ElementId( wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT) ))
				
				#if orientation is not "Vertical" and wallMinZ < startPoint.Z < wallMaxZ: 
				if wallMinZ < startPoint.Z < wallMaxZ: 
				# try:
					#print str(wall.Location.Curve)
				
					# A = startPoint  # MEP Curve
					# B = endPoint  # MEP Curve
					# C = wall.Location.Curve.Evaluate(0,True)
					# D = wall.Location.Curve.Evaluate(1,True)
					
					## flatten to 2D for intersection check.. 
					
					A = XYZ(startPoint.X, startPoint.Y, 0)  # MEP Curve
					B = XYZ(endPoint.X, endPoint.Y, 0)  # MEP Curve
					C = XYZ(wall.Location.Curve.Evaluate(0,True).X, wall.Location.Curve.Evaluate(0,True).Y, 0)
					D = XYZ(wall.Location.Curve.Evaluate(1,True).X, wall.Location.Curve.Evaluate(1,True).Y, 0)
				
				
					# print str( C , D )
			
			
					if intersect(A,B,C,D):
						
						
						print wall.Name + " intersection found with " + MEP_ElementSystemName + " at Z = " + str(startPoint.Z) + "  -   Wall min : " + str(wallMinZ) + "max: " + str(wallMaxZ)
						
						intersection = [wall.Name,
											MEP_Element.Id, 
											MEP_ElementSystem, 
											MEP_ElementSystemName, 
											line.Length,
											str(startPointXY),
											startPoint.X, 
											startPoint.Y, 
											startPoint.Z, 
											endPoint.X, 
											endPoint.Y, 
											endPoint.Z, 
											deltaX, 
											deltaY, 
											deltaZ,
											orientation
											]
						wallIntersections.append(intersection)
				
						intersection = line_intersection( (A,B), (C,D) )
						
						location = XYZ(intersection[0],intersection[1], startPoint.Z )
						
						## try this
						##b = MF_PlaceBWIC(location, BWICFamilySymbol,level)
						
						# st = SubTransaction(doc)
						# st.Start()
				
						#get some info about the wall
						wall_fire_rating = ' - '
						wall_thickness = 100 / 304.8  ## default 300mm converted to feet
						# try:
							# wall_thickness = wall.LookupParameter("Width").AsValueString()
							# wall_fire_rating = wall.LookupParameter("Fire Rating").AsValueString()
						# except Exception as e:
							# print str(e)
							# pass
						
						# RBS_PIPE_OUTER_DIAMETER
						
						
						 
						try:
							MEP_ElementWidth = MEP_Element.Diameter # need to get outer diameter!
							MEP_ElementHeight = MEP_Element.Diameter
							
						except:	
							MEP_ElementWidth = MEP_Element.Width
							MEP_ElementHeight = MEP_Element.Height
				
						# # incorporate a factor to adjust hole size according to MEP Element size
						sizeFactor = 1.5
						
						BWIC_width = sizeFactor* MEP_ElementWidth
						BWIC_depth = sizeFactor* wall_thickness  # length = wall thickness for walls 
						
						
						
						BWIC_height = 	sizeFactor* MEP_ElementHeight
						
						
						# special treatment for cable tray!
						if MEP_ElementsToCheck == "cabletrays":
							insulation_thickness = 0  # mm
							margin = 100 # mm
							try:
								min_hole_width = MEP_Element.Width + 2*(insulation_thickness/304.8)  + (margin/304.8)# ft
								min_hole_height = MEP_Element.Height + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
							except:
								min_hole_width = MEP_Element.Diameter + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
								min_hole_height = MEP_Element.Diameter + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
							
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
								min_hole_width = MEP_Element.Width + 2*(insulation_thickness/304.8)  + (margin/304.8)# ft
								min_hole_height = MEP_Element.Height + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
							except:
								min_hole_width = MEP_Element.Diameter + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
								min_hole_height = MEP_Element.Diameter + 2*(insulation_thickness/304.8) + (margin/304.8) # ft
							
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
							
							min_hole_size = outside_diameter + 2*insulation_thickness + (margin/304.8) # ft
							
							rounded_hole = roundup((min_hole_size)*304.8, 50) / 304.8  # round up to nearest 50mm
							
							BWIC_width = rounded_hole
							BWIC_height = rounded_hole 
						
						
						
						
						#place marker for MEP Element - matched to nominal dimensions
						
						
						
						BWICMarkerFamilySymbol = doc.GetElement(ElementId(4337768))
						
						MEP_Marker = doc.Create.NewFamilyInstance( 
										location, 
										BWICMarkerFamilySymbol,
										level,
										Structure.StructuralType.NonStructural
										)
						mep = MEP_Marker
						
						space_data = BWIC_assign_spaces(mep, wall.Id, location)
						
						
						mep.LookupParameter("MF_Width").Set(MEP_ElementWidth)
						mep.LookupParameter("MF_Length").Set(BWIC_depth)
						mep.LookupParameter("MF_Depth").Set(MEP_ElementHeight)
						
						# MEP marker
						mep.LookupParameter("MF_BWIC Building Element Id").Set(wall.Id.IntegerValue) 
						mep.LookupParameter("MF_BWIC MEP Element Id").Set(MEP_Element.Id.IntegerValue)
						
						
						ref_level = MEP_Element.ReferenceLevel
						
						## trying this
						ref_level = levels[levelIndex]
						
						
						
						# MEP marker
						mep_level_param = mep.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
						
						mep_level_param.Set(ref_level.Id)
						
						
						offset = MEP_Element.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).AsDouble() ## this does something wierd.. 
						
						
						
						# MEP marker
						mep.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(offset)
						
						mep.LookupParameter("MF_Package").Set(MEP_ElementSystemName)
						
						
						
						
						
						# MEP marker
						
						description = MEP_ElementSystemName + " vs. " + wall.Name  + " MARKER" # + "Fire Rating: " + wall_fire_rating
						mep.LookupParameter("MF_Description").Set(description)
						
						
						#ang = GetAngleFromMEPCurve(MEP_Element)
						
						
						#create rotation axis
						zDirection = XYZ(0,0,1)
						translation = Transform.CreateTranslation(  1 * zDirection)
						newPoint = translation.OfPoint (location)
						
						rotation_axis = Line.CreateBound(location, newPoint )  # for walls we are rotating about the z axis
						#get angle of wall f
						pp = XYZ(0,1,0)  # ,north' - or facing up.. y = 1
						qq = wall.Orientation
						angle = pp.AngleTo(qq)
				
						# cant rotate element into this position - probably because it is reference plane - ?
						
						ElementTransformUtils.RotateElement(doc, mep.Id, rotation_axis, angle)
						
						## if mep element size is larger than minimum threshold
						
						
						penetration_list.append([
							str(mep.Id),
							str(wall.Id),
							wall.Name,
							str(MEP_Element.Id),						
							MEP_ElementSystemName,
							str(location),
							location.X,
							location.Y,
							location.Z,
							BWIC_width*304.8,
							BWIC_height*304.8,
							str(wall.Orientation)])
						
						
						
						
						
						minimum_mep_element_size_threshold = 50 /304.8  #  50 mm 
						
						#b ##############################################################################
				
						if (MEP_ElementWidth + (2*insulation_thickness) ) > minimum_mep_element_size_threshold:
						
							#place BWIC hole including margins etc
							BWICInstance = doc.Create.NewFamilyInstance( 
											location, 
											BWICFamilySymbol,
											level,
											Structure.StructuralType.NonStructural
											)
							b = BWICInstance
							
							# BWIC_assign_spaces(b, wall.Id, location)
							
							b_level_param = b.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
							
							b_level_param.Set(ref_level.Id)
							
							b.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(offset)
							
							b.LookupParameter("MF_Package").Set(MEP_ElementSystemName)
							
							ElementTransformUtils.RotateElement(doc, b.Id, rotation_axis, angle) # rotates BWIC family to match orientation of wall
						
							
							
							b.LookupParameter("MF_Width").Set(BWIC_width)
							b.LookupParameter("MF_Length").Set(BWIC_depth )  ## match wall thickness
							
							 
							
							
							# for walls - set the MF_Depth parameter to match the "height "
							b.LookupParameter("MF_Depth").Set(BWIC_height)  ## TODO: set this based on the thickness of the wall / floor? WALL_ATTR_WIDTH_PARAM
							
												
							
							b.LookupParameter("MF_BWIC Building Element Id").Set(wall.Id.IntegerValue) 
							b.LookupParameter("MF_BWIC MEP Element Id").Set(MEP_Element.Id.IntegerValue)
							
							description = MEP_ElementSystemName + " vs. " + wall.Name  # + "Fire Rating: " + wall_fire_rating
							
							b.LookupParameter("MF_Description").Set(description)
						
							b.LookupParameter("MF_BWIC - From Space - Id").Set(space_data[0])
							b.LookupParameter("MF_BWIC - From Space - Name").Set(space_data[1])
							b.LookupParameter("MF_BWIC - To Space - Id").Set(space_data[2])
							b.LookupParameter("MF_BWIC - To Space - Name").Set(space_data[3])
						
							bwic_list.append([
								str(b.Id),
								str(wall.Id),
								wall.Name,
								str(MEP_Element.Id),						
								MEP_ElementSystemName,
								str(location),
								location.X,
								location.Y,
								location.Z,
								BWIC_width*304.8,
								BWIC_height*304.8,
								str(wall.Orientation)])
						
						##tb.Commit()
						# print "Total Wall Intersections: " + str(len(wallIntersections))
						# except Exception as e:
							# print str(e)
							# pass
		
		
		
		# Find intersections with Levels
		
		# floor intersections ########################################
		
		#tg.Assimilate()
		
		
		for level in levels:
			levelZ = level.ProjectElevation
			
		
			
			if MEP_ElementTop > levelZ and MEP_ElementBottom < levelZ:
				intersection = [level.Name, 
									levelZ,
									MEP_Element.Id, 
									MEP_ElementSystem, 
									MEP_ElementSystemName, 
									line.Length,
									str(startPointXY),
									startPoint.X, 
									startPoint.Y, 
									startPoint.Z, 
									endPoint.X, 
									endPoint.Y, 
									endPoint.Z, 
									deltaX, 
									deltaY, 
									deltaZ,
									orientation
									]
				intersections.append(intersection)
				
				location = XYZ(startPoint.X, startPoint.Y, levelZ)
				
				
				
				# # BWICInstance = doc.Create.NewFamilyInstance( 
									# # location, 
									# # BWICFamilySymbol,
									# # level,
									# # Structure.StructuralType.NonStructural
									# # )
									
				# b = BWICInstance
				# ang = GetAngleFromMEPCurve(MEP_Element)
				
				# ElementTransformUtils.RotateElement(doc, b.Id, line, ang) # rotates BWIC family to match rotation of duct passing through it
				
				
				
				
				# # BuiltinParameters
				# # SPACE_ASSOC_ROOM_NUMBER	"Room Number"
				# # SPACE_ASSOC_ROOM_NAME	"Room Name"
				
				# phase = list(doc.Phases)[-1] # get last phase of project
				# space = b.Space[phase]
				
				## print space.Id # ERROR - always gets same space Id 
				
				# # fix this for various system types / rectangular / round duct etc
				# try:
					# MEP_ElementWidth = MEP_Element.Diameter
					# MEP_ElementHeight = MEP_Element.Diameter
				# except:	
					# MEP_ElementWidth = MEP_Element.Width
					# MEP_ElementHeight = MEP_Element.Height
				
				# # # incorporate a factor to adjust hole size according to MEP Element size
				# b.LookupParameter("MF_Width").Set(1.5* MEP_ElementWidth)
				# b.LookupParameter("MF_Length").Set(1.5* MEP_ElementHeight) 
				# b.LookupParameter("MF_Depth").Set(50/304.8)  ## TODO: set this based on the thickness of the wall / floor?
				# b.LookupParameter("Offset").Set(0)
				
				# b.LookupParameter("MF_Package").Set(MEP_ElementSystemName)
				
				
				
		
			if abs(deltaX) > 0 and abs(deltaY) > 0 and abs(deltaZ > 0):
					MEP_ElementData = [MEP_Element.ReferenceLevel.Name,
										MEP_Element.ReferenceLevel.ProjectElevation,
										MEP_Element.Id, 
										MEP_ElementSystem, 
										MEP_ElementSystemName, 
										line.Length,
										str(startPointXY),
										startPoint.X, 
										startPoint.Y, 
										startPoint.Z, 
										endPoint.X, 
										endPoint.Y, 
										endPoint.Z, 
										deltaX, 
										deltaY, 
										deltaZ,
										orientation
										]
			
					slopingMEP_Elements.append(MEP_ElementData)
	
	#dt.Commit()
	
	#print bwic_list
	
	
	#sort bwic_list
	
	sorted_bwic_list = sorted(bwic_list[1:], key = lambda x: int(x[1]))
	
	groups = []
	uniquekeys = []
	
	from  itertools import groupby
	for k, g in groupby(sorted_bwic_list, lambda x: x[1]):
		groups.append(list(g))
		uniquekeys.append(k)
	print "--- Groups ---------------------------------------"	
	print str(groups)	
		
	
	# bwic_list.append([
							# str(b.Id),
							# str(wall.Id),
							# wall.Name,
							# str(MEP_Element.Id),						
							# MEP_ElementSystemName,
							# str(location),
							# str(location.X),
							# str(location.Y),
							# str(location.Z),
							# BWIC_width*304.8,
							# BWIC_height*304.8,
							# str(wall.Orientation)])
	
	
	
	group_x_dims = []
	
	
	items_with_nearby_neighbours = []
	
	for group in groups:
		x_mins = [ g[6] - ((0.5*g[9])/304.8) for g in group] 
		x_maxs = [ g[6] + ((0.5*g[9])/304.8) for g in group] 
		
		x_centres = [g[6] for g in group]
		
		x_centres.sort()
		
		x_separations = [(x - x_centres[i - 1])*304.8 for i, x in enumerate(x_centres)][1:]
		
		#group_sorted_by_x_coord = sorted( group, lambda x: x[6] )
		
		
		#sort group by x coordinate
		# for item , n in enumerate( group_sorted_by_x_coord ):
			
			# if item[n+1][6]  - item[n][6] < 100:
				# items_with_nearby_neighbours.append(item)
			
				
		
		group_total_x = (max(x_maxs) - min(x_mins) ) * 304.8 # converting to mm
		
		# sort by x coordinate, then get separation from next one?
		
		
		group_x_dims.append((group, [group_total_x], x_separations))
		
	print "BWIC Group Information ----------------"
	
	print group_x_dims
	
	
	
	
	MF_WriteToExcel("MEP_Element Data.xlsx", MEP_ElementsToCheck + " vs. Wall ", wallIntersections)
	
	MF_WriteToExcel("MEP_Element Data.xlsx", MEP_ElementsToCheck + " vs. Floor ", intersections)
	MF_WriteToExcel("MEP_Element Data.xlsx", "Sloping " + MEP_ElementsToCheck, slopingMEP_Elements)
	MF_WriteToExcel("MEP_Element Data.xlsx", "All " + MEP_ElementsToCheck, all_MEP_Elements)
	MF_WriteToExcel("MEP_Element Data.xlsx", "BWIC - " + MEP_ElementsToCheck, bwic_list)
	MF_WriteToExcel("MEP_Element Data.xlsx", "All Penetrations - " + MEP_ElementsToCheck, penetration_list)
	
	return groups # mep elements grouped by wall
	
	
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
	

