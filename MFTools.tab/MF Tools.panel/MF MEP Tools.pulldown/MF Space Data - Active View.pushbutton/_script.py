# -*- coding: utf-8 -*-
__title__ = 'MF Space Data'
__doc__ = """Space Magic
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_HeaderStuff import *

from MF_ExcelOutput import *

from MF_CheckBoundary import *

import clr
import os
import os.path as op
import pickle as pl

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

import System

from System import Array
from System import Enum

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


def is_on(a, b, c):
    "Return true if point c intersects the line segment from a to b."
    # (or the degenerate case that all 3 points are coincident)
    return (collinear(a, b, c)
            and (within(a.X, c.X, b.X) if a.X != b.X else 
                 within(a.Y, c.Y, b.Y)))

def collinear(a, b, c):
	"Return true iff a, b, and c all lie on the same line."
	return (b.X - a.X) * (c.Y - a.Y) == (c.X - a.X) * (b.Y - a.Y)

def within(p, q, r):
    "Return true iff q is between p and r (inclusive)."
    return p <= q <= r or r <= q <= p

	
def isBetween(a, b, c):
	crossproduct = (c.Y - a.Y) * (b.X - a.X) - (c.X - a.X) * (b.Y - a.Y)
	epsilon = 10
    # compare versus epsilon for floating point values, or != 0 if using integers
	if abs(crossproduct) > epsilon:
		return False

	dotproduct = (c.X - a.X) * (b.X - a.X) + (c.Y - a.Y)*(b.Y - a.Y)
	if dotproduct < 0:
		return False

	squaredlengthba = (b.X - a.X)*(b.X - a.X) + (b.Y - a.Y)*(b.Y - a.Y)
	if dotproduct > squaredlengthba:
		return False

	return True	

	####################
	
def pointOnLine(x1,y1,x2,y2,r):
	d = sqrt((x2-x1)^2 + (y2 - y1)^2) #distance
	r = n / d #segment ratio

	x3 = r * x2 + (1 - r) * x1 #find point that divides the segment
	y3 = r * y2 + (1 - r) * y1 #into the ratio (1-r):r
	
	return (x3,y3)

	
#########################################	

def fuzzyMatch(a,b, precision):
	return abs(a - b) <= precision
	
def ccw(A,B,C):
    return (C.Y-A.Y) * (B.X-A.X) > (B.Y-A.Y) * (C.X-A.X)

# Return true if line segments AB and CD intersect
def intersect(A,B,C,D):
    return ccw(A,C,D) != ccw(B,C,D) and ccw(A,B,C) != ccw(A,B,D)
	
	
def GetBoundaryGeneratingElement(boundarySegment):
			linkInstance = doc.GetElement( boundarySegment.ElementId )
			try:
					linkDoc = linkInstance.GetLinkDocument()
					linkedElementId =  boundarySegment.LinkElementId 
					generatingElement = linkDoc.GetElement(linkedElementId)
					return generatingElement
			except Exception as e:
				return str(e)	
				
###################################################				


	
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
links = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()  ## GLOBAL -- bad idea

def GetWindowTypesFromLink():
	
	
	filter = ElementCategoryFilter(BuiltInCategory.OST_Windows)
	
	builtInCats = List[BuiltInCategory]()
	builtInCats.Add(BuiltInCategory.OST_Doors)
	builtInCats.Add(BuiltInCategory.OST_Windows)
	
	filter = ElementMulticategoryFilter(builtInCats)
	
	
	linkDoc = links[0].GetLinkDocument()
	windowFamilies = FilteredElementCollector(linkDoc).WhereElementIsNotElementType().WherePasses(filter).ToElements()
	
	windowData = []
	for wt in windowFamilies:
		
		
		bbDeltaX = round((wt.GetOriginalGeometry(Options()).GetBoundingBox().Max.X - wt.GetOriginalGeometry(Options()).GetBoundingBox().Min.X) * 304.8, 0)
		bbDeltaY = round((wt.GetOriginalGeometry(Options()).GetBoundingBox().Max.Y - wt.GetOriginalGeometry(Options()).GetBoundingBox().Min.Y) * 304.8, 0)
		bbDeltaZ = round((wt.GetOriginalGeometry(Options()).GetBoundingBox().Max.Z - wt.GetOriginalGeometry(Options()).GetBoundingBox().Min.Z) * 304.8, 0)
		
		area = round((bbDeltaX * bbDeltaZ)/1000000, 2)
		
		windowData.append( ( wt.GetTypeId().IntegerValue, wt.Name,  bbDeltaX, bbDeltaY, bbDeltaZ, area ) )
	
	#windowData = [  (wt.Name, wt.GetTypeId().IntegerValue) for wt in windowFamilies]
	
	windowTypes = list(dict.fromkeys(windowData))
	windowTypeList = [wt for wt in windowTypes]
		
	
	return windowTypeList

windowTypes = 	GetWindowTypesFromLink()

windowTypeAreaDict = { k[0]:k[5] for k in windowTypes }
	
for wt in windowTypes:
	print str( wt )	

#print windowTypeAreaDict[8890489] 
	
#sys.exit()
###################################################################################################
filter = ElementCategoryFilter(BuiltInCategory.OST_Levels)

levels = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()

levels = list(levels)

#levelIds = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElementIds()

#### START LEVEL LOOP HERE


allLevelSpacesSummary = []
allLevelSpacesSummary.append(["Level", "Flat", "Space Id", "Space Name", "Floor Area", "Gross External Area", "Net External Wall Area", "Total Windows", "Total External Window Area", "Total Doors", "Total External Door Area"])


allLevelSpaceData = []
allLevelSpaceData.append(["Level", "Flat","Space Id", "Space Name", "Boundary Length", "Storey Height", "Gross Boundary Area", "Net Boundary Area", "Start", "End", "Boundary Normal", "Boundary Generated By Element (Id)", "Boundary Element Name", "Boundary Type", "Inserts"])


t = Transaction(doc, 'Export Space Data from Multiple Levels')
	 
t.Start()

levels = [levels[5]]  ## temp debug hack

for level in levels:

	currentLevel =  doc.ActiveView.GenLevel
	
	#currentLevel = level
	
	#levelFilter = ElementLevelFilter(levels[4].Id)

	levelFilter = ElementLevelFilter(currentLevel.Id)



	# storeyheight = height from current level to level above 
	#storeyHeight = 2400  ## temporary hack


	levelElevationDict = {l.Id: l.Elevation for l in levels}



	spacesFilter = ElementCategoryFilter(BuiltInCategory.OST_MEPSpaces)

	
						   
						   
						   

	spaces = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(spacesFilter).WherePasses(levelFilter).ToElements()


	#spaces = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(spacesFilter).ToElements()


	#pipes = FilteredElementCollector(doc).OfCategory(Pipes)

	#TransactionManager.Instance.EnsureInTransaction(doc)
	############
	options = SpatialElementBoundaryOptions()
		
	options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.CoreCenter

	all_SpacePerimeterSegments = [s.GetBoundarySegments(options) for s in spaces]
	all_BoundaryCurves = [[[s.GetCurve() for s in segment ] for segment in segments] for segments in all_SpacePerimeterSegments ]
		
	allPolys = [ [ [(c.Evaluate(0,True).X, c.Evaluate(0,True).Y) for c in sCurves] for sCurves in boundaryCurves] for boundaryCurves in all_BoundaryCurves]
	#####################	

	allBoundaryCurves = []
	allBoundaryElements = []

	allBoundaryCurvesWithElements = []

	allNormals = []
	linePoints = []

	spaceData = []




					



					
	rooms = []	
	spaceCurves = []
	spaceCurves.append(["Space Id", "x1", "y1", "z1"])

	spaceList = []
	spaceList.append(["Level", "Flat","Space Id", "Space Name", "Boundary Length", "Storey Height", "Gross Boundary Area", "Net Boundary Area", "Start", "End", "Boundary Normal", "Boundary Generated By Element (Id)", "Boundary Element Name", "Boundary Type", "Inserts"])

	extCurves = []
	intCurves = []


	
	view = doc.ActiveView 
	allSpacesSummary = []
	allSpacesSummary.append( ["Level", "Flat", "Space Id", "Space Name", "Floor Area", "Gross External Area", "Net External Wall Area", "Total Windows", "Total External Window Area", "Total Doors", "Total External Door Area"])
	for space in spaces:
		
		spaceExteriorWindowCount = 0
		spaceExteriorWindowArea = 0
							
		spaceExteriorDoorCount = 0
		spaceExteriorDoorArea = 0
		
		spaceExteriorWallArea = 0   ## (net of windows and doors)
		spaceGrossExteriorWallArea = 0
		
		if space.Location: # placed spaces only
			
			thisSpaceCurves = []
			
			spaceFloorArea = space.Area * 0.092903  # convert sq ft to m2
			
			#spaceBoundingElementIds = space.GetGeneratingElementIds()
			
			#spaceBoundingElements = [doc.GetElement(e) for e in spaceBoundingElementIds]
			
			
			
			rooms.append(space.Room)
			
			normals = []
			options = SpatialElementBoundaryOptions()
			
			options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.CoreCenter
		
			boundaries = space.GetBoundarySegments(options)  # how to get walls that generated these?!
			
			boundaryCurves = [[b.GetCurve() for b in boundary ] for boundary in boundaries]
			
			bElements = [[GetBoundaryGeneratingElement(b) for b in boundary] for boundary in boundaries]
			
			bCurvesWithElements = [ [ [b.GetCurve(), GetBoundaryGeneratingElement(b), b] for b in boundary] for boundary in boundaries]
			
			elementInserts = "-"
			
			
				
			for item in bCurvesWithElements:
			## create a list here
				list = []
				print currentLevel.Name + "--------- Space: " + str( space.Id ) + "--------------------------------------------"
				
				bCurves = [b[0] for b in item]
				bBoundaries = [b[2] for b in item]
				bElements = [b[1] for b in item]
				beCount = 0
				storeyHeight = 3060 # default
				
			
				
				
				
				for bE in bElements:
					## some of these elements are strings - error messages.. 
					c = bCurves[beCount]
					
					b = bBoundaries[beCount]
					beCount += 1
					
				
					
					bInsertArea = 0
					
					netBoundaryArea = 0
					netExteriorBoundaryArea = 0
					
					grossExteriorBoundaryArea = 0
					
					boundaryExteriorWindowCount =0
					boundaryExteriorWindowArea =0
							
									
					boundaryExteriorDoorCount =0
					boundaryExteriorDoorArea =0
					
					
				#try:
				
					if MF_ExteriorBoundaryCheck(b, space, allPolys):
						bType = "Exterior"
						extCurves.append(c)
					else:
						bType = "Interior"
						intCurves.append(c)
				
				#except Exception as e: 
					#bType =  str(b) + " : " + str(e)				
						
						
					spaceFlatName = ' - '
					try:
						#spaceName = space.Name
						spaceName = space.get_Parameter(BuiltInParameter.SPACE_ASSOC_ROOM_NAME).AsString()
						spaceFlatName = space.LookupParameter("Flat Name").AsString()
					
					except Exception as e: 
						spaceName =  str(e)
					
						
					try:
						bEId = bE.Id
					except Exception as e: 
						bEId = str(bE) + str(e)
					try:
						cLength = c.Length * 304.8  ## feet to mm
						
					except Exception as e: 
						cLength = str(e)	
						
					try:
						
						storeyHeight = 2400   ## storeyHeight already in mm
					except Exception as e: 
						storeyHeight = str(e)		
						
					try:
						
						bArea = round( cLength * storeyHeight / 1000000 , 2)   ## storeyHeight already in mm
					except Exception as e: 
						bArea = str(e)	
					
					try:
						direction = (c.Evaluate(1,True) - c.Evaluate(0,True)).Normalize()
						cNormal = XYZ.BasisZ.CrossProduct(direction).Normalize()
						#next_c = bCurves[beCount+1]
						
						cStart = c.Evaluate(0,True)
						cEnd = c.Evaluate(1,True)
						
					except Exception as e: 
						cNormal = str(e)	
						cStart = str(e)
						cEnd = str(e)
					
					try:
						bEName = bE.Name
					except Exception as e: 
						bEName= str(e)


					#inserts = []
					
					## handle Columns #############################################################
					column = False
					belongsToRoom = False
					if "Wall" not in bEName:   ## terrible hack
						column = True
						inserts = ["Column: " + str(column) ]
					else:
					
						try:
							linkInstance = doc.GetElement( bE.Id )
							linkDoc = links[0].GetLinkDocument()
							inserts = []
							
							
							
							for i in bE.FindInserts(False,False,False,False):
							
								insert = linkDoc.GetElement(i)
								
								
								
								## does insert belong to this room? 
								
								insertType = linkDoc.GetElement(insert.GetTypeId())
								width = " -width- "
								try:
									
									phases = linkDoc.Phases
									
									phase = phases[phases.Size - 1] # retrieve the last phase of the project
								
									
									fromRoom = insert.FromRoom[phase].get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
									belongsToRoom = False
									if fromRoom == spaceName:
										belongsToRoom = True
								
								except Exception as e: 
									fromRoom = str(e)
								
								#insertOnBoundary = False 
								
								try:
									locationPoint = insert.Location.Point
									
									a = XYZ(c.Evaluate(0,True).X, c.Evaluate(0,True).Y, 0)
									b = XYZ(c.Evaluate(1,True).X, c.Evaluate(1,True).Y, 0)
									
									insertXY = XYZ(locationPoint.X, locationPoint.Y, 0)
									
									tolerance = 3.5
			
									dir = 1
			
									translation = Transform.CreateTranslation( dir * tolerance * cNormal)
									outerPoint = translation.OfPoint (locationPoint) ## is this anywhere near?
									translation = Transform.CreateTranslation( -1* dir * tolerance * cNormal)
									
									innerPoint = translation.OfPoint (locationPoint)
									normalLine = Line.CreateBound(innerPoint, outerPoint)
									
									#### draw lines #############################
			
									# drawNormal = doc.Create.NewDetailCurve(view, normalLine)
									# ogs = OverrideGraphicSettings()
									# ogs.SetProjectionLineColor(Color(255,0,255))
									# ogs.SetProjectionLineWeight(9)
									# view.SetElementOverrides(drawNormal.Id, ogs)
									
																
									
									
									insertOnBoundary = intersect(a,b,innerPoint,outerPoint)
									onBoundary = insertOnBoundary
									
								except Exception as e: 
									locationPoint = str(e)
									insertOnBoundary = str(e)
									
								
								try:
									
									
									
									params = insertType.Parameters
									paramNames = [p.Definition.Name for p in params]
									rough_width = MF_GetTypeParameterValueByName(insertType, "Rough Width")
									width = MF_GetTypeParameterValueByName(insertType, "Width")
									sec_sill_height = MF_GetTypeParameterValueByName(insertType, "Secondary sill hight")
									
									dims = "Rough Width : " + str(rough_width) + " --- Width : " + str(width) + " --- Sill Height : " + str(sec_sill_height) 
									
									#area = windowTypeAreaDict[insert.Id.IntegerValue]
									
								except Exception as e: 
									width = str(e)
								
							
								#inserts = [(insert.Id.IntegerValue, insert.Category.Name, insert.Name, dims , fromRoom ) ]
								
								## if wall has same orientation as curve
								#if belongsToRoom and bE.Orientation == normal:  ## INVESTIGATE
								#if belongsToRoom :
								insertArea = windowTypeAreaDict[insert.GetTypeId().IntegerValue]
								
								
								
								if onBoundary:
									bInsertArea += insertArea
									
									
									if bType == "Exterior":
										if insert.Category.Name == "Windows":
									
											boundaryExteriorWindowCount +=1
											boundaryExteriorWindowArea += insertArea
										if insert.Category.Name == "Doors":
											
											boundaryExteriorDoorCount +=1
											boundaryExteriorDoorArea += insertArea
										
										#netExteriorBoundaryArea += insertArea
									
									
									inserts.extend([ (insert.Id.IntegerValue, insert.Category.Name, insert.Name, insertArea,  dims , str(locationPoint), "On Boundary: " + str(insertOnBoundary), fromRoom, "Belongs to Room: " + str(belongsToRoom), "Column: " + str(column) ) ] )
								#inserts.extend([ (insert.Id.IntegerValue, insert.Category.Name, insert.Name, dims , fromRoom,) ] )
								
							try:
								netBoundaryArea = bArea - bInsertArea
							
							except Exception as e: 
								netBoundaryArea = str(e)
							
							
							
							if bType == "Exterior":
								try:
									grossExteriorBoundaryArea += bArea
								except Exception as e: 
									grossExteriorBoundaryArea = str(e)	
								
								
								
								try:						
									
									netExteriorBoundaryArea += netBoundaryArea
								
								except Exception as e: 
									netExteriorBoundaryArea = str(e)		
							
						except Exception as e: 
							
							inserts = [str(e)]
					
					
					
					
					list = [currentLevel.Name, spaceFlatName, space.Id, spaceName, cLength, storeyHeight,  bArea, netBoundaryArea, str(cStart), str(cEnd) , str(cNormal), bEId, bEName, bType, str(inserts)]	
					
					spaceList.append(list)
				
					spaceExteriorWindowCount += boundaryExteriorWindowCount
					spaceExteriorWindowArea += boundaryExteriorWindowArea
				
					spaceExteriorDoorCount += boundaryExteriorDoorCount
					spaceExteriorDoorArea += boundaryExteriorDoorArea
					
					try:
						spaceExteriorWallArea += netExteriorBoundaryArea
					except Exception as e:
						spaceExteriorWallArea = str(e)
						
					try:
						spaceGrossExteriorWallArea += grossExteriorBoundaryArea
					except Exception as e:
						spaceGrossExteriorWallArea = str(e)	
			
			spaceSummary = [currentLevel.Name, spaceFlatName, space.Id, spaceName, spaceFloorArea, spaceGrossExteriorWallArea, spaceExteriorWallArea, spaceExteriorWindowCount, spaceExteriorWindowArea, spaceExteriorDoorCount, spaceExteriorDoorArea]
		allSpacesSummary.append(spaceSummary)	
	
allLevelSpaceData.append(spaceList[1:])
allLevelSpacesSummary.append(allSpacesSummary[1:])	

#### end level loop here ################################################

			#print "bCurvesWithElements: " + str(bCurvesWithElements)
		
		#print "Space: " + space.Id + " : Inserts" + str(bCurvesWithElements[2])
		
		
		# #boundaryElements = [[doc.GetElement( b.ElementId )for b in boundary ] for boundary in boundaries]
		# boundaryElements = []
		# for bnd in boundaries:
			# for segment in bnd:
				# linkInstance = doc.GetElement( segment.ElementId )
				# try:
					# linkDoc = linkInstance.GetLinkDocument()
					# linkedElementId =  segment.LinkElementId 
					# generatingElement = linkDoc.GetElement(linkedElementId)
					# boundaryElements.append(generatingElement.Name)
				# except Exception as e:
					# boundaryElements.append(str(e))
		
		
		
		# allBoundaryElements.append(boundaryElements)
		
		# allBoundaryCurves.extend(boundaryCurves)
		
		# allBoundaryCurvesWithElements.extend(bCurvesWithElements)
		
		# boundaryPoints = [[( b.GetCurve().Evaluate(0,True).X ,b.GetCurve().Evaluate(0,True).Y, b.GetCurve().Evaluate(0,True).Z ) for b in boundary ] for boundary in boundaries]
		
		# for b in boundaryPoints:
			# for xyz in b:
				# thisSpaceCurves.append([space.Id,  xyz[0], xyz[1], xyz[2] ])
			# thisSpaceCurves.append(thisSpaceCurves[0])  # close the curve with the first point again 
		# spaceCurves.extend(thisSpaceCurves)
		
		
		
		# midpoints = [[c.Evaluate(0.5, True) for c in boundary] for boundary in boundaryCurves]
		
	
		
		
		# spaceData.append([str(space.Id),  space.Level.Name,  " (" + str(space.Location.Point.X) + ', ' + str(space.Location.Point.Y) + ")", str(boundaryPoints)  ])

		
		
t.Commit()

##################################

		
#for level in levels:
	#print level.Name + " ---- Elevation: " + str(level.ProjectElevation)
###########	

#print str(spaceList)

#print str(rooms)
# t = Transaction(doc, 'Draw Space Boundaries as Detail Lines')
 
# t.Start()
# view = doc.ActiveView 

# allCurves = [[c for c in sCurves] for sCurves in allBoundaryCurves]

# #allCurves = []

# allPolygons = [ [ (c.Evaluate(0,True).X, c.Evaluate(0,True).Y) for c in sCurves] for sCurves in allBoundaryCurves ]
# allExteriorCurves = []
# allInteriorCurves = []

# allExteriorBoundaryElements = []
# allInteriorBoundaryElements = []


# exteriorCurves = []
# spaceCount = 0
# curveCount = 0
# interiorCurves = []

# exteriorBoundaryElements = []
# interiorBoundaryElements = []


# #for sCurves in allBoundaryCurves:
# bcCount = []
# for item in allBoundaryCurvesWithElements:
	# sCurves = [sc[0] for sc in item]
	# sElements = [se[1] for se in item]
	# #sElementInserts = [sei[2] for sei in item]

	# spaceCount += 1
	# spaceBoundingElements = []
	# exteriorCurves = []
	# #interiorCurves = []
	# ##count = 0
	
	
	# for c in sCurves:
		# #curveCount += 1
		# #allCurves.append(c)
		
		
		# thisCurve = c
		# detailLine = doc.Create.NewDetailCurve(view, c)
		# ogs = OverrideGraphicSettings()
		
		# ogs.SetProjectionLineColor(Color(0,0,255))
		
		# view.SetElementOverrides(detailLine.Id, ogs)
		
		# direction = (c.Evaluate(1,True) - c.Evaluate(0,True)).Normalize()
		# normal = XYZ.BasisZ.CrossProduct(direction).Normalize()
		# startpoint = c.Evaluate(0,True)
		# midpoint = c.Evaluate(0.5,True)
		# endpoint = c.Evaluate(1,True)
		
		# tolerance = 3.5
		
		# dir = 1
		
		# translation = Transform.CreateTranslation( dir * tolerance * normal)
		# newPoint = translation.OfPoint (midpoint) ## is this anywhere near?
		
		# polygon = [ (l.Evaluate(0,True).X, l.Evaluate(0,True).Y) for l in sCurves]
		# # check if the new translated point is inside the current space boundary
		# if point_inside_polygon(newPoint.X, newPoint.Y, polygon):
			# # if yes, move it the opposite way
			# dir = -1
			# translation = Transform.CreateTranslation( dir * tolerance  * normal)
			# newPoint = translation.OfPoint (midpoint)
		# count = 0
		# # check if the point is in any of the other polygons
		# polygonCount = 0
		# for poly in allPolygons:
			# polygonCount += 1
			# if point_inside_polygon(newPoint.X, newPoint.Y, poly):
				# count +=1
				# #print "Point from curve " + str(curveCount) + " on space " + str(spaceCount) + " is contained in polygon " + str(polygonCount)
				# interiorCurves.append(c)
				# interiorBoundaryElements.append(sElements[sCurves.index(c)])
				# #interiorBoundaryElements.append(e)
				# break
		# if count > 0:
			# temp = 0
		
			
			
		# else:
			# exteriorCurves.append(c)
			# exteriorBoundaryElements.append(sElements[sCurves.index(c)])
			# #exteriorBoundaryElements.append(e)
			
			# normalLine = Line.CreateBound(midpoint, newPoint)
		
			# drawNormal = doc.Create.NewDetailCurve(view, normalLine)
			# ogs = OverrideGraphicSettings()
			# ogs.SetProjectionLineColor(Color(255,0,255))
			# ogs.SetProjectionLineWeight(9)
			# view.SetElementOverrides(drawNormal.Id, ogs)
			
			# exteriorWallLine = Line.CreateBound(startpoint, endpoint)
			
			# # exteriorLine = doc.Create.NewDetailCurve(view, c)
			
			# # ogs = OverrideGraphicSettings()
			# # ogs.SetProjectionLineColor(Color(0,255,0))
			# # ogs.SetProjectionLineWeight(12)
			# # view.SetElementOverrides(exteriorLine.Id, ogs)
			
			# ########################
			
			# remotePoint = Transform.CreateTranslation( dir * 15  * normal).OfPoint(midpoint)
			
			# ray = Line.CreateBound(midpoint, remotePoint)
			
			# intersectCount = 0
			# for lines in allCurves:
				# for l in lines:
					# A = ray.Evaluate(0,True)
					# B = ray.Evaluate(1,True)
					# C = l.Evaluate(0,True)
					# D = l.Evaluate(1,True)
					
					# if intersect(A,B,C,D ):
						# intersectCount += 1
				
			# if intersectCount <= 1:
			
				# drawRay = doc.Create.NewDetailCurve(view, ray)
				# ogs.SetProjectionLineColor(Color(255,0,0))
				# ogs.SetProjectionLineWeight(6)
				# view.SetElementOverrides(drawRay.Id, ogs)
				
				# exteriorLine = doc.Create.NewDetailCurve(view, c)
			
				# ogs = OverrideGraphicSettings()
				# ogs.SetProjectionLineColor(Color(0,255,0))
				# ogs.SetProjectionLineWeight(12)
				# view.SetElementOverrides(exteriorLine.Id, ogs)
					
		# curveCount += 1
		
	# allExteriorCurves.extend(exteriorCurves)	
	# allInteriorCurves.extend(interiorCurves)
	
	# allExteriorBoundaryElements.extend(exteriorBoundaryElements)
	# allInteriorBoundaryElements.extend(interiorBoundaryElements)
	
	
	# #allBoundingElements.extend(spaceBoundingElements)

# #print 'Polygon count: ' + str( len(allPolygons) )

# #print str(allExteriorBoundaryElements)	

# #allExteriorBoundaryElementIds = [e.Id for e in allExteriorBoundaryElements]


# #print str(allExteriorCurves)

# for element in allExteriorBoundaryElements:
	# # try:
		# # print "Inserts: " + str( element.FindInserts(False, False, False, False) )
	# # except Exception as e:
		# # print str(element) + "  :  " + str(e)
	
	# ogs = OverrideGraphicSettings()
	# ogs.SetProjectionLineColor(Color(0,255,0))
	# try:
		# temp = 0  ## disable this bit for now
		# #doc.ActiveView.SetElementOverrides(element.Id, ogs)
		# #doc.ActiveView.HideElements(element.Id)  # expected Icollection of ElementId
	# except Exception as e:
		# print str(element) + "  :  " + str(e)

# #doc.ActiveView.HideElements(allExteriorBoundaryElementIds)		

# # for c in allCurves:
	# # if c not in interiorCurves:
		# # exteriorLine = doc.Create.NewDetailCurve(view, c)
		# # ogs.SetProjectionLineColor(Color(0,255,0))
		# # ogs.SetProjectionLineWeight(6)
		
	# # view.SetElementOverrides(exteriorLine.Id, ogs)	
# t.Commit()		


MF_WriteToExcel("SpaceData5", "allLevelSpaceData", spaceList)

MF_WriteToExcel("SpaceData6", "allLevelSpacesSummary", allSpacesSummary)


# excel = Excel.ApplicationClass()   

# from System.Runtime.InteropServices import Marshal

# excel = Marshal.GetActiveObject("Excel.Application")

# excel.Visible = True
# excel.DisplayAlerts = False   

# ###################################

# #filename = 'C:\Users\e.green\Desktop\SheetListDataExport.xlsx'

# desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 

# filename = desktop + '\SpaceData1.xlsx'

# # finding a workbook that's already open

# workbooks = [wb for wb in excel.Workbooks if wb.FullName == filename]
# if workbooks:
    # workbook = workbooks[0]
# else:
# #Workbooks
# #if workbook exists, try to open it
	# try:
		# workbook = excel.Workbooks.Open(filename)
	# except:
		# # if not, create a new one
		# workbook = excel.Workbooks.Add()
		# #save it with the desired name
		# workbook.SaveAs(filename)

		# # oopen it
		# workbook = excel.Workbooks.Open(filename)




# ws = workbook.Sheets.Item["SpaceData"]

# ######################################################

# def ColIdxToXlName(idx):
    # if idx < 1:
        # raise ValueError("Index is too small")
    # result = ""
    # while True:
        # if idx > 26:
            # idx, r = divmod(idx - 1, 26)
            # result = chr(r + ord('A')) + result
        # else:
            # return chr(idx + ord('A') - 1) + result


# #ws = workbook.Worksheets.Add()


# #############################################################################
# ### Objects Visibile in View



# exportData = spaceList


# lastRow = len(exportData)
# #lastColumn = len(exportData[1])

# totalColumns = len(max(exportData,key=len))

# #totalColumns = 4# temporary hack

# lastColumn = totalColumns

# lastColumnName = ColIdxToXlName(totalColumns)

# xlrange = ws.Range["A1", lastColumnName+str(lastRow)]

# a = Array.CreateInstance(object, len(exportData),totalColumns)

# #exportData[1:] = sorted(exportData[1:],key=lambda x: x[1])  # ignore header row

# i = 0 


# while i < lastRow:
	# j = 0
	# while j < totalColumns:
	
		# a[i,j] = exportData[i][j]
		# j += 1
	
	
	# i += 1

# xlrange.Value2 = a 

# ws.Range(ws.Cells(1,1), ws.Cells(1,lastColumn)).Font.Bold = True
# ws.Range(ws.Cells(1,1), ws.Cells(lastRow,4)).Columns.AutoFit()
# ws.Range(ws.Cells(1,3), ws.Cells(lastRow,lastColumn)).Columns.AutoFit()
# ws.Range(ws.Cells(1,1), ws.Cells(lastRow,lastColumn)).AutoFilter()

# #workbook.Sheets.Item["SpaceDiagram"].Activate