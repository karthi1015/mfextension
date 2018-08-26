# -*- coding: utf-8 -*-
__title__ = 'MF Space Data - Active View'
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
horizontalPipes = []
verticalPipes = []

filter = ElementCategoryFilter(BuiltInCategory.OST_Levels)

levels = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()

#levelIds = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElementIds()

currentLevel =  doc.ActiveView.GenLevel

#levelFilter = ElementLevelFilter(levels[4].Id)

levelFilter = ElementLevelFilter(currentLevel.Id)



# storeyheight = height from current level to level above 
#storeyHeight = 2400  ## temporary hack


levelElevationDict = {l.Id: l.Elevation for l in levels}



spacesFilter = ElementCategoryFilter(BuiltInCategory.OST_MEPSpaces)

links = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()
                       
					   
					   

spaces = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(spacesFilter).WherePasses(levelFilter).ToElements()


#spaces = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(spacesFilter).ToElements()


#pipes = FilteredElementCollector(doc).OfCategory(Pipes)

#TransactionManager.Instance.EnsureInTransaction(doc)
############
options = SpatialElementBoundaryOptions()
	
options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.CoreCenter

all_SpacePerimeterSegments = [s.GetBoundarySegments(options) for s in spaces]
all_BoundaryCurves = [[[s.GetCurve() for s in segment ] for segment in segments] for segments in all_SpacePerimeterSegments ]
#### check this....	
allPolys = [ [ [(c.Evaluate(0,True).X, c.Evaluate(0,True).Y) for c in sCurves] for sCurves in boundaryCurves] for boundaryCurves in all_BoundaryCurves]


#flat_list = [item for sublist in l for item in sublist]

#flat_list = [item for sublist in l for item in sublist]

#allPolys = [ [ [(c.Evaluate(0,True).X, c.Evaluate(0,True).Y) for c in sCurves] for sCurves in boundaryCurves] for boundaryCurves in all_BoundaryCurves]

allPolys = []
allLines = []
for boundaryLoop in all_BoundaryCurves:
	for spaceCurves in boundaryLoop:
		spaceLoop = []
		for c in spaceCurves:
			allLines.append(c)
			spaceLoop.append((c.Evaluate(0,True).X, c.Evaluate(0,True).Y))
		allPolys.append(spaceLoop)
#####################	

allBoundaryCurves = []
allBoundaryElements = []

allBoundaryCurvesWithElements = []

allNormals = []
linePoints = []

spaceData = []

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

				
rooms = []	
spaceCurves = []
spaceCurves.append(["Space Id", "x1", "y1", "z1"])

spaceList = []
spaceList.append(["Level", "Flat","Space Id", "Space Name", "Boundary Length", "Height", "Gross Boundary Area", "Net Boundary Area", "Start", "End", "Boundary Normal", "Boundary Generated By Element (Id)", "Boundary Element Name", "Boundary Type", "Inserts"])

extCurves = []
intCurves = []


t = Transaction(doc, 'Mark Boundarys with Detail Lines')
 
t.Start()
view = doc.ActiveView 
drawLines = True

allSpacesSummary = []
allSpacesSummary.append( ["Level", "Flat", "Space Id", "Space Name", "Floor Area", "Height", "Gross External Area", "Net External Wall Area", "Total Windows", "Total External Window Area", "Total Doors", "Total External Door Area"])

print currentLevel.Name + " ----------All Polygons: -----  " 
print str(allPolys)
print currentLevel.Name + " ----------First Polygon: -----  "
print str(allPolys[0])

#sys.exit()

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
			print "Space: " + str( space.Id ) + "--------------------------------------------"
			
			
			
			bCurves = [b[0] for b in item]
			bBoundaries = [b[2] for b in item]
			bElements = [b[1] for b in item]
			beCount = 0
			## storeyHeight = 3060 # default
			
		
			
			
			
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
			
				if MF_ExteriorBoundaryCheck(b, space, allPolys, allLines, view, drawLines):
					bType = "Exterior"
					extCurves.append(c)
					
					##################################
					if drawLines:
						exteriorLine = doc.Create.NewDetailCurve(view, c)
						ogs = OverrideGraphicSettings()
						ogs.SetProjectionLineColor(Color(0,255,0))
						ogs.SetProjectionLineWeight(12)
						view.SetElementOverrides(exteriorLine.Id, ogs)
					#############################################
				else:
					bType = "Interior"
					intCurves.append(c)
					
					##################################
					if drawLines:
						interiorLine = doc.Create.NewDetailCurve(view, c)
						ogs = OverrideGraphicSettings()
						ogs.SetProjectionLineColor(Color(0,0,255))
						ogs.SetProjectionLineWeight(5)
						view.SetElementOverrides(interiorLine.Id, ogs)
					#############################################
			
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
					
					storeyHeight = 3060   ## storeyHeight already in mm
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
						
						# get inserts in this boundary element at the current Level
						
						# check if z coordinate of  insert location lies between the current level elevation and the height of the storey
						
						#bEInserts = [ ins for ins in bE.FindInserts(False,False,False,False) if linkDoc.GetElement(ins).Location.Point.Z is space.Level.Name]
						
						
						
						for i in bE.FindInserts(False,False,False,False):
						#for i in bEInserts:
						
							insert = linkDoc.GetElement(i)
							#insert = i
							
							
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
							
							insertOnBoundary = False 
							onBoundary = False
							try:
								locationPoint = insert.Location.Point
								
								
								
								a = XYZ(c.Evaluate(0,True).X, c.Evaluate(0,True).Y, 0)
								b = XYZ(c.Evaluate(1,True).X, c.Evaluate(1,True).Y, 0)
								
								insertXY = XYZ(locationPoint.X, locationPoint.Y, 0)
								
								tolerance = 1  ## less that the width of the insert but thicker than its host
		
								dir = 1
		
								translation = Transform.CreateTranslation( dir * tolerance * cNormal)
								outerPoint = translation.OfPoint (locationPoint) ## is this anywhere near?
								translation = Transform.CreateTranslation( -1* dir * tolerance * cNormal)
								
								innerPoint = translation.OfPoint (locationPoint)
								normalLine = Line.CreateBound(innerPoint, outerPoint)
								
								#### draw lines to mark inserts #############################
								if drawLines:
									drawNormal = doc.Create.NewDetailCurve(view, normalLine)
									ogs = OverrideGraphicSettings()
									ogs.SetProjectionLineColor(Color(0,0,255))
									ogs.SetProjectionLineWeight(6)
									view.SetElementOverrides(drawNormal.Id, ogs)
								
															
								
								
								insertOnBoundary = intersect(a,b,innerPoint,outerPoint)  ## in 2d
								# now check if on same level by checing the z elevation
								
								curveZ = c.Evaluate(0,True).Z
								
								if abs(curveZ - locationPoint.Z ) < 0.1 :   ## check if curveZ is (approx) same as insertZ
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
		
		spaceSummary = [currentLevel.Name, spaceFlatName, space.Id, spaceName, spaceFloorArea, storeyHeight,spaceGrossExteriorWallArea, spaceExteriorWallArea, spaceExteriorWindowCount, spaceExteriorWindowArea, spaceExteriorDoorCount, spaceExteriorDoorArea]
	allSpacesSummary.append(spaceSummary)	
		
t.Commit()	


MF_WriteToExcel("ActiveView - SpaceData - " + currentLevel.Name , "SpaceData", spaceList)

MF_WriteToExcel("ActiveView - AllSpaceSummary - " + currentLevel.Name, "AllSpaceSummary", allSpacesSummary)


