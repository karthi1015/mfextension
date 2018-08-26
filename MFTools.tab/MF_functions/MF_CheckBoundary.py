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


def MF_ExteriorBoundaryCheck(boundary, space, allPolys, allLines, view, draw):
	# get curve from Boundary
	## get level of current space
	#levelFilter = ElementLevelFilter(space.Level.Id)
	## get all spaces at current level
	#spacesFilter = ElementCategoryFilter(BuiltInCategory.OST_MEPSpaces)
	#spaces = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(spacesFilter).WherePasses(levelFilter).ToElements()
	
	normals = []
	#options = SpatialElementBoundaryOptions()
	
	#options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.CoreCenter
	
	
	#all_SpacePerimeterSegments = [s.GetBoundarySegments(options) for s in spaces]
	#all_boundaryCurves = [[s.GetCurve() for s in segment ] for segment in all_SpacePerimeterSegments ]
	
	#allPolygons = [ [ (c.Evaluate(0,True).X, c.Evaluate(0,True).Y) for c in sCurves] for sCurves in all_BoundaryCurves ]
	
	
	
	c = boundary.GetCurve()
	
	direction = (c.Evaluate(1,True) - c.Evaluate(0,True)).Normalize()
	normal = XYZ.BasisZ.CrossProduct(direction).Normalize()
	startpoint = c.Evaluate(0,True)
	midpoint = c.Evaluate(0.5,True)
	endpoint = c.Evaluate(1,True)
	
	tolerance = 1
	
	rayLength = 25
	
	dir = 1
	
	translation = Transform.CreateTranslation( dir * tolerance * normal)
	newPoint = translation.OfPoint (midpoint) ## is this anywhere near?
	
	remotePoint = Transform.CreateTranslation( dir * rayLength  * normal).OfPoint(midpoint)
	
	options = SpatialElementBoundaryOptions()
	options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.CoreCenter
	
	spacePerimeterSegments = space.GetBoundarySegments(options)
	
	boundaryCurves = [[s.GetCurve() for s in segment ] for segment in spacePerimeterSegments]
	#boundaryCurves = [s.GetCurve() for s in spacePerimeterSegments]
	#print str(boundaryCurves)
	polygon = [ (l.Evaluate(0,True).X, l.Evaluate(0,True).Y) for l in boundaryCurves[0]]
	#print str(polygon)
	
	# check if the new translated point is inside the current space boundary
	if point_inside_polygon(newPoint.X, newPoint.Y, polygon):
		# if yes, move it the opposite way
		dir = -1
		translation = Transform.CreateTranslation( dir * tolerance  * normal)
		newPoint = translation.OfPoint (midpoint)
		
		remotePoint = Transform.CreateTranslation( dir * rayLength  * normal).OfPoint(midpoint)
	
	
	count = 0
	# check if the point is in any of the other polygons
	polygonCount = 0
	
	for p in allPolys:
		try:
			#poly = p[0]
			poly = p
			#print str(poly)
			polygonCount += 1
			if point_inside_polygon(newPoint.X, newPoint.Y, poly):
				count +=1
			#print "Point from curve " + str(curveCount) + " on space " + str(spaceCount) + " is contained in polygon " + str(polygonCount)
			#return "Interior"
		except Exception as e:  ## check this ################
			print "Polygon error: Space Id :" + str(space.Id) + " : " +str(e)
			pass
			
	# if count > 0:
		# temp = 0
		# return False  # interior
	# else:
		# return True # exterior
		
	# remotePoint = Transform.CreateTranslation( dir * 15  * normal).OfPoint(midpoint)
			
	ray = Line.CreateBound(midpoint, remotePoint)
	
	intersectCount = 0
	for l in allLines:
		
		A = ray.Evaluate(0,True)
		B = ray.Evaluate(1,True)
		C = l.Evaluate(0,True)
		D = l.Evaluate(1,True)
		
		if intersect(A,B,C,D ):
			intersectCount += 1
			
	
	# check if these go in the right direction	
	
	# drawRay = doc.Create.NewDetailCurve(view, ray)
	# ogs = OverrideGraphicSettings()
	# ogs.SetProjectionLineColor(Color(0,0,255))
	# ogs.SetProjectionLineWeight(9)
	# view.SetElementOverrides(drawRay.Id, ogs)	
	
	# if ray intersects with no other lines, the boundary must be exterior?		
	
	
	if intersectCount <= 1:
		exterior = 1
		if draw:
			drawRay = doc.Create.NewDetailCurve(view, ray)
			ogs = OverrideGraphicSettings()
			ogs.SetProjectionLineColor(Color(255,0,0))
			ogs.SetProjectionLineWeight(6)
			view.SetElementOverrides(drawRay.Id, ogs)	

	if count > 0 :
		temp = 0
		#return False  # interior
	else:
		normalLine = Line.CreateBound(midpoint, newPoint)
		if draw:
			drawNormal = doc.Create.NewDetailCurve(view, normalLine)
			ogs = OverrideGraphicSettings()
			ogs.SetProjectionLineColor(Color(255,0,255))
			ogs.SetProjectionLineWeight(9)
			view.SetElementOverrides(drawNormal.Id, ogs)
		
		#return True # exterior
		
	#if count < 1 and intersectCount < 1:
	if count < 1 and intersectCount <= 1:
		return True
	else:
		return False
	