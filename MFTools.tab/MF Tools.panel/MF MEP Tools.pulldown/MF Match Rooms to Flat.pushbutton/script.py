# -*- coding: utf-8 -*-
__title__ = 'MF Match Rooms to Flat'
__doc__ = """Flat Magic
"""

__helpurl__ = ""

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
	
def MF_SetParameterByName(el, paramName, value):
	for param in el.Parameters:
		if param.IsShared and param.Definition.Name == paramName:
			param.Set(value)
		elif param.Definition.Name == paramName:
			param.Set(value)	

#####################################

filter = ElementCategoryFilter(BuiltInCategory.OST_Levels)

levels = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()

levelFilter = ElementLevelFilter(levels[4].Id)

spacesFilter = ElementCategoryFilter(BuiltInCategory.OST_MEPSpaces)



spaces = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(spacesFilter).WherePasses(levelFilter).ToElements()

spaces = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(spacesFilter).ToElements()


#choose a level and get the Level.Id (I have just chosen one at index [1] for this example, you can use your own logic to find the Level you need or bring one in as a variable)...


lvlId = levels[4].Id


#Create ElementLevelFilter... 
filter = ElementLevelFilter(lvlId)

areaFilter = ElementCategoryFilter(BuiltInCategory.OST_Areas)

roomFilter = ElementCategoryFilter(BuiltInCategory.OST_Rooms)

links = FilteredElementCollector( doc ).OfCategory( BuiltInCategory.OST_RvtLinks )

for l in links:
	linkInstance = doc.GetElement( l.Id )
				
linkDoc = linkInstance.GetLinkDocument()

#Collect all the Elements that pass the ElementLevelFilter...
areas = FilteredElementCollector(linkDoc).OfClass(SpatialElement).WherePasses(areaFilter).ToElements()

#Collect all the Elements that pass the ElementLevelFilter...
rooms =  FilteredElementCollector(doc).OfClass(SpatialElement).WherePasses(filter).WherePasses(roomFilter).ToElements()

roomLocations = []
spaceLocations = []
placedRooms = []
placedSpaces = []
for s in spaces:
	if s.Location:
	
		sName = s.GetParameters("Name")[0].AsString()
		
		sNumber = s.GetParameters("Number")[0].AsString()
		placedSpaces.append(s)
		spaceLocations.append(s.Location.Point)
		
		#print "Space" + str(s.Location.Point) + " " + str(s.Level.Name) 

t = Transaction(doc, 'Match Spaces to Flats')
 
t.Start()

		
for a in areas:
	if a.Location:
		
		
		
		aName = a.GetParameters("Name")[0].AsString()
		
		aNumber = a.GetParameters("Number")[0].AsString()
		
		aComments = a.GetParameters("Comments")[0].AsString()
		
		
		
		aCurves = []
		
			
		options = SpatialElementBoundaryOptions()
		
		options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.CoreCenter
	
		boundaries = a.GetBoundarySegments(options)  # how to get walls that generated these?!
		
		boundaryCurves = [[b.GetCurve() for b in boundary ] for boundary in boundaries]
		boundaryPoints = [[( b.GetCurve().Evaluate(0,True).X ,b.GetCurve().Evaluate(0,True).Y, b.GetCurve().Evaluate(0,True).Z ) for b in boundary ] for boundary in boundaries]
		
		for b in boundaryPoints:
			for xyz in b:
				aCurves.append((xyz[0], xyz[1], xyz[2] ))
				
			aCurves.append(aCurves[0])  # close the curve with the first point again 
			
			polygon = [(p[0], p[1]) for p in aCurves	]
		
		spacesInArea = []
		
		
		
		for sp in spaces:
			if sp.Location and sp.Level.Name == a.Level.Name:
				if point_inside_polygon(sp.Location.Point.X, sp.Location.Point.Y, polygon):
					spacesInArea.append(sp)
					
					fNumber = a.Number.replace("Flat", "")
					
					MF_SetParameterByName(sp, "Flat Name", aNumber)
					MF_SetParameterByName(sp, "Flat Code", aName)
					MF_SetParameterByName(sp, "Flat Number", fNumber)
					try:
						sp.LookupParameter("Flat Type").Set(aComments)
					except Exception as e: 
						#print str(e)
						pass
					
					
					
					#sp.get_Parameter("Flat Type").Set(aComments)
				
		
		#print "Area: " + str(a.Location.Point) + " \t " + str(a.Level.Name) + " \t " + str(aName) + " \t " + " \t " + aNumber + " \t " + str(aCurves)
		for space in spacesInArea:
			temp = 0
			#print "Spaces: " + space.GetParameters("Name")[0].AsString()

t.Commit()			