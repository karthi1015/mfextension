# -*- coding: utf-8 -*-
__title__ = 'MF BWIC Assign Spaces'
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


from MF_ExcelOutput import *

#from MF_CheckIntersections import *

# dt = Transaction(doc, "Duct Intersections")
# dt.Start()

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


def GetBoundaryGeneratingElement(boundarySegment):
			linkInstance = doc.GetElement( boundarySegment.ElementId )
			try:
					linkDoc = linkInstance.GetLinkDocument()
					linkedElementId =  boundarySegment.LinkElementId 
					generatingElement = linkDoc.GetElement(linkedElementId)
					return generatingElement
			except Exception as e:
				return False






# select a group and convert to larger hole

# prompt for selection

# with selected ids


def get_rooms_from_walls():

	# get spaces
	
	space_walls = []
	
	spacesFilter = ElementCategoryFilter(BuiltInCategory.OST_MEPSpaces)
	spaces = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(spacesFilter).ToElements()
	
	for space in spaces:
		
		options = SpatialElementBoundaryOptions()
		boundary_segments = space.GetBoundarySegments( options )
		
		for segment in boundary_segments:
		
			for s in segment:
				
				if GetBoundaryGeneratingElement(s):
					wall = GetBoundaryGeneratingElement(s)
				
					space_walls.append([wall.Id.IntegerValue, wall.Name, space.Id.IntegerValue, space.get_Parameter(BuiltInParameter.SPACE_ASSOC_ROOM_NAME).AsString() ])
				
					#print "Wall: " + str(wall_id.Id) + " --  Space: " + str(space.Id)
				
	#print "Space vs Walls:"		
	
	sorted_walls = sorted(space_walls, key = lambda x: x[0])
	
	# group spaces by wall - all spaces that share this wall
	from itertools import groupby
	
	spaces_grouped_by_walls = []
	uniquekeys = []
	
	for k, g in groupby(sorted_walls, lambda x: x[0]):
		spaces_grouped_by_walls.append(list(g))
		uniquekeys.append(k)
	
	#print "sorted_walls:"
	
	#print sorted_walls
	
	#print "spaces_grouped_by_walls:"
	
	#print spaces_grouped_by_walls
	
	return [uniquekeys, spaces_grouped_by_walls]
	
	
import math

def roundup(x, interval):
	return int(math.ceil(x/interval)) * interval

	
from rpw import db

@db.Transaction.ensure('Assign Spaces to BWIC wall penetrations')
def BWIC_assign_spaces(el):
############
		
		
		
		
		# get from room and to room
		
		global wall_keys
		global spaces_grouped_by_wall
		
		# get location
		
		try:
		
			location = el.Location.Point
			
			
			
			level = doc.GetElement(el.LevelId)  ## hopefully all the items shoudl have the same level property if they have been grouped!
			
			
			facingOrientation = el.FacingOrientation
			
			
			
			BWIC_wall_id = el.LookupParameter("MF_BWIC Building Element Id").AsInteger()
		
			space_matches = spaces_grouped_by_wall[ wall_keys.index(BWIC_wall_id) ]
			
			space_names= [x[3] for x in space_matches]
			
			space_ids = [x[2] for x in space_matches]
			
			print "Space Matches:"
			print space_names
			
			options = SpatialElementBoundaryOptions()
		
			options.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.CoreCenter
			
			#space_boundaries = [doc.GetElement(ElementId(s)).GetBoundarySegments(options) for s in space_ids]
			
			spaces_as_polygons = []
			for s in space_matches:
				space = doc.GetElement(ElementId(s[2]))
				
				space_name = s[3]
				
				space_boundary_segments = list(space.GetBoundarySegments(options))
				#print "space_boundary_segments"
				#print list(space_boundary_segments)
				
				space_curves = [[s.GetCurve() for s in segment] for segment in space_boundary_segments]
				
				
				space_polygon = [(sc.Evaluate(0,True).X, sc.Evaluate(0,True).Y) for sc in space_curves[0]]
				
				#print "space_polygon"
				#print space_polygon
				
				spaces_as_polygons.append([space.Id, space_name, space_polygon])
			
			
			
			#print "spaces_as_polygons	"
			#print spaces_as_polygons		
			
		
			
			## get polygons in X,Y to do point_in_polygon test.. 
			
			dir = 1
			tolerance = 100 / 304.8
			
			
			translation = Transform.CreateTranslation( dir * tolerance * facingOrientation)
			p1 = translation.OfPoint (location) 
			translation = Transform.CreateTranslation( -1 * dir * tolerance * facingOrientation)
			p2 = translation.OfPoint (location) 
			
			bwic_rooms = []
		
		
		
			for s in spaces_as_polygons:
				space_polygon = s[2]
				
				
				
				if point_inside_polygon( p1.X, p1.Y, space_polygon):
					bwic_rooms.append(s[1])
					bwic_to_room_id = int(s[0].IntegerValue)
					bwic_to_room_name = s[1]
				if point_inside_polygon( p2.X, p2.Y, space_polygon):
					bwic_rooms.append(s[1])
					bwic_from_room_id = int(s[0].IntegerValue)
					bwic_from_room_name = s[1]
			
			print "bwic_rooms"
			print bwic_rooms
			
			print "bwic_from_room"
			print bwic_from_room_name
			
			print "bwic_to_room"
			print bwic_to_room_name
			
			
			el.LookupParameter("MF_BWIC - From Space - Id").Set(bwic_from_room_id)
			el.LookupParameter("MF_BWIC - From Space - Name").Set(bwic_from_room_name)
			el.LookupParameter("MF_BWIC - To Space - Id").Set(bwic_to_room_id)
			el.LookupParameter("MF_BWIC - To Space - Name").Set(bwic_to_room_name)

		except Exception as e:
			print str(e)



def assign_spaces_to_bw_holes():  #
	

	
	ids = uidoc.Selection.GetElementIds()
	
	
	

	
	for id in ids:
		#el = doc.GetElement(ElementId(int(id)))
		el = doc.GetElement(id)
		
		if el.LookupParameter("MF_BWIC - From Space - Id").AsInteger() == 0 :
		
			BWIC_assign_spaces(el)
		
		

wall_keys = get_rooms_from_walls()[0]
	
spaces_grouped_by_wall = get_rooms_from_walls()[1]		
		
		
assign_spaces_to_bw_holes()


