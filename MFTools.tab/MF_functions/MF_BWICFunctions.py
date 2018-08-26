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

#@db.Transaction.ensure('Assign Spaces to BWIC wall penetrations')
def BWIC_assign_spaces(el, wall_id, loc):
############
		
		bwic_from_room_id = 0
		bwic_from_room_name = ' - '
			
		bwic_to_room_id = 0
		bwic_to_room_name = ' - '
		
		
		# get from room and to room
		
		#global wall_keys
		#global spaces_grouped_by_wall
		
		# get location
		
		try:
		
			location = loc
			
			
			
			level = doc.GetElement(el.LevelId)  ## hopefully all the items shoudl have the same level property if they have been grouped!
			
			
			facingOrientation = el.FacingOrientation
			
			
			
			#BWIC_wall_id = el.LookupParameter("MF_BWIC Building Element Id").AsInteger()
			
			BWIC_wall_id = wall_id.IntegerValue
		
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

		return [bwic_from_room_id, bwic_from_room_name, bwic_to_room_id, bwic_to_room_name]

# def assign_spaces_to_bw_holes():  #
	

	
	# ids = uidoc.Selection.GetElementIds()
	
	
	

	
	# for id in ids:
		# #el = doc.GetElement(ElementId(int(id)))
		# el = doc.GetElement(id)
		
	
		
		# BWIC_assign_spaces(el)
		
		

wall_keys = get_rooms_from_walls()[0]
	
spaces_grouped_by_wall = get_rooms_from_walls()[1]		
		
		
# assign_spaces_to_bw_holes()

@db.Transaction.ensure('Merge BWIC Holes')
def merge_bw_holes(group):  #
	
	#takes a selected set of bwic holes and replaces them with one object
	
	#get min and max coords of group
	# delta x is the new width
	# delta y ia the new height
	# centre point is new location
	
	
	# # Prompt user to select elements and to merge
	# from Autodesk.Revit.UI.Selection import ObjectType
	
	# try:
		# with forms.WarningBar(title="Select Elements to Merge"):
			# references = uidoc.Selection.PickObjects(ObjectType.Element, "Select Elements to Merge")
	# except Exception as e:
		# print str(e)
        
	# print references
	
	ids = [g[0].Id for g in group ]

	######### pre-existing selection
	
	# ids = uidoc.Selection.GetElementIds()
	
	wall_keys = get_rooms_from_walls()[0]
	
	spaces_grouped_by_wall = get_rooms_from_walls()[1]
	
	
	
	
	
	bwic_dim_list = []
	bwic_location_list = []
	
	package_list = "Combined Services: AUTO MERGED : "
	
	for id in ids:
		#el = doc.GetElement(ElementId(int(id)))
		el = doc.GetElement(id)
		
		
		
		
		
		
		
		
		
		
		# get location
		
		location = el.Location.Point
		
		bbox = el.get_BoundingBox(None)
		
		print str(bbox.Max)
		print str(bbox.Min)
		
		
		dimX = float(el.LookupParameter("MF_Width").AsDouble())  # AsValueString readse the value displayed in the GUI (in mm)
		dimY = float(el.LookupParameter("MF_Depth").AsDouble())  ## these will need to change when proper parameters are defined
		dimZ = float(el.LookupParameter("MF_Length").AsDouble())
		
		penetration_depth = float(el.LookupParameter("MF_Length").AsDouble())  ## ie. thickness of wall
		
		package = el.LookupParameter("MF_Package").AsString()
		package_list += el.LookupParameter("MF_Package").AsString() + ", "
		
		el.LookupParameter("MF_Package").Set(package + " (merged)")
		
		
		level = doc.GetElement(el.LevelId)  ## hopefully all the items shoudl have the same level property if they have been grouped!
		
		
		# if delta x == 0
		facingOrientation = el.FacingOrientation  
		
		#if abs(facingOrientation[1]) > 0.001 :
		
		x1 = location.X - dimX/2
		x2 = location.X + dimX/2
		
		x1 = bbox.Min.X
		x2 = bbox.Max.X
		
		
		
		y1 = location.Y - dimY/2
		y2 = location.Y + dimY/2
		
		y1 = bbox.Min.Y
		y2 = bbox.Max.Y
		
		
		z1 = location.Z - dimZ/2
		z2 = location.Z + dimZ/2
		
		z1 = bbox.Min.Z
		z2 = bbox.Max.Z
		
		# if abs(facingOrientation[0]) > 0.001 :
		
		# x1 = location.X - dimY/2
		# x2 = location.X + dimY/2
		
		
		
		# y1 = location.Y - dimX/2
		# y2 = location.Y + dimX/2
		
		# z1 = location.Z - dimX/2
		# z2 = location.Z + dimZ/2
		
		
		bwic_dim_list.append([x1,y1,z1,x2,y2,z2])  # this describes the 3d box of each element
		
		
		
		
		bwic_location_list.append([location.X, location.Y, location.Z])
		
		# dont need to do this here - spaces already assigned elements being merged... 
		# get the wall the e elements are hosted in though... 
		BWIC_wall_id = el.LookupParameter("MF_BWIC Building Element Id").AsInteger()
		
		# wallId = ElementId(BWIC_wall_id)
		
		# space_data = BWIC_assign_spaces(el, wallId, location)
		
		
		
	min_loc_x = min(x[0] for x in bwic_location_list)
	min_loc_y = min(x[1] for x in bwic_location_list)
	
	max_loc_x = max(x[0] for x in bwic_location_list)
	max_loc_y = max(x[1] for x in bwic_location_list)
	
	min_x = min(x[0] for x in bwic_dim_list)
	max_x = max(x[3] for x in bwic_dim_list)
	
	min_y = min(x[1] for x in bwic_dim_list)
	max_y = max(x[4] for x in bwic_dim_list)
	
	min_z = min(x[2] for x in bwic_dim_list)
	max_z = max(x[5] for x in bwic_dim_list)
	
	new_dimX = (max_x - min_x)
	new_dimY = (max_y - min_y)
	new_dimZ = (max_z - min_z)
	
	new_width = new_dimX
	new_height = new_dimZ
	
	bbMax = XYZ(max_x, max_y, max_z)
	bbMin = XYZ(min_x, min_y, min_z)
	
	bbLine = Line.CreateBound(bbMin, bbMax)
	
	#new_location = XYZ( ( min_x + max_x )/2, ( min_y + max_y )/2, ( min_z + max_z )/2 )
	
	#new_location = bbLine.Evaluate(0.5, True)
	
	new_location = XYZ( ( min_x + max_x )/2, location.Y, ( min_z + max_z )/2 )
	
	if (max_loc_x - min_loc_x <0.01):  # if all x coords are the same implies the holes are in a y axis wall
		new_width = new_dimY
		new_height = new_dimZ
		
		new_location = XYZ(  location.X, ( min_y + max_y )/2, ( min_z + max_z )/2 )
		
	
		
		# [[124.82351554450474, 38.149876866312894, 26.164698162729657, 124.96131082009529, 38.287672141903442, 26.656824146981627], [125.29267564949158, 38.149876866312887, 26.164698162729657, 125.43047092508213, 38.287672141903435, 26.656824146981627]]
		
	print 	"bwic_dim_list"
	print 	bwic_dim_list
	
	print "str([(min_x, min_y, min_z), (max_x, max_y, max_z) ])"
	print str([(min_x, min_y, min_z), (max_x, max_y, max_z) ])
	
	print "str([(min_loc_x, min_loc_y), (max_loc_x, max_loc_y)])"
	print str([(min_loc_x, min_loc_y), (max_loc_x, max_loc_y)])
	
	print "str( [ new_dimX*304.8, new_dimY*308.8, new_dimZ*304.8, new_location] )"
	print str( [ new_dimX*304.8, new_dimY*308.8, new_dimZ*304.8, new_location] )
	
	
	# Get the BWIC family to place - this should be selectable by the user and/or based on what type of penetration it will be
	BWICFamilySymbol = doc.GetElement(ElementId(1290415))
	
	#st = Subtransaction(doc, "Merge BWICs")
	#st.Start()
	
	BWICInstance = doc.Create.NewFamilyInstance( 
										new_location, 
										BWICFamilySymbol,
										level,
										Structure.StructuralType.NonStructural
										)
	
	b = BWICInstance
	
	
	
	space_data = BWIC_assign_spaces(b, ElementId(BWIC_wall_id), new_location)
	
	margin = 100 / 304.8 # 200mm converted to ft
	#margin = 0
	
	# if wall orientation is facing in the x - swap dim Y  and dim X?
	
	
	
	
	
	rounded_width = roundup((new_width + margin)*304.8, 50) / 304.8 
	# print "rounded_width"
	# print rounded_width
	
	b.LookupParameter("MF_Width").Set(rounded_width)
	#b.LookupParameter("MF_Length").Set(new_dimZ     )  ## match wall thickness
						
		
				 

						
	# for walls - set the MF_Depth parameter to match the "height "
	# this is project Z dimension
	rounded_height =  roundup((new_height + margin)*304.8, 50) / 304.8 
	
	# print "rounded_height"
	# print rounded_height
	
	b.LookupParameter("MF_Depth").Set(rounded_height )  # this the project Z direction... 
	
	

	b.LookupParameter("MF_Length").Set( penetration_depth )  # in plan for wall penetrations this is ALWAYS the wall thickness
	
	
	b.LookupParameter("MF_Package").Set(package_list)
	
	b.LookupParameter("MF_BWIC Building Element Id").Set(BWIC_wall_id) 		
	
	offset = new_location.Z - level.ProjectElevation ## this does something wierd.. 
						
	b.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(offset)
	
	############################
	zDirection = XYZ(0,0,1)
	translation = Transform.CreateTranslation(  1 * zDirection)
	newPoint = translation.OfPoint (new_location)
	
	rotation_axis = Line.CreateBound(new_location, newPoint )  # for walls we are rotating about the z axis
	#get angle of wall f
	pp = XYZ(0,1,0)  # ,north' - or facing up.. y = 1
	qq = facingOrientation
	angle = pp.AngleTo(qq)

	
	ElementTransformUtils.RotateElement(doc, b.Id, rotation_axis, angle)
	
	#uidoc.ActiveView.IsolateElementsTemporary(select_ids)
	#st.Commit()	
	
	#min_x = 	
	return True

	#############################################################

def group_nearby_points_vertical(li,vertical_tolerance_mm=500, horizontal_tolerance_mm = 500):
    
	# li is a list of elements, 
	
	li_sorted_by_z = sorted(li, key = lambda x: int(x[0].Location.Point.Z)) # sorted by z
	
	vertical_tolerance_ft = vertical_tolerance_mm / 304.8
	horizontal_tolerance_ft = vertical_tolerance_mm / 304.8
	
	out = []
	last = li_sorted_by_z[0]
	for x in li_sorted_by_z:
	
		el = x[0]
        
		loc_x = el.Location.Point.X
		loc_y = el.Location.Point.Y
		loc_z = el.Location.Point.Z
		
		last_loc_x = last[0].Location.Point.X
		last_loc_y = last[0].Location.Point.Y
		last_loc_z = last[0].Location.Point.Z
		
		
		if abs( loc_z - last_loc_z ) > vertical_tolerance_ft:
			yield out
			out = []
		out.append(x)
		last = x
	yield out	

def group_nearby_points_horizontal(li, horizontal_tolerance_mm = 500):
    
	# li is a list of elements, 
	
	li_sorted_by_x = sorted(li, key = lambda x: int(x[0].Location.Point.X)) # sorted by x
	
	
	horizontal_tolerance_ft = horizontal_tolerance_mm / 304.8
	
	out = []
	last = li_sorted_by_x[0]
	for x in li_sorted_by_x:
	
		el = x[0]
        
		loc_x = el.Location.Point.X
		loc_y = el.Location.Point.Y
		loc_z = el.Location.Point.Z
		
		last_loc_x = last[0].Location.Point.X
		last_loc_y = last[0].Location.Point.Y
		last_loc_z = last[0].Location.Point.Z
		
		
		if abs( loc_x - last_loc_x ) > horizontal_tolerance_ft:
			yield out
			out = []
		out.append(x)
		last = x
	
	
	
	
	
	
	yield out		
	
def find_mergeable_holes(groups):
	#print str(groups)	
	

	# first group by wall id

	# then group by spaces

	# then group by proximity
		
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

	from itertools import groupby



	all_x_overlaps = []
	all_y_overlaps = []
	
	all_x_overlap_holes = []
	all_y_overlap_holes = []

	separation_threshold = 500

	for group in groups:
		x_mins = [ float(g[6]) - ((0.5*g[9])/304.8) for g in group] 
		x_maxs = [ float(g[6]) + ((0.5*g[9])/304.8) for g in group] 
		
		x_centres = [g[6] for g in group]
		
		x_overlaps = []
		y_overlaps = []
		
		x_overlap_holes = []
		y_overlap_holes = []
		
		sorted_group_by_x = sorted(group, key = lambda x: x[6])
		
		sorted_group_by_y = sorted(group, key = lambda x: x[7])
		
		
		
		print "###################################################################"
		print "New Group of BWIC Objects---------------------------------------"
		i = 0
		
		
		
		
		group_string = ''
		merge_string = ''
		print "Items in group: " + str(len(sorted_group_by_x))
		
		print "sorted_group_by_x --------------------"
		print sorted_group_by_x
		
		print "sorted_group_by_y ----------------------"
		print sorted_group_by_y
		
		for i, sg  in enumerate(sorted_group_by_x):
			print "sg contains----------------------------"
			
			#sg contains----------------------------

		#dot product
		
		

		# ['4250562', '2785448', 'A-25-M3-FBA-Wall-WallType2-SingleLeaf140mmBlockTileBothSides', '2170037', 'RETURN 6', '(127.851730768, 38.218774504, 27.444225722)', 127.85173076760174, 38.218774504108154, 27.444225721784782, 187.5, 187.5, '(0.000000000, 1.000000000, 0.000000000)']



		# sg contains----------------------------



		# ['4250563', '2785448', 'A-25-M3-FBA-Wall-WallType2-SingleLeaf140mmBlockTileBothSides', '2170046', 'RETURN 6', '(132.772990610, 38.218774504, 27.444225722)', 132.77299061012144, 38.21877450410814, 27.444225721784782, 187.5, 187.5, '(0.000000000, 1.000000000, 0.000000000)']



		# sg contains----------------------------



		# ['4250564', '2785448', 'A-25-M3-FBA-Wall-WallType2-SingleLeaf140mmBlockTileBothSides', '2170097', 'RETURN 6', '(136.053830505, 38.218774504, 27.444225722)', 136.05383050513552, 38.218774504108126, 27.444225721784775, 450.00000000000006, 187.5, '(0.000000000, 1.000000000, 0.000000000)']
			
			hole = sg		# unpacking 2nd item of sublist
			
			print hole
			x_overlap_pair = []
			y_overlap_pair = []
			x_overlap_hole_pair = []
			y_overlap_hole_pair = []
			
			x_overlaps_string = ''
			
			x_centre = float(hole[6])*304.8
			
			x_min = float( hole[6])*304.8   - (float( hole[9]) / 2 )
			x_max = float( hole[6])*304.8   + (float( hole[9]) / 2 )
			
		
			
			first_hole_id = hole[0]
			first_hole = hole
			
			if i < (len(sorted_group_by_x)-1):
			
				next_hole = sorted_group_by_x[i+1]
				
				hole = next_hole
				
				this_hole_id = hole[0]
				this_hole = hole
				
				x_centre_next = float(hole[6])*304.8
				
				x_min_next = float( hole[6])*304.8  - (float( hole[9]) / 2 )
				x_max_next = float( hole[6])*304.8   + (float( hole[9]) / 2 )
				
			
				
				if ( (x_centre_next - x_centre) > 0.001 )and ((x_min_next - x_max) < separation_threshold ):
					x_overlap_pair.append([first_hole_id, this_hole_id] )
					
					x_overlap_hole_pair.append([first_hole, this_hole] )
					
					x_overlaps.extend(x_overlap_pair)		
					all_x_overlaps.extend(x_overlap_pair)
					
					x_overlap_holes.extend(x_overlap_hole_pair)
					all_x_overlap_holes.extend(x_overlap_hole_pair)
			
			
			i = i+1
		
		print "Items in group: " + str(len(sorted_group_by_y))
		for j, sgy  in enumerate(sorted_group_by_y):
			print "sg contains----------------------------"
			
			
			
			hole = sgy		# unpacking 2nd item of sublist
			
			print hole
			
			y_overlap_pair = []
			
			
			
			x_centre = float(hole[6])*304.8
			
			x_min = float( hole[6])*304.8   - (float( hole[9]) / 2 )
			x_max = float( hole[6])*304.8   + (float( hole[9]) / 2 )
			
			
			y_centre = float(hole[7])*304.8
			y_min = float( hole[7])*304.8   - (float( hole[9]) / 2 )  # work this out basis of orientation axis... 
			y_max = float( hole[7])*304.8   + (float( hole[9]) / 2 )
			
			first_hole_id = hole[0]
			first_hole = hole
			
			if j < (len(sorted_group_by_y)-1):
			
				next_hole = sorted_group_by_y[j+1]
				
				hole = next_hole
				
				this_hole_id = hole[0]
				this_hole = hole
				
				
				y_centre_next = float(hole[7])*304.8
				
				y_min_next = float( hole[7])*304.8   - (float( hole[9]) / 2 )  # work this out basis of orientation axis... 
				y_max_next = float( hole[7])*304.8   + (float( hole[9]) / 2 )
				
				
				if ( (y_centre_next - y_centre) > 0.001 ) and ((y_min_next - y_max) < separation_threshold ):
					y_overlap_pair.append([first_hole_id, this_hole_id] )
					
					y_overlap_hole_pair.append([first_hole, this_hole] )
					
					y_overlaps.extend(y_overlap_pair)		
					all_y_overlaps.extend(y_overlap_pair)
			
					y_overlap_holes.extend(y_overlap_hole_pair)
					all_y_overlap_holes.extend(y_overlap_hole_pair)
			
			j = j+1
		
		
		
		
		print "######################################################"
		print "######################################################"
		
		# print "Group of Element Ids:"
		# print group_string
		# print "Mergable Element Ids"
		# print merge_string
		
		print "######################################################"
		print "######################################################"
		print "x_overlaps:"
		
		# remove any duplicates from x_overlaps array
		
		#x_overlaps = set(x_overlaps)
		x_overlaps_string = ';'
		for pair in x_overlaps:
			#print pair
			if len(pair)>1:
				x_overlaps_string += str( pair[0]) + ";" + str( pair[1]) +"; "
			print x_overlaps_string	
		
		print "######################################################"
		
		x_centres.sort()
		
		x_separations = [(float(x) - float(x_centres[i - 1]))*304.8 for i, x in enumerate(x_centres)][1:]
		
		
		
		
			
				
		
		group_total_x = (max(x_maxs) - min(x_mins) ) * 304.8 # converting to mm
		
		# sort by x coordinate, then get separation from next one?
		
		
		group_x_dims.append((group, [group_total_x], x_separations))
		
	#print "BWIC Group Information ----------------"



	#print group_x_dims
	#print GetBWICInfo(doc)

	print "all_x_overlaps:"
		
	# remove any duplicates from x_overlaps array

	#x_overlaps = set(x_overlaps)
	x_overlaps_string = ';'
	y_overlaps_string = ';'

	select_ids = []

	for pair in all_x_overlaps:
		#print pair
		id = ElementId(int(pair[0]))
		
		select_ids.append(id)  
		if len(pair)>1:
			x_overlaps_string += str( pair[0]) + ";" + str( pair[1]) +"; "
			id = ElementId(int(pair[1]))
			select_ids.append(id)
		print x_overlaps_string	

	print "all_y_overlaps:"	
	for pair in all_y_overlaps:
		#print pair
		id = ElementId(int(pair[0]))
		
		select_ids.append(id)  
		if len(pair)>1:
			y_overlaps_string += str( pair[0]) + ";" + str( pair[1]) +"; "
			id = ElementId(int(pair[1]))
			select_ids.append(id)
		print y_overlaps_string		

		
	select_ids = List[ElementId](select_ids)
	# uidoc.Selection.SetElementIds(select_ids)
	# t = Transaction(doc, "Isolate BWICs")
	# t.Start()
	# uidoc.ActiveView.IsolateElementsTemporary(select_ids)
	# t.Commit()	

	close_bwics = [doc.GetElement(id) for id in select_ids]
	
	merge_bws = []
	
	for bw in close_bwics:
		space_pair = str(bw.LookupParameter("MF_BWIC - From Space - Id").AsInteger() ) + " - " + str(bw.LookupParameter("MF_BWIC - To Space - Id").AsInteger() ) 
		
		wall_id = bw.LookupParameter("MF_BWIC Building Element Id").AsInteger()
		
		merge_bws.append([bw, wall_id, space_pair, bw.Id, bw.Location.Point.X, bw.Location.Point.Y, bw.Location.Point.Z])
		
		#### group these by nearby coordinates... 
	
	# group merge_bws by space_pair
	
	sorted_merge_bws = sorted(merge_bws, key = lambda x: ( x[1], x[2]) )
	
	bws_grouped_by_space_pair = []
	uniquekeys = []
	
	for k, g in groupby(sorted_merge_bws, lambda x: ( x[1], x[2]) ):
		bws_grouped_by_space_pair.append(list(g))
		uniquekeys.append(k)
	
	#print bws_grouped_by_space_pair
	
	for group in bws_grouped_by_space_pair:
		print group
	
	# need to group by wall first, then group by space pair?
	# merge similar types (ducts / pipes / cable tray) - successful!!
	# similar z elevations , max delta_x and delta_z to avoid really big holes
	
	return bws_grouped_by_space_pair
	

class BoundingBox(object):
    """
    A 2D bounding box
    """
    def __init__(self, points):
        if len(points) == 0:
            raise ValueError("Can't compute bounding box of empty list")
        self.minx, self.miny = float("inf"), float("inf")
        self.maxx, self.maxy = float("-inf"), float("-inf")
        for x, y in points:
            # Set min coords
            if x < self.minx:
                self.minx = x
            if y < self.miny:
                self.miny = y
            # Set max coords
            if x > self.maxx:
                self.maxx = x
            elif y > self.maxy:
                self.maxy = y
    @property
    def width(self):
        return self.maxx - self.minx
    @property
    def height(self):
        return self.maxy - self.miny
    def __repr__(self):
        return "BoundingBox({}, {}, {}, {})".format(
            self.minx, self.maxx, self.miny, self.maxy)	
	
	
@db.Transaction.ensure('Check Intersection - Accessories')
def MF_CheckIntersectionAccessory(MEP_ElementsToCheck, building_Element, doc):
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
	
	level = selected_level
	
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
	
	filter = ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory)
	
	duct_accessories = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()
	
	duct_accessory_collector = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter)
	
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
	
	
	wallIntersections = []
	#intersection filter..
	for d in list(duct_accessories):
		print d.Symbol.Family.Name 
		print (d.Location.Point)
		
		MEP_Element = d
		MEP_ElementSystemName = MEP_Element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
		
		MEP_ElementSystem = MEP_Element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString()
		
		# create imaginary normal lines from location centre of object
		
		normal = d.HandOrientation
		
		orientation = normal
		
		centre = d.Location.Point
		tolerance = 300/304.8
		
		translation = Transform.CreateTranslation( -1 * tolerance  * normal)
		startPoint = translation.OfPoint (centre)
		
		translation = Transform.CreateTranslation( 1 * tolerance  * normal)
		endPoint = translation.OfPoint (centre)
		
		line = Line.CreateBound(startPoint, endPoint)


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
											# str(startPointXY),
											startPoint.X, 
											startPoint.Y, 
											startPoint.Z, 
											endPoint.X, 
											endPoint.Y, 
											endPoint.Z, 
											# deltaX, 
											# deltaY, 
											# deltaZ,
											orientation
											]
						wallIntersections.append(intersection)
				
						intersection = line_intersection( (A,B), (C,D) )
						
						location = XYZ(intersection[0],intersection[1], startPoint.Z )	
						
						BWICMarkerFamilySymbol = doc.GetElement(ElementId(4337768))
						
						MEP_Marker = doc.Create.NewFamilyInstance( 
										location, 
										BWICMarkerFamilySymbol,
										level,
										Structure.StructuralType.NonStructural
										)
						mep = MEP_Marker
						
						space_data = BWIC_assign_spaces(mep, wall.Id, location)
						
						wall_thickness = 150/304.8
						
						try:
							width = d.LookupParameter("Damper Width").AsDouble()
							height = d.LookupParameter("Damper Height").AsDouble()
							
							mep.LookupParameter("MF_Length").Set(wall_thickness)
						
							margin = 1
							
							mep.LookupParameter("MF_Width").Set(margin*width)
							
							mep.LookupParameter("MF_Depth").Set(margin*height)
							
						except:
							pass
						
						#ref_level = d.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
						
						## trying this
						#ref_level = levels[levelIndex]
						
						
						
						# MEP marker
						mep_level_param = mep.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
						
						mep_level_param.Set(level.Id)
						
						
						offset = MEP_Element.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble() ## this does something wierd.. 
						
						
						
						# MEP marker
						mep.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(offset)
						
						mep.LookupParameter("MF_Package").Set(MEP_ElementSystemName)
						
						mep.LookupParameter("MF_BWIC Building Element Id").Set(wall.Id.IntegerValue) 
						mep.LookupParameter("MF_BWIC MEP Element Id").Set(MEP_Element.Id.IntegerValue)
						
						
						
						# MEP marker
						
						description = d.Symbol.Family.Name + " : " + MEP_ElementSystemName + " vs. " + wall.Name  + " MARKER" # + "Fire Rating: " + wall_fire_rating
						mep.LookupParameter("MF_Description").Set(description)
						
						#duct_accessories_intersecting_wall = duct_accessory_collector.WherePasses(intersectionFilter)
	
		#print "Intersection: " + wall.Name + " : " + str( len(list(duct_accessories_intersecting_wall) ))
	
	
