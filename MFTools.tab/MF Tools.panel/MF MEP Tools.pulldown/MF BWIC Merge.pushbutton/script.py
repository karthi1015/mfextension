# -*- coding: utf-8 -*-
__title__ = 'MF BWIC Merge'
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

clr.AddReference("System")
from System.Collections.Generic import List





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


from MF_BWICFunctions import *




@db.Transaction.ensure('Merge BWIC Holes')
def merge_bw_holes():  #
	
	#takes a selected set of bwic holes and replaces them with one object
	
	#get min and max coords of group
	# delta x is the new width
	# delta y ia the new height
	# centre point is new location
	
	
	# Prompt user to select elements and to merge
	from Autodesk.Revit.UI.Selection import ObjectType
	
	try:
		with forms.WarningBar(title="Select Elements to Merge"):
			references = uidoc.Selection.PickObjects(ObjectType.Element, "Select Elements to Merge")
	except Exceptions.OperationCanceledException:
		return False
        
	#print references
	
	ids = [ref.ElementId for ref in references ]
	
	if len(list(references)) == 0:
		return False
	
	######### pre-existing selection
	
	# ids = uidoc.Selection.GetElementIds()
	
	wall_keys = get_rooms_from_walls()[0]
	
	spaces_grouped_by_wall = get_rooms_from_walls()[1]
	
	
	
	
	
	bwic_dim_list = []
	bwic_location_list = []
	
	package_list = "Combined Services: "
	
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

#tg = TransactionGroup(doc)
#tg.Start()	


while merge_bw_holes(	):
	 
	
	pass
#tg.Assimilate()
#t.Commit()	




