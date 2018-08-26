# -*- coding: utf-8 -*-
__title__ = 'MF Pipe Data'
__doc__ = """Pipe Magic
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



from MF_CheckIntersections import MF_CheckIntersections

# dt = Transaction(doc, "Duct Intersections")
# dt.Start()

MF_CheckIntersections("ducts", 'dummy', doc)

# dt.Commit()


#sys.exit()


horizontalPipes = []
verticalPipes = []

filter = ElementCategoryFilter(BuiltInCategory.OST_PipeCurves)

pipes = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()

filter = ElementCategoryFilter(BuiltInCategory.OST_DuctCurves)

ducts = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()

filter = ElementCategoryFilter(BuiltInCategory.OST_Levels)

levels = FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(filter).ToElements()

# get all spaces

# get the boundaries of each spaces

# check for intersections of mep curves (ducts, cable tray, pipes etc) with space boundaries (which are usually generated from wall objects)

# to optimise intersection checks:

# seperte by level to reduce the number of checks

# all spaces at level n vs all mep curves at level n (check z coordinate in case reference level is incorrectly set)

# only check non vertical pipes vs walls, and non horizontal pipes against floors.. 


# group vertical pipe segments into 'pipe runs' - group together by (x1,y1) coordinates (


curves = []
linePoints = []

def fuzzyMatch(a,b, precision):
	return abs(a - b) <= precision

def MF_SetParameterByName(el, paramName, value):
	for param in el.Parameters:
		#if param.IsShared and param.Definition.Name == paramName:
		if param.Definition.Name == paramName:
			param.Set(value)

def cross(a, b):
    c = [a[1]*b[2] - a[2]*b[1],
         a[2]*b[0] - a[0]*b[2],
         a[0]*b[1] - a[1]*b[0]]

    return c

	
intersections = []
slopingPipes = []
allPipes = []



verticalPipeRuns = []

headings = ["level.Name", 
								"levelZ",
								"pipe.Id", 
								"pipeSystem", 
								"pipeSystemName", 
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
								"pipeOrientation"
							
								]
intersections.append(headings)




slopingPipes.append(headings)

verticalPipes = []
horizontalPipes = []

allPipes.append(headings)

# Get the BWIC family to place - this should be selectable by the user and/or based on what type of penetration it will be
BWICFamilySymbol = doc.GetElement(ElementId(1290415))


testXYs = [ 
			(2,1),
			(1,2),
			(3,4),
			(3.5, 4.7),
			(2,1),
			(1,2),
			(3,4),
			(3.5, 6.7),
			(2,1),
			(1,2),
			(3,4),
			(6.5, 4.7)
		


]

#sortedXYs = sorted(testXYs, lambda x: x)
sortedXYsV2 = sorted(
						testXYs, 
						key=lambda x: 
						[x[0],x[1]]
					)


sortedXYs = sorted(
						testXYs, 
						key=lambda x: 
						[x[1],x[0]]
					)	
print "Raw:"
print testXYs

print "Sorted [x[0],x[1]] : "
print sortedXYsV2

print "Sorted [x[1],x[0]]: "
print sortedXYs


from itertools import groupby



group_list = []
for key, group in groupby(sortedXYs, lambda x: [x[0], x[1]]):
        group_list.append(list(group))

print "grouped List "
print group_list




dictionary = dict(enumerate(sortedXYs)) # gest unique items

print "Dictionary dict(sortedXYs):"
print dictionary

#sys.exit()


t = Transaction(doc, __title__)
t.Start()


## convert this into a function for every MEP curve type? (pipes, duct, cabletray)
for pipe in pipes:
	
	line = pipe.Location.Curve
	
	curves.append(line)
	
	startPoint = line.GetEndPoint(0)
	endPoint = line.GetEndPoint(1)
	
	startPointXY = ('%.2f' % startPoint.X, '%.2f' % startPoint.Y )
	
	
	linePoints.append([startPoint, endPoint])
	
	p = 0.1
	
	pipeTop = max(startPoint.Z, endPoint.Z)
	pipeBottom = min(startPoint.Z, endPoint.Z)
	
	crossesLevel = "Crosses Level(s) :"
	
	pipeSystem = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString()
	pipeSystemName = pipe.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
	
	deltaX = round(endPoint.X - startPoint.X, 2)
	deltaY = round(endPoint.Y - startPoint.Y, 2)
	deltaZ = round(endPoint.Z - startPoint.Z, 2)
	
	
	
	
	tolerance = 0.01
	
	pipeOrientation = ' - '
	
	if ( abs(deltaX) > tolerance or abs(deltaY) > tolerance ) and abs(deltaZ) > tolerance:
		pipeOrientation = "Sloped"
	
	
	
	if ( abs(deltaX) < tolerance and abs(deltaY) < tolerance ) and abs(deltaZ) > tolerance:
		pipeOrientation = "Vertical"
		verticalPipes.append(pipe)
		
	if ( abs(deltaX) > tolerance or abs(deltaY) > tolerance ) and abs(deltaZ) < tolerance:
		pipeOrientation = "Horizontal"
		horizontalPipes.append(pipe)
	
	pipeData = [pipe.ReferenceLevel.Name,
							pipe.ReferenceLevel.ProjectElevation,
							pipe.Id, 
							pipeSystem, 
							pipeSystemName, 
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
							pipeOrientation
							]

	allPipes.append(pipeData)
	
	# Find intersections with Levels
	
	
	
	for level in levels:
		levelZ = level.ProjectElevation
		
	
		
		if pipeTop > levelZ and pipeBottom < levelZ:
			intersection = [level.Name, 
								levelZ,
								pipe.Id, 
								pipeSystem, 
								pipeSystemName, 
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
								pipeOrientation
								]
			intersections.append(intersection)
			
			location = XYZ(startPoint.X, startPoint.Y, levelZ)
			
			
			
			BWICInstance = doc.Create.NewFamilyInstance( 
								location, 
								BWICFamilySymbol,
								level,
								Structure.StructuralType.NonStructural
								)
								
			b = BWICInstance
			
			pipeDiameter = pipe.Diameter
			
			# incorporate a factor to adjust hole size according to pipe size
			b.LookupParameter("MF_Width").Set(1.5* pipeDiameter)
			b.LookupParameter("MF_Length").Set(1.5* pipeDiameter)
			b.LookupParameter("MF_Depth").Set(50/304.8)
			b.LookupParameter("Offset").Set(0)
			
			b.LookupParameter("MF_Package").Set(pipeSystemName)
			
			# MF_SetParameterByName(b, "MF_Width", pipeDiameter)
			# MF_SetParameterByName(b, "MF_Length", pipeDiameter)
			
	
	if abs(deltaX) > 0 and abs(deltaY) > 0 and abs(deltaZ > 0):
			pipeData = [pipe.ReferenceLevel.Name,
								pipe.ReferenceLevel.ProjectElevation,
								pipe.Id, 
								pipeSystem, 
								pipeSystemName, 
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
								pipeOrientation
								]
	
			slopingPipes.append(pipeData)
	
	
t.Commit()	

for level in levels:
	print level.Name + " ---- Elevation: " + str(level.ProjectElevation)


for vPipe in verticalPipes:
	vPipeMidPoint = vPipe.Location.Curve.Evaluate(0.5,True)
	
	if (vPipeMidPoint.X, vPipeMidPoint.Y) not in verticalPipeRuns:
		verticalPipeRuns.append((vPipeMidPoint.X, vPipeMidPoint.Y))


verticalPipeData = [x for x in allPipes[1:] if x[-1] is "Vertical"]

		
sortedVerticalPipeData = sorted(
						verticalPipeData, 
						key=lambda x: 
						[x[4], x[6] ]
					)						



MF_WriteToExcel("Pipe Data.xlsx", "Floor Intersections", intersections)
MF_WriteToExcel("Pipe Data.xlsx", "Sloping Pipes", slopingPipes)
MF_WriteToExcel("Pipe Data.xlsx", "All Pipes", allPipes)
MF_WriteToExcel("Pipe Data.xlsx", "Vertical Pipes", sortedVerticalPipeData)
					
## now group pipes in to viertical runs - eg to find stacks 					
#pipeRuns = dict(enumerate([ (x[4], x[6]) for x in sortedVerticalPipeData]))

pipeRunSet = set([(x[4], x[6]) for x in sortedVerticalPipeData]) # split this by system.. FW 1.X etc 

pipeSystemSet = set([x[4] for x in sortedVerticalPipeData])

# create list of 'sets' 
setLists = []

# group by systems

sortedVerticalPipeData_grouped = []

#group_list = []
for key, group in groupby(sortedVerticalPipeData, lambda x: x[4]):
        sortedVerticalPipeData_grouped.append(list(group))

# for item in sortedVerticalPipeData:
	# for system in pipeSystemSet:
		# setLists.append(set([s[6] for s in item if x[4] is system]))

# #pipeRunList = list(pipeRuns)

pipeSystemRunSets = []
for group in sortedVerticalPipeData_grouped:
	
	systemSet = set([x[6] for x in group])
	pipeSystemRunSets.append([group[0][4], list(systemSet)])


print pipeSystemRunSets # [ SYSTEM 1 [(x1, y1), (x2, y2), ... ]] [ SYSTEM 2 [(x1, y1), (x2, y2), ... ]]

print pipeRunSet
print pipeSystemSet
headings.append("Global Pipe Run ID")
headings.append("System Pipe Run ID")
headings.append("System Pipe Run Numeric ID")
pipeRunData = []
#pipeRunData.append(headings)
for item in sortedVerticalPipeData:
	#print x
	#print (x[4], x[6])
	match = (item[4], item[6])
	pipeRunId = list(pipeRunSet).index(match)
	
	
	pipeSystems = [p[0] for p in pipeSystemRunSets]
	
	
	systemIndex = pipeSystems.index(item[4])
	
	coordSet = pipeSystemRunSets[systemIndex][1]
	
	coordIndex = coordSet.index(item[6])
	
	#pipeSystemRunId = item[4]+':'+str(systemIndex)+"."+str(coordIndex)
	pipeSystemRunId = item[4]+"."+'{0:03}'.format(coordIndex)
	pipeSystemRunNumericId = '{0:03}'.format(systemIndex)+"."+'{0:03}'.format(coordIndex)
	
	item.append(pipeRunId)
	item.append(pipeSystemRunId)
	item.append(pipeSystemRunNumericId)
	pipeRunData.append(item)
		
verticalPipeRunData = []
verticalPipeRunData.append(headings)
verticalPipeRunData.extend(pipeRunData)
		
#print "Vertical Pipes" + str(sortedVerticalPipeData)	+ " : " +  str(len(sortedVerticalPipeData) )


MF_WriteToExcel("Pipe Data.xlsx", "Vertical Pipes Runs", verticalPipeRunData )



