# -*- coding: utf-8 -*-
__title__ = 'MF Lab Test'
__doc__ = """MF Lab Test
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_HeaderStuff import *


## START HERE #############


textNotes = FilteredElementCollector(doc).OfClass(TextNote).ToElements()

tags = FilteredElementCollector(doc).OfClass(IndependentTag).ToElements()


textNoteData = []

textNoteData.append(["Text", "Element ID", "Owner View", "Text Note Type Name", "Font", "Text Size"])

tagData = []

tagData.append(["Text", "Element ID", "Owner View", "Tag Name", "Label Name", "Font", "Text Size"])



log = []

def MF_GetParameterValueByName(el, paramName):
	for param in el.Parameters:
		if param.IsShared and param.Definition.Name == paramName:
			paramValue = el.get_Parameter(param.GUID)
			return paramValue.AsString()
	        
def MF_SetParameterByName(el, paramName, value):
	for param in el.Parameters:
		#if param.IsShared and param.Definition.Name == paramName:
		if param.Definition.Name == paramName:
			param.Set(value)
					

for t in textNotes:
	#print t.Text + "\t : " + str(t.Id) + "\t : " + str(doc.GetElement(t.OwnerViewId).Name) + "\t : " + t.Name
	
	font = t.TextNoteType.get_Parameter(BuiltInParameter.TEXT_FONT).AsString()
	textSize = t.TextNoteType.get_Parameter(BuiltInParameter.TEXT_SIZE).AsValueString()
	
	textNoteData.append([(t.Text ),(str(t.Id)),( str(doc.GetElement(t.OwnerViewId).Name)),(  t.Name), (font), (textSize)])
					



for tag in tags:

	familyDoc = doc.EditFamily(doc.GetElement(tag.GetTypeId()).Family)
	
	labels = FilteredElementCollector(familyDoc).OfClass(TextElement).ToElements()
	# first label in the tag family
	# what about multiple types?
	label = labels[0]
	
	
	
	labelName = label.Name
	
	labelType = familyDoc.GetElement(label.GetTypeId())
	
	font = labelType.get_Parameter(BuiltInParameter.TEXT_FONT).AsString()
	textSize = labelType.get_Parameter(BuiltInParameter.TEXT_SIZE).AsValueString()
	
	
	tagData.append([(tag.TagText ),(str(tag.Id)),( str(doc.GetElement(tag.OwnerViewId).Name)),( tag.Name),( labelName ), (font), (textSize) ])

################################# 
# Write To Excel



from MF_ExcelOutput import *

MF_WriteToExcel("TextData.xlsx", "Tags", tagData)
MF_WriteToExcel("TextData.xlsx", "TextNotes", textNoteData)

#################################

# Read from Excel

from MF_ExcelInput import MF_ReadFromExcel

inputData = 	MF_ReadFromExcel("TextData.xlsx", "Tags")

for row in inputData:
	print str(row)
