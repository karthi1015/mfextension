# -*- coding: utf-8 -*-
__title__ = 'MF Family Fixer - Import'
__doc__ = """Fix Family Issues
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

sys.path.append("\lib")

from MF_HeaderStuff import *


	
from MF_CustomForms import *

from MF_MultiMapParameters import *

import itertools
from itertools import *

from MF_ExcelOutput import *



# options = ["option 1", "option 2", "option 3"]


# test = SelectFromDoubleList.show(options,
			# title='Choose Parameter to Import',
			# width=800,
			# height=800,
													 # multiselect=False)	



# log = []

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
					



# from MF_ExcelOutput import *

# MF_WriteToExcel("TextData.xlsx", "Tags", tagData)
# MF_WriteToExcel("TextData.xlsx", "TextNotes", textNoteData)

#################################

# Read from Excel

from MF_ExcelInput import *


from System import Guid


from rpw.ui.forms import select_file

#file = select_file('Excel File (*.xlsx)|*.xlsx', 'Excel File (*.xlsm)|*.xlsm' )

#file = "C:\Users\e.green\Desktop\j6276 - Master View Template Settings2.xlsm"

#inputData = 	MF_OpenExcelAndRead(file, None, 20 )  # limit to import 20 rows of data - to see what we are dealing with


# user selects stuff - need to ask user to select column containing view template ids

#pairs = MF_MultiMapParameters(inputData)

from time import *





#inputData = 	MF_OpenExcelAndRead(file, "Filter Update" )  # now read in all of the data.. 

#importData = inputData

#headerRow = importData[0]

#paramPairs = pairs

#idColumnIndex = pairs[0][2] ## index of column containing element ids



def str2bool(v):
  return str(v).lower() in ("yes", "true", "t", "1")
  
 

def getParameterValue(fp, ft):
	if fp.StorageType == StorageType.Double:
		val = ft.AsDouble(fp)
		return val	
	elif fp.StorageType == StorageType.Integer: 	
		val = ft.AsInteger(fp)
		return val	
	elif fp.StorageType == StorageType.String: 	
		val = ft.AsString(fp)
		return val	
	elif fp.StorageType == StorageType.ElementId: 	
		val = ft.AsElementId(fp)
		return val	
	else: 
		val = ft.AsValueString(fp)
   	
   	return val	 

file = "C:\Users\e.green\Desktop\Family Data.xlsx"

#files = select_file('Revit Family File (*.rfa)|*.rfa', multiple = True )

alldocs = []
app = __revit__.Application
# for f in files:  ## modify this to point to the file in the input sheet

	# alldocs.append(app.OpenDocumentFile(f))


	
time = strftime("%Y-%m-%d %H%M%S", localtime())

path = 'Y:\Revit MEP\Revit Development\_Work in progress\FamilyParameterTest\modifedFiles' + time

from System.IO import Directory

Directory.CreateDirectory(path)

#build list of 'actions' from spreadsheet inputs
#(modify, add, delete parameters)
params = []
addParams = []
deleteParams = []


renameTypes = []

#these are the docs we to operate on
docs = []
docIds = []


modify = 	MF_OpenExcelAndRead(file, "Modify" )  

paramsToModify = modify[1:]  # skip header row

print "Modify Parameters: --------------------------"

print str(paramsToModify) 

add = 	MF_OpenExcelAndRead(file, "Add" )  

paramsToAdd = add[1:]  # skip header row

print "Add Parameters: ----------------------"

print str(paramsToAdd) 


delete = 	MF_OpenExcelAndRead(file, "Delete" )  

paramsToDelete = delete[1:]  # skip header row

print "Delete Parameters: --------------------------"

print str(paramsToDelete) 





typesToRename = MF_OpenExcelAndRead(file, "Rename Types" )  


docIndexes = []
docNames = []

log = []








if paramsToDelete:
	for item in paramsToDelete:

		if item[0]:  #check input has a value
			docIndex = item[0]
			
			docName = item[1]
		
			docIndexes.append(docIndex)
		
			paramName = item[2]
			
			paramValue = item[5]
			
			typeIndex = item[9]
			
			typeName = item[8]
			
			paramGUID = item[4]
			
			docPath = item[16]
		
			deleteParams.append([paramName, paramValue, typeIndex, typeName, docIndex, paramGUID, docPath ])
			
			
			d = docPath
			 ## do we need to open this now?
		
			#d = alldocs[int(docIndex)]
		
			if d not in docs:
				docs.append(d)
				docNames.append(docName)
				docIds.append(int(docIndex))

if paramsToModify:
	for item in paramsToModify:
	
		if item[0]:  #check input has a value
			docIndex = item[0]
			
			docName = item[1]
		
			docIndexes.append(docIndex)
		
			paramName = item[2]
			
			paramValue = item[5]
			
			typeIndex = item[9]
			
			typeName = item[8]
			
			paramGUID = item[4]
			
			pInstanceOrType = item[12]
			
			docPath = item[16]
		
			params.append([paramName, paramValue, typeIndex, typeName, docIndex, paramGUID, pInstanceOrType, docPath])
			
			
		
			#d = alldocs[int(docIndex)]
			
			d = docPath
		
			if d not in docs:
				docs.append(d)
				docNames.append(docName)
				docIds.append(int(docIndex))

if paramsToAdd:
	for item in paramsToAdd:
		
		if item[0]:  #check input has a value
			docIndex = item[0]
			
			docName = item[1]
		
			docIndexes.append(docIndex)
		
			paramName = item[2]
			
			paramValue = item[5]
			
			paramGroup = item[3]
			
			dataType = item[6]
		
			paramGUID = item[4]
			
			pInstanceOrType = item[12]
			
			docPath = item[16]
			
			addParams.append([paramName, paramValue, paramGroup, dataType, docIndex, paramGUID, pInstanceOrType, docPath ])
			
			d = docPath
		
			if d not in docs:
				docs.append(d)
				docNames.append(docName)
				docIds.append(int(docIndex))
			
			
if typesToRename:
	for item in typesToRename[1:]:  #skip first item (header row)
		
		if item[0]:  #check input has a value
			docIndex = item[0]
			
			docName = item[1]
		
			docIndexes.append(docIndex)
		
			
			
			existingTypeName = item[2]
			
			newTypeName = item[4]
			
			typeIndex = item[3]
			
		
			renameTypes.append([docIndex, docName, existingTypeName, newTypeName, typeIndex])		
			
			
		
			filePath = item[16]
			
			
			d = docPath
			
			#d = alldocs[int(docIndex)]
		
			if d not in docs:
				docs.append(d)
				docNames.append(docName)
				docIds.append(int(docIndex))




#wrap input inside a list (if not a list)
if not isinstance(docs, list): docs = [docs]
if not isinstance(params, list): params = [params]

if not isinstance(addParams, list): addParams = [addParams]

if not isinstance(deleteParams, list): deleteParams = [deleteParams]

if not isinstance(renameTypes, list): renameTypes = [renameTypes]


print "AddParams: ---------------"
print str(addParams)
print "DeleteParams: ---------------"
print str(deleteParams)
print "modifyParams: ---------------"
print str(params)

#######sys.exit()

#default document set to DocumentManager.Instance.CurrentDBDocument
#if docs[0] == 'Current.Document':
#	docs = [DocumentManager.Instance.CurrentDBDocument]
#else: pass

i=0
#---DELETING PARAMETERS---#
#core data processing
log.append( ' ') 



for d in docs:
	#TransactionManager.Instance.EnsureInTransaction(doc)
	
	doc = app.OpenDocumentFile(d)
	tr = Transaction(doc, "Delete Parameters")
	tr.Start()
	#delete unwanted parameters
	
	
	for  idx, item in enumerate(deleteParams): 
			
			pName = item[0]
			pGUID = item[5]
			
			param = doc.FamilyManager.get_Parameter(pName)
			
			doc_i = int(item[4])			
						
			if param:
			
				
		
				if docIds[i] == doc_i:
					try:
						doc.FamilyManager.RemoveParameter(param)
						
						log.append( [( time),  ( doc.Title ), ('Remove Parameter'),( pName ), (pGUID),( ' Removed successfully' ) ] )
					except Exception as e:
						log.append( [( time), ( doc.Title ), ('Remove Parameter'),( pName ), (pGUID),( ' Error removing parameter: '+ str(e) ) ] )  
	i=i+1	#doc counter
	tr.Commit()
i=0
#---ADDING PARAMETERS---#
#core data processing

log.append( ' ') 
for d in docs:

	doc = app.OpenDocumentFile(d)
	#TransactionManager.Instance.EnsureInTransaction(doc)
	tr = Transaction(doc, "Add Parameters")
	tr.Start()
	
	
	#add parameters first
	for  idx, item in enumerate(addParams): 
		pName = item[0]
		pgName = item[2]
		
		pValue = item[1]
		
		pGUID = item[5]
		
		pInstanceOrType = item[6]
		
		pIsInstance = True
		if pInstanceOrType == "Type":
			pIsInstance = False
		
		pType = System.Enum.Parse(clr.GetClrType(ParameterType), 'Text' ) # FIX THIS TO READ STRING FROM EXCEL
		
		pGroup = System.Enum.Parse(clr.GetClrType(BuiltInParameterGroup), pgName )
		
		if pGUID != " - ":  ## Use ExternalDefinition method to add the shared parameter with specified GUID
			opt = ExternalDefinitionCreationOptions(pName, pType)
			
			spfilepath = path + '\SP.txt'
			
			f = open(spfilepath,'w')
			f.close()
			
			app.SharedParametersFilename = spfilepath
			
			defFile = app.OpenSharedParameterFile()
			
			tempGroup = defFile.Groups.Create(pName)
			defs = tempGroup.Definitions
			
			opt = ExternalDefinitionCreationOptions(pName, pType)
			opt.GUID = Guid(pGUID)
			#opt.GUID = pGUID
			
			
			
			#pExtDfn = Definitions.Create(opt, True)
			
			pExtDfn = defs.Create(opt)
		
		
		
		
		doc_i = int(item[4])
		
		if docIds[i] == doc_i:
			
			
			
			try:
				if pGUID != " - ":
					# if Shared Parameter
					doc.FamilyManager.AddParameter(pExtDfn, pGroup, pIsInstance )
				else:
					doc.FamilyManager.AddParameter(pName, pGroup, pType, pIsInstance )	
				
				param = doc.FamilyManager.get_Parameter(pName)
				doc.FamilyManager.Set(param, pValue)   ## this does not handle Types yet.... 
				
				log.append( [( time),  ( doc.Title ), ('Add Parameter'),( pName ), (pGUID),( ' Parameter added successfully'  ) ] )
			except Exception as e:
				
				log.append( [( time),  ( doc.Title ), ('Add Parameter'),( pName ), (pGUID),( ' Error adding parameter: '+str(e) ) ] )
				
	
	i=i+1	#doc counter
	tr.Commit()
i=0
#---MODIFYING PARAMETERS---#
#core data processing
log.append( ' ') 

for d in docs:

	doc = app.OpenDocumentFile(d)
	#TransactionManager.Instance.EnsureInTransaction(doc)	
	tr = Transaction(doc, "Modify Parameters")
	tr.Start()
	
	log.append( ' ') 
	for  idx, item in enumerate(params): 
			param = doc.FamilyManager.get_Parameter(item[0])
			
			#value = item[1] + '__' + time
			
			value = item[1] 
			# this doesnt get the value of the current type 
			type_i = int(item[2]) - 1  # start at zero
			type_name = item[3]
			doc_i = int(item[4])  
			
			
			
			
			
			
			if param:
				try:
					#doc.FamilyManager.RemoveParameter(param)
					
					#log.append( param.Definition.Name + ': Parameter removed successfully')
					types = doc.FamilyManager.Types
					
					fTypesIterator = types.ForwardIterator()
					#fTypesIterator.Reset()
					j = 0
					#while fTypesIterator.MoveNext(): #does not seem to move to next type.. 
					
										
					
					for t in types:
					
						previousValue = 'empty'	
						newValue = 'empty'	
						
						doc.FamilyManager.CurrentType = t
						#doc.FamilyManager.CurrentType = t
						
										
						#t = fTypesIterator.Current	
						#check the value before setting it
						if t.HasValue(param):
							previousValue = getParameterValue(param, t)
							#previousValue = t.AsString(param)
						 
						#if current position in the typeset matches the type number in the input sheet
						#if  docIds[i] == doc_i and j == type_i: 
						if  docIds[i] == doc_i and t.Name == type_name: 
							
							
							#set the value						
							try:
								doc.FamilyManager.Set(param, value)
								#check the value after setting it
								if t.HasValue(param):
									newValue = getParameterValue(param, t)
									#newValue = t.AsString(param)
								
								log.append( [( time),  ( doc.Title ), ('Update Parameter'), ( param.Definition.Name ),(  'updated to: '),( value), (str( type_i)), ( t.Name ),('New Value: ') ,( newValue), ('Previous Value: '), (previousValue)] )
							except Exception as ex:		
								log.append( [( time),  ( doc.Title ), ('Update Parameter'), ( param.Definition.Name ),(  'NOT updated to: '),( value), (str( type_i)), ( t.Name ), ('Error: ' +str(ex)) ] )
													
							
						j = j+1	
				
				except Exception as e:
					log.append(' Parameter skipped: ' + str(e) )			
	tr.Commit()
#---RENAMING TYPES---#
#core data processing
log.append( ' ') 

for d in docs:

	doc = app.OpenDocumentFile(d)
	#TransactionManager.Instance.EnsureInTransaction(doc)	
	tr = Transaction(doc, "Rename Types")
	tr.Start()
	
	log.append( ' ') 
	for  idx, item in enumerate(renameTypes):
			
			
			
			doc_i = int(item[0])
			
			existingTypeName = item[2]
			newTypeName = item[3]
			
			
			#if param:
			try:
				
				types = doc.FamilyManager.Types
				
				fTypesIterator = types.ForwardIterator()
				
				j = 0
				
				
									
				
				for t in types:
				
					previousValue = 'empty'	
					newValue = 'empty'	
					
					doc.FamilyManager.CurrentType = t
					#doc.FamilyManager.CurrentType = t
					
									
					
					if  docIds[i] == doc_i and t.Name == existingTypeName: 
						
						
						#set the value						
						try:
							doc.FamilyManager.RenameCurrentType(newTypeName )
								#newValue = t.AsString(param)
							
							log.append( [( time),  ( doc.Title ), ('Rename Type'), ( existingTypeName ),(  'updated to: '),( t.Name), ( t.Name ),('New Value: ') ,( newTypeName), ('Previous Value: '), (existingTypeName)] )
						except Exception as ex:		
							log.append( [( time),  ( doc.Title ), ('Rename Type'), ( t.Name ),(  'NOT updated to: '),( newTypeName ),  ( t.Name ), ('Error: ' +str(ex)) ] )
												
						
					j = j+1	
			
			except Exception as e:
				log.append(' Parameter skipped: ' + str(e) )		
	
	
	tr.Commit()

	

	famDoc = doc
	#Overwrite option
	overwrite = SaveAsOptions()
	overwrite.OverwriteExistingFile = True
	try:
        #category = fam.FamilyCategory.Name
        #newPath = pathBuilder(path, category, subfolder)
        #famDoc = doc.EditFamily(fam)
		doc.SaveAs(path + '\\' + docNames[i] + '.rfa', overwrite)
		
		log.append( [( time), ( doc.Title ), ('File Saved'),  ( ' copied to '), ( docNames[i]), ('in folder'), (path)] )
		
		
		#doc.Close()
        
	except Exception as e:
		log.append([( time), ( doc.Title ), ('File Save Error'),  ('Error Saving File:' + str(e) ) ] )
		pass
	log.append( ' ') 
	i=i+1




	
	

#output assigned the OUT variable
#OUT = docs, docIndexes, log, deleteParams, params, paramsToAdd

print log

#MF_WriteToExcel("Family Data.xlsx", "Log", log)

