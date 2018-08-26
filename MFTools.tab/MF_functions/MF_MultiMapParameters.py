
from MF_HeaderStuff import *
### choose column containing element ids

from MF_WindowStuff import *

def MF_MultiMapParameters(inputData):

	headerRow = inputData[0]
	options = [x for x in headerRow if "id" in str(x).lower()]
	#
	if len(options) == 0:
		forms.alert("Element ID column not found in Excel file.. ",
						title = "Import Parameters from Excel - ID column not found"
		
		)
		sys.exit()
	
	#options.extend(headerRow)
	selectedIds = forms.SelectFromList.show(options,
				title='Choose Column Containing Element IDs',
				width=800,
				height=800,
				multiselect=False)	

		
		

	#find index of selected item	
	#print "Selected Field: " + str(selected[0])



	idColumnIndex = headerRow.index(selectedIds[0]) ## need to return this to main somehow



	# check what elements we have in the input sheet 

	sampleDataRow = inputData[1] # look at first row of data

	## choose column containing element Ids.. 

	elementIdstring = sampleDataRow[idColumnIndex]  ## temporary

	sampleElement = doc.GetElement(ElementId(int(elementIdstring)))

	sampleElementParams = sampleElement.Parameters

	remainingParams = sampleElementParams



	### choose parameter (s) to import from sheet by column heading	

	headerRow = inputData[0]
	options = []
	#

	paramPairs = []

	remainingOptions =  inputData[0]
	
	options = remainingOptions
		
	selected = forms.SelectFromList.show(options,
			title='Choose Parameter(s) to Import',
			width=800,
			height=800,
			multiselect=True)	
			
	importFields =  selected

	#remainingOptions.remove(selected[0])
	
	
	
	
	
	

	for fieldToImport in importFields:

		# options = remainingOptions
		
		# selected = forms.SelectFromList.show(options,
				# title='Choose Parameter to Import',
				# width=800,
				# height=800,
				# multiselect=False)	
				
		# fieldToImport =  str(selected[0])

		# remainingOptions.remove(selected[0])
		
		

		options = [p.Definition.Name for p in remainingParams]

		selected = forms.SelectFromList.show(options,
				title='Choose Element Parameter to Update with value from: "' + fieldToImport + '"' ,
				width=800,
				height=800,
				multiselect=False)	

		selectedParamName = str(selected[0])
		
		#remainingParams.remove(selected[0])
		
		
		
		paramPairs.append([fieldToImport, selectedParamName, idColumnIndex])


	#print str(paramPairs)
	
	return paramPairs





