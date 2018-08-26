# -*- coding: utf-8 -*-
__title__ = 'MF Create Filters'
__doc__ = """MF Create Filters
"""

__helpurl__ = ""
import os
import sys
## Add Path to MF Functions folder 	

# path_to_this_script = os.path.realpath(__file__)

# current_path = path_to_this_script.split("MFTools.extension")[0]

# print current_path + " MFTools.extension\MFTools.tab\MF_functions"

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_HeaderStuff import *

#print "all ok"

#import listview

#import treeview

#import combobox

#import colour_dialog

import itertools
#import xaml_test
def str2bool(v):
  return str(v).lower() in ("yes", "true", "t", "1")
  
  ############################################

def GetBuiltInParam(paramName):
	builtInParams = System.Enum.GetValues(BuiltInParameter)
	
	test = []
	
	for i in builtInParams:
		if i.ToString() == paramName:
			test.append(i)
			break
		else:
			continue
	return test[0]
	
def group(seq, sep):
    g = []
    for el in seq:
        if el == sep:
            yield g
            g = []
	g.append(el)
    yield g
	
def MF_GetParameterValueByName(el, paramName):
	for param in el.Parameters:
		if param.IsShared and param.Definition.Name == paramName:
		#if param.Definition.Name == paramName:
			paramValue = el.get_Parameter(param.GUID)
			return paramValue.AsString()
		elif param.Definition.Name == paramName: #handle project parameters?
			#paramValue = el.get_Parameter(paramName)
			return param.AsValueString()	

def MF_GetParameterByName(el, paramName):
	for param in el.Parameters:
		if param.IsShared and param.Definition.Name == paramName:
		#if param.Definition.Name == paramName:
			param = el.get_Parameter(param.GUID)
			return param
		elif param.Definition.Name == paramName: #handle project parameters?
			#paramValue = el.get_Parameter(paramName)
			return param			
			
	"""
	Filter 1	
	4289815	
	['Lighting Devices  ,  
	Fire Alarm Devices  ,  
	Data Devices  ,  
	
	Communication Devices  ,  
	Security Devices  ,  
	Nurse Call Devices  ,  
	Telephone Devices  ,  
	Air Terminals  , 
	Mechanical Equipment  ,  
	Lighting Fixtures  ,  
	Electrical Fixtures  ,  
	Electrical Equipment  , 
	Generic Models  ,  ', 
	['Classification.Uniclass.Ss.Number', ' - ', 'equals', '60', '60', 'Autodesk.Revit.DB.FilterStringRule', ' --- ']]
	"""


## START HERE #############

system_code = "Ss_60_40"

uniclass_groups = {
"Drainage Collection" :	"50_30",
"Gas Supply" :	"55_20",
"Fire Extinguishing Supply"	:"55_30",
"Water Supply":	"55_70",
"Space Heating & Cooling" :	"60_40",
"Ventilation" :	"65_40",
"Electrical Power Generation" :	"70_10",
"Electrical Distribution" :	"70_30",
"Lighting"	:"70_80",
"Communication"	:"75_10",
"Security":	"75_40",
"Safety and Protection" 	:"75_50",
"Control and Management":	"75_70",
"Protection" :	"75_80"
}



c = "Mechanical Equipment, Electrical Equipment"
f = system_code

collector = FilteredElementCollector(doc)

bindings = doc.ParameterBindings


catname = str(c[0])
bic = System.Enum.GetValues(BuiltInCategory) 
cats, bics = [], []
for i in bic:
    try:
        cat = Revit.Elements.Category.ById(ElementId(i).IntegerValue)
        cats.append(cat)
        bics.append(i)
    except:
        pass
 
for i, b in zip(cats, bics):
    if catname == str(i): 
        ost = b 

#print ost

mech_el= FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).FirstElement()
elec_el= FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).FirstElement()

print str(elec_el)

result_string = ""



# for p in el.Parameters:
	# if p.Definition in bindings:
		# result_string += "\n" + p.Definition.Name + ": Parameter Found!!! \n\n"
	# else:
		# result_string += "\n" + p.Definition.Name + ": Not Found...  \n\n"

		
# Classification Parameters apply to all objects, so we can get the GUID from any element with the following hack
		
p_m = MF_GetParameterByName(mech_el, "Classification.Uniclass.Ss.Number")
p_e = MF_GetParameterByName(elec_el, "Classification.Uniclass.Ss.Number")


			
			#print param.Name 

# UI.TaskDialog.Show(
		# "Result:", 
		# "Name: " + p_m.Definition.Name + "\n GUID: " + str(p_m.GUID ) + "\n\n" +
		
		# "Name: " + p_e.Definition.Name + "\n GUID: " + str(p_e.GUID )
		# )

# r = [
	# 'Family Name', 
	# 'SYMBOL_FAMILY_NAME_PARAM', 
	# 'contains', 
	# 'Jaga', 
	# 'Jaga', 
	# 'Autodesk.Revit.DB.FilterStringRule', 
	# ' --- '
	# ]



	
r = [
	'Family Name', 
	'-- ', 
	'contains', 
	'Jaga', 
	system_code, 
	'Autodesk.Revit.DB.FilterStringRule', 
	' --- '
	]	
	
inp = [f, c, r]	

from rpw import db

@db.Transaction.ensure('Create Filter from Excel')	
def mf_create_filter(input, param_name):	

	fName = input[0]
	categories = input[1]
	ruleData = input[2]
	
	#print ruleData
	

	log = []

	#categories = a[4]  # one long comma separated string from excel cell
	catList = map(str.strip, categories.split(',')) # list of string category names
	cats = []
	cList = []
	for c in catList:
		try:
		
			
			cat = doc.Settings.Categories.get_Item(c)
			cats.append(cat)  # Autodesk DB. Category
		except: pass	

	ruleStringList = list(group(ruleData, ' --- '))

	lst = ruleData
	w = ' --- '
	spl = [list(y) for x, y in itertools.groupby(lst, lambda z: z == w) if not x]

	ruleStringList = spl

	#allRules.append(ruleData)
	#allRules.append(ruleStringList)

	##updates.append([(vtId, fName, fId, categories, cats, ruleData) ])
	#############
	doit = 1

	typedCatList = ' - '

	try:
		for cat in cats:
			
			#cList.append(i)  # cannot cast System.String to Element Id
			#cList.append(ElementId(cat.Id))  # ERROR: Expected Built In Parameter - got 'ElementId'
			
			
			cList.append(cat.Id)
			
			typedCatList = List[ElementId](cList)    #unable to cast string to element id
			#typedCatList = cList  ## temporary debug
			#create rule list for filter
	except Exception as e:
		log.append("cat error: "+ str(e) )

	#typedCatList = cList  ## temporary debug


	## check if filter already exists
	existingFilters = FilteredElementCollector(doc).OfClass(ParameterFilterElement)
			
	existingFilterIds = []
	existingFilterNames = []
			
	for ef in existingFilters:
		existingFilterIds.append(ef.Id)  # this is a list of ElementIds
		existingFilterNames.append(ef.Name) 


			
	rules = []

	for r in ruleStringList:
		paramName = r[1]
		comparator = r[2]
		
		pValue = r[4]
		try:	
			bip = GetBuiltInParam(param_name)
			
			rule_param = ElementId(bip)
			
		except:
		
			mech_el= FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MechanicalEquipment).FirstElement()
			p_m = MF_GetParameterByName(mech_el, param_name)
			
			 # project parameter 
			rule_param = p_m.Id
		#string contains
		if comparator == "contains":
			rules.append(ParameterFilterRuleFactory.CreateContainsRule(rule_param, pValue, False) )
		# not contains
		if comparator == "does not contain":
			rules.append(ParameterFilterRuleFactory.CreateNotContainsRule(rule_param, pValue, False) )
		
			
		
	try:	
			
			## ClearRules for existing filter
			
			## Update with new rules ? 
		
			#fId = int(filterIDs[0])
		
		#t = Transaction(doc, __title__)
		
		#t.Start()
		match = "not matched"
		logRow = []
		# update existing filters
		
		
		# if ElementId(fId) in existingFilterIds: # pfe does not get defined?
			
			# match = "matched"			
			# pfe = doc.GetElement(ElementId(fId))
			
			# pFilter = pfe
			# # get previous name
			# previousName = pfe.Name
			# # update name
			# pfe.Name = fName
			# pfe.ClearRules()
			# pfe.SetRules(rules)
			# ## ParameterFilterElement.SetRules( rules)
			
			
			
			# ruleDataString ='(' + '),( '.join(ruleData) + ')'
			
			# logRow = [ (time), ("Filter " ),( previousName), ( " updated to "),(  pfe.Name), ("Rules:") ]
			# logRow.extend(ruleData) 
			
			# ####
			

			# ####
			
			
			# log.append(logRow) 
		
		# filterMatches.append(match)
		# #add new filters?
		# logRow = [ (time) ]
		
		
		#if fId == 0:
		try:
			## try creating filter
			
			## need to check if fName matches any existing filter name... 
			if fName not in existingFilterNames:
				pfe = ParameterFilterElement.Create(doc, fName, typedCatList, rules)
			
				pFilter = pfe			
			
				createResult = [("Filter Created :" ),( pfe.Name)]
			else:
				createResult = [("Filter name :" ),( fName), ("already in use")]
			
				i = existingFilterNames.index(fName)
				id = existingFilterIds[i]
				
				pfe = doc.GetElement(id)
				
				pFilter = pfe	
			
			logRow.extend(createResult)
			#view.AddFilter(pfe.Id)
		
		
		except Exception as e: 
			createResult = [( "Error Creating Filter : "),(  str(e) )]
			logRow.extend(createResult)
		pass
		log.append(logRow) 
		
		# Try Adding the filter pfe to the view - if it is already added it will throw an error
		# try:
			# viewFilters = view.GetFilters()
			
			# if pFilter.Id not in viewFilters:
			
				# view.AddFilter(pFilter.Id)
						
				# addResult = [(fName),("added to " ),( view.Name)]
			# else: 
				# addResult = [(fName),("already applied to " ),( view.Name)]
			
			# logRow.extend(addResult) 
		# except Exception as e: 
			# addResult = [ ("Error Adding Filter"),(fName),("to view"), (view.Name), (str(e)) ]
			# logRow.extend(addResult) 
		# pass
		
		# #set the overrides for the filter now that it is in in the View
		# try:
			# view.SetFilterOverrides(pFilter.Id, ogs)
			
			# view.SetFilterVisibility(pFilter.Id,visibility)
			# setResult = [(fName),("overrides updated in " ),( view.Name)]
			# logRow.extend(setResult)
		# except Exception as e: 
			# setResult = [ ("Error setting overrides for Filter"),(fName),("in view"), (view.Name), (str(e)) ]
			# logRow.extend(setResult) 
		# pass 	
		
		
		
		log.append(logRow) 
		
		
		#t.Commit()
	except Exception as e:
		log.append(str(e))	
		
	#cLists.append(cList)


	# Place your code below this line

	# Assign your output to the OUT variable.
	#print str( [updates, allRules, log, bicats, cLists, existingFilterIds, filterMatches, colours] )

	print log

uniclass_groups = {
"Drainage Collection":"50_30",
"Gas Supply":"55_20",
"Fire Extinguishing Supply":"55_30",
"Water Supply":"55_70",
"Space Heating & Cooling":"60_40",
"Ventilation":"65_40",
"Electrical Power Generation":"70_10",
"Electrical Distribution":"70_30",
"Lighting":"70_80",
"Communication":"75_10",
"Security":"75_40",
"Safety and Protection":"75_50",
"Control and Management":"75_70",
"Protection" :"75_80"
}



	
#for item in sorted(uniclass_groups.items(), key=lambda x:x[1] ):

categories = """
Mechanical Equipment, 
Electrical Equipment, 
Electrical Devices, 
Lighting Fixtures,
Lighting Devices, 
Nurse Call Devices, 
Security Devices, 
Data Devices, 
Plumbing Fixtures, 
Furniture
"""

f = system_code

	
inp = [f, c, r]		

for key, val in uniclass_groups.items():
	
	rule = [
	'blank', 
	'-- ', 
	'contains', 
	'Jaga', 
	val, 
	'blank', 
	' --- '
	]
	
	
	
	p_name = "Classification.Uniclass.Ss.Number"
	filter_name = "Uniclass - Ss_"+val + " - "  + key
	
	create_filter([filter_name, categories, rule], p_name)
	
	p_name = "Classification.Uniclass.EF.Number"
	filter_name = "Uniclass - EF_"+val + " - "  + key
	
	mf_create_filter([filter_name, categories, rule], p_name)



