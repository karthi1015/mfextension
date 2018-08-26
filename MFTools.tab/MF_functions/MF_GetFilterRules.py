
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

def MF_GetFilterRules(f):
	types = []
	pfRuleList = []
	categories  = ''
	for c in f.GetCategories():
			categories += Category.GetCategory(doc, c).Name + "  ,  "
	
	
	for rule in f.GetRules():
				
				
				try:
					comparator = ""
					ruleValue = ""
					ruleInfo = ""
					
					if rule.GetType() == FilterDoubleRule:
						ruleValue = "filterdoublerule"
						
						fdr = rule
						
						if (fdr.GetEvaluator().GetType() == FilterNumericLess):
							comparator = "<"
						elif (fdr.GetEvaluator().GetType() == FilterNumericGreater):
							comparator = ">"
							
						ruleValue = fdr.RuleValue.ToString()
						
						ruleName = ruleValue
						
					if rule.GetType() == FilterStringRule:
						ruleValue = "filterstringrule"	
						
						fsr = rule 
							
						if (fsr.GetEvaluator().GetType() == FilterStringBeginsWith):
							comparator = "starts with"
						elif (fsr.GetEvaluator().GetType() == FilterStringEndsWith):
							comparator = "ends with"
						elif (fsr.GetEvaluator().GetType() == FilterStringEquals):
							comparator = "equals"
						elif (fsr.GetEvaluator().GetType() == FilterStringContains):
							comparator = "contains"
							
						ruleValue = fsr.RuleString
						
						ruleName = ruleValue
					
					if rule.GetType() ==FilterInverseRule:	
					# handle 'string does not contain '
						
						fInvr = rule.GetInnerRule()
						
						if (fsr.GetEvaluator().GetType() == FilterStringBeginsWith):
							comparator = "does not start with"
						elif (fsr.GetEvaluator().GetType() == FilterStringEndsWith):
							comparator = "does not end with"
						elif (fsr.GetEvaluator().GetType() == FilterStringEquals):
							comparator = "does not equal"
						elif (fsr.GetEvaluator().GetType() == FilterStringContains):
							comparator = "does not contain"
						
						ruleValue = fInvr.RuleString
						
						ruleName = ruleValue
						
					if rule.GetType() == FilterIntegerRule:
						
						#comparator = "equals" 
						ruleValue = "filterintegerrule"	
						
						fir = rule
						
						if (fir.GetEvaluator().GetType() == FilterNumericEquals):
							comparator = "="
						elif (fir.GetEvaluator().GetType() == FilterNumericGreater):
							comparator = ">" 
						elif (fir.GetEvaluator().GetType() == FilterNumericLess):
							comparator = "<"  				
						  
						ruleValue =  fir.RuleValue
						  
						
						
					
					if rule.GetType() ==FilterElementIdRule:
						
						comparator = "equals" 
						feidr = rule
						
						ruleValue = doc.GetElement(feidr.RuleValue)
						
						ruleName = ruleValue.Abbreviation
						
						t = ruleValue.GetType()
						#ruleName = doc.GetElement(ruleValue.Id).Name
						types.append(t)
						
					
					
					
					
					paramName = ""
					bipName = " - "
					if (ParameterFilterElement.GetRuleParameter(rule).IntegerValue < 0):
						
						bpid = f.GetRuleParameter(rule).IntegerValue
						
						#bp = System.Enum.Parse(clr.GetClrType(ParameterType), str(bpid) )
						
						#paramName = LabelUtils.GetLabelFor.Overloads[BuiltInParameter](bpid)
						#paramName = doc.get_Parameter(ElementId(bpid)).ToString()
						paramName = Enum.Parse( clr.GetClrType( BuiltInParameter ), str(bpid) )
						param = Enum.Parse( clr.GetClrType( BuiltInParameter ), str(bpid) )
						bipName = param.ToString()
						
						#YESSS
						paramName = LabelUtils.GetLabelFor.Overloads[BuiltInParameter](param)
					
					else:
						paramName = doc.GetElement(ParameterFilterElement.GetRuleParameter(rule)).Name
					#paramName = doc.GetElement(ParameterFilterElement.GetRuleParameter(rule)).Name
						
				
					
				
					#ruleData += "'" + paramName + "' " + comparator + " " + "'" + ruleValue.ToString() + "'" + Environment.NewLine;	
					#ruleInfo += "'" + str(paramName) + "' " + comparator + " " + "'" + ruleValue.ToString() + "---"
					try:
						ruleInfo += "" + str(paramName) + " - " + bipName + " -  " + comparator + " - " +  ruleValue.ToString() + " - " + ruleName + " - " + rule.GetType().ToString() + "   ---   "
						
						pfRuleList += [  (str(paramName)), ( bipName  ), (comparator), (ruleValue.ToString()) ,(ruleName) , (rule.GetType().ToString() ), (" --- " )]
						
					except:
						ruleInfo += "" + str(paramName) + " - " + bipName + " - " + comparator + " -  " +  ruleValue.ToString() + " -  " + ruleName.ToString() + " - " + rule.GetType().ToString() + "   ---   "		
						
						pfRuleList += [  (str(paramName)) , ( bipName ), (comparator), (ruleValue.ToString()) ,( ruleName.ToString() ) , (rule.GetType().ToString() ), (" --- " )]
				
				#ruleList.append([str(paramName), comparator, ruleName])
				#ruleInfo = ( (str(paramName) , comparator , ruleValue.ToString() ) )
				#filterRuleList.append(pfRuleList) 
				
				#ruleValues.append(ruleValue)
				
				except Exception as e:
					print str(e)
			
				#sublist.extend(pfRuleList)
			
				return [categories, pfRuleList]