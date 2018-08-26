# -*- coding: utf-8 -*-
__title__ = 'MF Create Printing Sheets'
__doc__ = """MF Create Printing Sheets
"""

__helpurl__ = ""

import sys
## Add Path to MF Functions folder 	

sys.path.append(os.path.realpath(__file__).split("MFTools.extension")[0] + "MFTools.extension\MFTools.tab\MF_functions")

from MF_HeaderStuff import *


## START HERE #############

def _ask_for_titleblock(self):
	no_tb_option = 'No Title Block'
	titleblocks = DB.FilteredElementCollector(revit.doc)\
					.OfCategory(DB.BuiltInCategory.OST_TitleBlocks)\
					.WhereElementIsElementType()\
					.ToElements()

	tblock_dict = {'{}: {}'.format(tb.FamilyName,
								   revit.ElementWrapper(tb).name): tb
				   for tb in titleblocks}
	options = [no_tb_option]
	options.extend(tblock_dict.keys())
	selected_titleblocks = forms.SelectFromList.show(options,
													 multiselect=False)
	if selected_titleblocks:
		if no_tb_option not in selected_titleblocks:
			self._titleblock_id = tblock_dict[selected_titleblocks[0]].Id
		else:
			self._titleblock_id = DB.ElementId.InvalidElementId
		return True

	return False

@staticmethod
def _create_placeholder(sheet_num, sheet_name):
	with DB.Transaction(revit.doc, 'Create Placeholder') as t:
		try:
			t.Start()
			new_phsheet = DB.ViewSheet.CreatePlaceholder(revit.doc)
			new_phsheet.Name = sheet_name
			new_phsheet.SheetNumber = sheet_num
			t.Commit()
		except Exception as create_err:
			t.RollBack()
			logger.error('Error creating placeholder sheet {}:{} | {}'
						 .format(sheet_num, sheet_name, create_err))

def _create_sheet(self, sheet_num, sheet_name):
	with DB.Transaction(revit.doc, 'Create Sheet') as t:
		try:
			t.Start()
			new_phsheet = DB.ViewSheet.Create(revit.doc,
											  self._titleblock_id)
			new_phsheet.Name = sheet_name
			new_phsheet.SheetNumber = sheet_num
			t.Commit()
		except Exception as create_err:
			t.RollBack()
			logger.error('Error creating sheet sheet {}:{} | {}'
						 .format(sheet_num, sheet_name, create_err))

def create_sheets(self, sender, args):
	self.Close()

	if self._process_sheet_code():
		if self.sheet_cb.IsChecked:
			create_func = self._create_sheet
			transaction_msg = 'Batch Create Sheets'
			if not self._ask_for_titleblock():
				script.exit()
		else:
			create_func = self._create_placeholder
			transaction_msg = 'Batch Create Placeholders'

		with DB.TransactionGroup(revit.doc, transaction_msg) as tg:
			tg.Start()
			for sheet_num, sheet_name in self._sheet_dict.items():
				logger.debug('Creating Sheet: {}:{}'.format(sheet_num,
															sheet_name))
				create_func(sheet_num, sheet_name)
			tg.Assimilate()
	else:
		logger.error('Aborted with errors.')

############################################################

#Collect View Templates from Project

viewTemplates = []
collector = FilteredElementCollector(doc).OfClass(View)
for i in collector:
	if i.IsTemplate == True:
		viewTemp = i
		viewTemplates.append(i)
		
		
class ViewOption(BaseCheckBoxItem):
    def __init__(self, view_element):
        super(ViewOption, self).__init__(view_element)

    @property
    def name(self):
		
        
		return '{} ({}) '.format(self.item.ViewName, self.item.ViewType)

		

# #select multiple
# class LevelOption(BaseCheckBoxItem):
    # def __init__(self, level_element):
        # super(LevelOption, self).__init__(level_element)

    # @property
    # def name(self):
		
        
		# return '{} ({}) '.format(self.item.Name)


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


# filter view templates for 500 - Printing Views only


printingTemplates = [x for x in viewTemplates if "Pr" in x.Name ]
## Drawing / Working templates
#printingTemplates = [x for x in viewTemplates if "100 - Wo" in x.Name ]

selected = []
return_options = SelectFromCheckBoxes.show(
			sorted([ViewOption(x) for x in printingTemplates],
				   key=lambda x: x.name),
			title="Select Foreground Views to Create",
			button_name="Create Foreground Views for Selected Templates for Each Level",
			width=800)
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]

foregroundViewTemplates = list(selected)

# selected = []

# bgTemplates = [x for x in viewTemplates if "Bg" in x.Name ]

# return_options = SelectFromCheckBoxes.show(
			# sorted([ViewOption(x) for x in bgTemplates],
				   # key=lambda x: x.name),
			# title="Select Background Views to Create",
			# button_name="Create Background Views for Selected Templates for Each Level",
			# width=800)
# if return_options:
		# selected = [x.unwrap() for x in return_options if x.state]
		# backgroundViewTemplates = list(selected)
# else:		
		# backgroundViewTemplates = []
# ### cahnge this to single selection


fixTemplates = [x for x in viewTemplates if "Fx" in x.Name ]

selected = []
return_options = SelectFromCheckBoxes.show(
			sorted([ViewOption(x) for x in fixTemplates],
				   key=lambda x: x.name),
			title="Select Fix Template",
			button_name="Apply Fix Template to New Views",
			width=800)
if return_options:
		selected = [x.unwrap() for x in return_options if x.state]
		fixTemplate = list(selected)
else: 
		fixTemplate = []	




titleblocks = FilteredElementCollector(revit.doc)\
					.OfCategory(DB.BuiltInCategory.OST_TitleBlocks)\
					.WhereElementIsElementType()\
					.ToElements()
					
no_tb_option = 'No Title Block'
	

tblock_dict = {'{}: {}'.format(tb.FamilyName,
								   revit.ElementWrapper(tb).name): tb
				   for tb in titleblocks}
options = [no_tb_option]
options.extend(tblock_dict.keys())
selected_titleblocks = forms.SelectFromList.show(options,
													 multiselect=False)		


titleblock_id = 	tblock_dict[selected_titleblocks[0]].Id											 

									 



viewFamilyTypes = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()



## choose template sheet to get viewport positions from

collector = FilteredElementCollector(doc).OfClass(ViewSheet)

sheets = [x for x in collector  if x.GetAllPlacedViews() ]  ## only show sheets with a placed view

options = []
sheet_dict = {'{}: {}'.format(s.Name,
								   revit.ElementWrapper(s).name): s
				   for s in sheets}
options.extend(sheet_dict.keys())
selected_sheet = forms.SelectFromList.show(options,
												title="Choose Template Sheet to get Viewport Location",
												button_name="Select",
												width=800,
													 multiselect=False)	
													 
template_sheetId = sheet_dict[selected_sheet[0]].Id		

templateSheet = sheet_dict[selected_sheet[0]]

## choose Viewport to get positions from



viewports = templateSheet.GetAllViewports()



options = []
vp_dict = {'{}: {}'.format(doc.GetElement(vp).Name,
								   doc.GetElement(doc.GetElement(vp).ViewId).Name): vp
				   for vp in viewports}
options.extend(vp_dict.keys())
selected_vp = forms.SelectFromList.show(options,
										title="Choose Viewport match Location",
												button_name="Select",
												width=800,
													 multiselect=False)	
													 
template_vpId = vp_dict[selected_vp[0]]
template_vp = doc.GetElement(vp_dict[selected_vp[0]])

template_vp_location = 	template_vp.GetBoxCenter()


	
# create Floor Plan Views for each Level and View Template
name = "Floor Plan"
id = -1
for viewType in viewFamilyTypes:
	typeName = viewType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
	if typeName == name:
		viewTypeId = viewType.Id
		break


		
t = Transaction(doc)
t.Start(__title__)		

for level in levels:		

	for fvt in foregroundViewTemplates:
		try:
			#Create Foreground Floor Plan
			fv = ViewPlan.Create(doc, viewTypeId, level.Id) 
			
			#Set Name of new Floor Plan - based on level name and View Template name
			fv.Name = fvt.Name + " - " + level.Name + " - Print"
			
					
			#Apply View Template  
			fv.ViewTemplateId = fvt.Id
			
			#Apply Links Fix Template ## TEMP HACK
			
			if fixTemplate[0] is not None:
				fv.ApplyViewTemplateParameters(fixTemplate[0])
		
		except Exception as e:
			print(str(e))
			pass
	# ## should only allow one choice
	# #for bvt in backgroundViewTemplates:
		# if backgroundViewTemplates[0] is not None:
			# bvt = backgroundViewTemplates[0]
		
			# try:
				
				
				# #Create Background Floor Plan
				# bv = ViewPlan.Create(doc, viewTypeId, level.Id) 
				
				# #Set Name of new Floor Plan - based on level name and View Template name
				# bv.Name = fvt.Name + " - " + level.Name + " - Background"
				
				# #Apply View Template  
				# bv.ViewTemplateId = bvt.Id
				
				# #Apply Links Fix Template ## TEMP HACK
				
				# if fixTemplate[0] is not None:
					# bv.ApplyViewTemplateParameters(fixTemplate[0])
			
			# except Exception as e:
				# print(str(e))
				# pass		

		## create sheets

		new_sheet = ViewSheet.Create(revit.doc,
											  titleblock_id)
											  
		new_sheet.Name = "Printing Sheet - " + 	fvt.Name + " - " + level.Name	
		
		## create sensible sheet number here... 
		
		vp_location = template_vp_location
		
		
		fvp = Viewport.Create(doc, new_sheet.Id,fv.Id,  vp_location )
		


t.Commit()		


