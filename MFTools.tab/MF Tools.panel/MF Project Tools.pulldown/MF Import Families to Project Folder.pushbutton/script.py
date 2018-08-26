# -*- coding: utf-8 -*-
__title__ = 'MF Import Families to Project Folder'
__doc__ = """MF Import Families to Project Folder
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
import os
import shutil
import glob

def recursive_copy_files(source_path, destination_path, override=False):
	"""
	Recursive copies files from source  to destination directory.
	:param source_path: source directory
	:param destination_path: destination directory
	:param override if True all files will be overridden otherwise skip if file exist
	:return: count of copied files
	"""
	files_count = 0
	if not os.path.exists(destination_path):
		os.mkdir(destination_path)
	items = glob.glob(source_path + '/*')
	print "items:" + str(len(items))
	for item in items:
		if os.path.isdir(item):
			print "folder found"
			path = os.path.join(destination_path, item.split('/')[-1])
			files_count += recursive_copy_files(source_path=item, destination_path=path, override=override)
		else:
			file = os.path.join(destination_path, item.split('/')[-1])
			if not os.path.exists(file) or override:
				shutil.copyfile(item, file)
				files_count += 1
	return files_count			

	
	
def copytree(src, dst, symlinks = False, ignore = None):
  if not os.path.exists(dst):
    os.makedirs(dst)
    shutil.copystat(src, dst)
  lst = os.listdir(src)
  if ignore:
    excl = ignore(src, lst)
    lst = [x for x in lst if x not in excl]
  for item in lst:
    s = os.path.join(src, item)
    d = os.path.join(dst, item)
    if symlinks and os.path.islink(s):
      if os.path.lexists(d):
        os.remove(d)
      os.symlink(os.readlink(s), d)
      try:
        st = os.lstat(s)
        mode = stat.S_IMODE(st.st_mode)
        os.lchmod(d, mode)
      except:
        pass # lchmod not available
    elif os.path.isdir(s):
      copytree(s, d, symlinks, ignore)
    else:
      shutil.copy2(s, d)
	  
import shutil
 
def copyDirectory(src, dest):
    try:
        shutil.copytree(src, dest)
    # Directories are the same
    except shutil.Error as e:
        print('Directory not copied. Error: %s' % e)
    # Any error saying that the directory doesn't exist
    except OSError as e:
        print('Directory not copied. Error: %s' % e)	  

				
families = FilteredElementCollector(doc).OfClass(Family)

centralModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())



 
 
 
family_library_path = "Y:\Revit MEP\Revit Development\_Work in progress\_CARLOTTA WIP"


basePath = centralModelPath.split("\MF Model",1)[0] 

loadPath = basePath + "\Project Families"

target_path = basePath + "\Project Families\Imported"

#imported_files = copytree(family_library_path, target_path)

copyDirectory(family_library_path , target_path)

#print loadPath

print ( "imported to", target_path, "from", family_library_path )







