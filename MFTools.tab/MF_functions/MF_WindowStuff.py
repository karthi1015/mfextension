import clr
import os
import os.path as op
import pickle as pl

import sys
import subprocess
import time

import struct

import rpw
from rpw import doc, uidoc, DB, UI

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

import System

from System import Array

clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

try:
    from pyrevit.versionmgr import PYREVIT_VERSION
except:
    from pyrevit import versionmgr
    PYREVIT_VERSION = versionmgr.get_pyrevit_version()

pyRevitNewer44 = PYREVIT_VERSION.major >= 4 and PYREVIT_VERSION.minor >= 5

if pyRevitNewer44:
    from pyrevit import script, revit, forms
    from pyrevit.forms import *
    output = script.get_output()
    logger = script.get_logger()
    linkify = output.linkify
    from pyrevit.revit import doc, uidoc, selection
    selection = selection.get_selection()

else:
    from scriptutils import logger
    from scriptutils.userinput import SelectFromList, SelectFromCheckBoxes
    from revitutils import doc, uidoc, selection
	


from System.Windows.Forms import (
    Form, Panel, Label,
    TextBox, DockStyle, Button,
    ScrollBars, Application, DataGridView,
    DataGridViewColumnHeadersHeightSizeMode,
    MessageBox, MessageBoxButtons,
    MessageBoxIcon
)

from Autodesk.Revit.UI import TaskDialog

from System.Windows.Forms import *

from System.Drawing import (
    Point, Size,
    Font, FontStyle,
    GraphicsUnit
)

clr.AddReference("System.Data")

from System.Data import DataSet
from System.Data.Odbc import OdbcConnection, OdbcDataAdapter

msgBox = TaskDialog

headers = ["Title ", "Title ", "Title ", "Title ", "Title "]
signals = ["Data ", "Data ", "Data ", "Data ", "Data "]

array_str = Array.CreateInstance(str, len(headers))


class DataGridViewQueryForm(Form):

        def __init__(self):
            self.Text = 'Signals'
            self.ClientSize = Size(942, 255)
            self.MinimumSize = Size(500, 200)

            self.setupDataGridView()


        def setupDataGridView(self):            
            self._dataGridView1 = DataGridView()
            self._dataGridView1.AllowUserToOrderColumns = True
            self._dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            self._dataGridView1.Dock = DockStyle.Fill
            self._dataGridView1.Location = Point(0, 111)
            self._dataGridView1.Size = Size(506, 273)
            self._dataGridView1.TabIndex = 3
            self._dataGridView1.ColumnCount = len(headers)
            self._dataGridView1.ColumnHeadersVisible = True
            for i in range(len(headers)):
                self._dataGridView1.Columns[i].Name = headers[i]

            for j in range(len(signals)):

				for k in range(len(headers)):
					array_str[k] = signals[j][k]
				self._dataGridView1.Rows.Add(array_str)

            self.Controls.Add(self._dataGridView1)





#Application.Run(DataGridViewQueryForm())