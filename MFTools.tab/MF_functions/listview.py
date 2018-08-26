import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import Application, Form, StatusBar
from System.Windows.Forms import ListView, View, ColumnHeader, ComboBox
from System.Windows.Forms import ListViewItem, DockStyle, SortOrder
from System.Drawing import Size



from os import listdir
from os.path import isfile, join

mypath = "Y:\Revit MEP\Revit Development\_Work in progress\_CARLOTTA WIP\MXF Families (COBIe)"

global familyfiles

familyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]


class IForm(Form):

    def __init__(self):
		self.Text = 'ListBox'
		name = ColumnHeader()
		name.Text = 'Name'
		name.Width = -1
		year = ColumnHeader()
		year.Text = 'Type'
		year.Width = -1
		
		cb = ComboBox()
		#cb.Parent = self
		#cb.Location = Point(50, 30)

		cb.Items.AddRange(("Ubuntu",
			"Mandriva",
			"Red Hat",
			"Fedora",
			"Gentoo"))

		cb.SelectionChangeCommitted += self.OnChanged

		self.SuspendLayout()

		lv = ListView()
		lv.Parent = self
		lv.FullRowSelect = True
		lv.GridLines = True
		lv.AllowColumnReorder = True
		lv.Sorting = SortOrder.Ascending
		lv.Columns.AddRange((name, year))
		lv.ColumnClick += self.OnColumnClick

		# for act in actresses.keys():
			# item = ListViewItem()
			# item.Text = act
			# item.SubItems.Add(str(actresses[act]))
			# lv.Items.Add(item)
		for f in familyfiles:
			item = ListViewItem()
			item.Text = f
			item.SubItems.Add("Family File")
			lv.Items.Add(item)
			
		lv.Dock = DockStyle.Fill
		lv.Click += self.OnChanged

		self.sb = StatusBar()
		self.sb.Parent = self
		lv.View = View.Details

		self.ResumeLayout()

		self.Size = Size(850, 800)
		self.CenterToScreen()
    

    def OnChanged(self, sender, event):

        name = sender.SelectedItems[0].SubItems[0].Text
        born = sender.SelectedItems[0].SubItems[1].Text
        self.sb.Text = name + ', ' + born
    

    def OnColumnClick(self, sender, event):

        if sender.Sorting == SortOrder.Ascending:
            sender.Sorting = SortOrder.Descending
        else: 
            sender.Sorting = SortOrder.Ascending


Application.Run(IForm())