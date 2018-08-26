# cellfmt.py

"""
Test of IronPython's ability to change
individual cell formats in a dot Net
DataGridView control.
Intended behavior of DataGridViewControl:
    1) highlights row and column of mouse
       position simultaneously
    2) mock selection in color scheme
       other than Windows default
    3) selection of individual cells
       and individual rows allowed
       (no multiple selection)
    4) auto copy selections to clipboard
       in a manner that allows direct
       pasting into Excel
Geared toward use with ~1000 rows and
~10 columns.
Intended to make tracking of active row
and data easier.
"""
# XXX - row/column tracking is slow with
# XXX       more than 500 rows
# XXX   could try to optimize by tracking cell indices

# XXX - PageDown, PagUp, Home, and Arrow Keys will
# XXX       still highlight and select cells in the 
# XXX          Windows default fashion.

# XXX - could write more code for column header width
# XXX       adjustment and row width adjustment
# XXX       - columnwidth adjustment clears mock selection
# XXX       - rowwidth adjustment selects row above row
# XXX             boundary

import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import Form
from System.Windows.Forms import DataGridView
from System.Windows.Forms import DataGridViewContentAlignment
from System.Windows.Forms import Application
from System.Windows.Forms import Control
from System.Windows.Forms import Clipboard
from System.Windows.Forms import DataFormats
from System.Windows.Forms import DataObject

from System.Drawing import Point
from System.Drawing import Size
from System.Drawing import Font
from System.Drawing import FontStyle
from System.Drawing import Color

from System import Text

from System.IO import MemoryStream

# formatting constants
MDDLCNTR = DataGridViewContentAlignment.MiddleCenter
BOLD = Font(Control.DefaultFont, FontStyle.Bold)
REGL = Font(Control.DefaultFont, FontStyle.Regular)
SELECTCOLOR = Color.LightSkyBlue
ROWCOLOR = Color.Yellow
COLUMNCOLOR = Color.Cyan
MOUSEOVERCOLOR = Color.GreenYellow
REGULARCOLOR = Color.White
ROWHDRWDTH = 65

CSV = DataFormats.CommaSeparatedValue

# hack for identifying mouseovered cell
#    mousing over one of the header cells yields
#        an event.RowIndex value of -1       
INDEXERROR = -1

NUMCOLS = 3
HEADERS = ['positive', 'negative', 'flat']
TESTDATA = (['happy', 'sad', 'indifferent'],
            ['ebullient', 'despondent', 'phlegmatic'],
            ['elated', 'depressed', 'apathetic'],
            ['fired up', 'bummed out', "doesn't care"],
            ['psyched', 'uninspired', 'blah'])
NUMROWS = len(TESTDATA)

def getcellidxs(event):
    """
    From a mouse event on the DataGridView,
    returns the row and column indices of 
    the cell as a 2 tuple of integers.
    """
    # this is redundant with DataGridForm.getcell
    #    class methods were not handling unpacking
    #        of tuple with cell in it
    #    trying a separate function
    print event.RowIndex, event.ColumnIndex
    return event.RowIndex, event.ColumnIndex

def resetcellfmts(gridcontrol, numrows):
    """
    Initialize formatting of all data
    cells in the grid control.
    """
    # need to cycle through all cells individually
    #    to reset them
    for num in xrange(numrows):
        row = gridcontrol.Rows[num]
        for cell in row.Cells:
            # skip over selected cell(s)
            if cell.Style.BackColor != SELECTCOLOR:
                cell.Style.Font = REGL
                cell.Style.BackColor = REGULARCOLOR 

def resetheaderfmts(gridcontrol, rowidx, colidx):
    """
    Reset BackColor on "Header" cells for 
    rows and columns.
    """
    col = gridcontrol.Columns[colidx]
    col.HeaderCell.Style.BackColor = Color.Empty
    # for row header formats, don't clear selected
    row = gridcontrol.Rows[rowidx]
    if row.HeaderCell.Style.BackColor != SELECTCOLOR:
        row.HeaderCell.Style.BackColor = Color.Empty

def clearselection(gridcontrol):
    """
    This works as a separate function,
    but not as part of the form class.
    Clears gridview selection.
    """
    # clear selection
    gridcontrol.ClearSelection()

def mockclearselection(gridcontrol):
    """
    Clears mock selection on custom 
    color scheme for DataGridView.
    """
    # have to cycle through all cells
    rows = gridcontrol.Rows
    for row in rows:
        # deal with selected header, if any
        row.HeaderCell.Style.BackColor = Color.Empty
        cells = row.Cells
        for cell in cells:
            if cell.Style.BackColor == SELECTCOLOR:
                cell.Style.BackColor = REGULARCOLOR
                cell.Style.Font = REGL

def copytoclipboard(args):
    """
    Put data on Windows clipboard 
    in csv format.
    """
    csvx = ""
    if len(args) == 0:
        # clear clipboard
        print 'clearing clipboard'
    elif len(args) == 1:
        csvx = args[0]
    else:
        csvx = ""
        csvx = ','.join(args)
    dobj = DataObject()
    # hack from MSDN PostID 238181
    # this is a bit bizarre,
    #    but it works for getting csv data into Excel
    txt = Text.Encoding.Default.GetBytes(csvx)
    memstr = MemoryStream(txt)
    dobj.SetData('CSV', memstr)
    dobj.SetData(CSV, memstr)
    Clipboard.SetDataObject(dobj)

class DataGridForm(Form):
    """
    Container for the DataGridView control
    I'm trying to test.
    DataGridView is customized to have row
    and column of mouse-over'd cell highlighted.
    Also, there is a customized selection 
    color and selection limitations (one cell
    or one row at a time).
    """
    def __init__(self, numcols, numrows):
        """
        numcols is the number of columns
        in the grid.
        """
        self.Text = 'DataGridView Cell Format Test'
        self.ClientSize = Size(400, 175)
        self.MinimumSize = Size(400, 175)
        self.dgv = DataGridView()
        self.numcols = numcols
        self.numrows = numrows
        self.setupdatagridview()
        self.adddata()
        self.formatheaders()
        # clears Windows default selection on load
        clearselection(self.dgv)
    def setupdatagridview(self):
        """
        General construction of DataGridView control.
        Bind mouse events as appropriate.
        """
        self.dgv.Location = Point(0, 0)
        self.dgv.Size = Size(375, 150)
        self.Controls.Add(self.dgv)
        # have to have columns defined before inserting rows
        self.dgv.ColumnCount = self.numcols
        # center all text in all data cells by default
        self.dgv.DefaultCellStyle.Alignment = MDDLCNTR
        # use Mouse events for contingency that actual
        #    position is required
        #    otherwise, can use events without "Mouse"
        #       in them
        # CellMouseEnter event for formatting
        self.dgv.CellMouseEnter += self.onmouseovercell
        # CellMouseLeave event for formatting
        self.dgv.CellMouseLeave += self.onmouseleavingcell
        # another try at MouseClick (avoiding default select color)
        self.dgv.CellMouseUp += self.onmouseclickcell
        # add empty rows first
        for num in xrange(self.numrows):
            self.dgv.Rows.Add()
        # format empty cells
        resetcellfmts(self.dgv, self.numrows)
        # lock control so user cannot do anything to it datawise
        self.dgv.AllowUserToAddRows = False
        self.dgv.AllowUserToDeleteRows = False
        self.dgv.ReadOnly = True
        self.dgv.ClearSelection()
    def formatheaders(self):
        """
        Get header row and left side column 
        populated and formatted.
        """
        # give names to columns
        # if the name is the same as the desired header caption,
        #     the Name attribute will take care of the caption
        for num in xrange(self.numcols):
            # need to center text in header row 
            #    separate from data rows
            col = self.dgv.Columns[num]
            col.Name = HEADERS[num]
            # slightly left of center on headers
            col.HeaderCell.Style.Alignment = MDDLCNTR
            # sets font and font style
            col.HeaderCell.Style.Font = BOLD 
            col.HeaderCell.Style.ForeColor = Color.MidnightBlue
        # put numbers on rows
        for num in xrange(self.numrows):
            row = self.dgv.Rows[num]
            # get sequential numeric label on side of row
            row.HeaderCell.Value = str(num + 1)
            # sets font and font style
            row.HeaderCell.Style.Font = BOLD 
            row.HeaderCell.Style.ForeColor = Color.Blue
        # XXX - clear button is implicit, not explicit
        self.dgv.TopLeftHeaderCell.Value = 'CLEAR'
        self.dgv.TopLeftHeaderCell.Style.Font = BOLD
        self.dgv.TopLeftHeaderCell.Style.ForeColor = Color.Blue
        self.dgv.RowHeadersWidth = ROWHDRWDTH
    def adddata(self):
        """
        Put data into the grid, 
        row by row, column by column.
        """
        # go off indices of rows for placing data in cells
        for num in xrange(self.numrows):
            row = self.dgv.Rows[num]
            # iterator for data - places in correct column
            dat = (datax for datax in TESTDATA[num]) 
            for cell in row.Cells:
                cell.Value = dat.next()
    def getcell(self, event):
        """
        Gets DataGridViewCell that is responding
        to an event.
        Attempt to minimize code duplication by 
        applying method to multiple events.
        """
        colidx = event.ColumnIndex
        rowidx = event.RowIndex
        if rowidx > INDEXERROR and colidx > INDEXERROR:
            # to get a specific cell, need to first
            #    get row the cell is in,
            #        then get cell indexed by column
            row = self.dgv.Rows[rowidx]
            cell = row.Cells[colidx]
            return cell
        else:
            return None
    def onmouseovercell(self, sender, event):
        """
        Change format of data cells
        when mouse passes over them.
        """
        # had to separate these into two functions
        #    problems with tuple unpacking
        cell = self.getcell(event)
        rowidx, colidx = getcellidxs(event)
        if cell:
            # 1) take care of row and column header formatting
            col = self.dgv.Columns[colidx]
            col.HeaderCell.Style.BackColor = COLUMNCOLOR
            row = self.dgv.Rows[rowidx]
            # only change if the row header is not selected
            if row.HeaderCell.Style.BackColor != SELECTCOLOR:
                row.HeaderCell.Style.BackColor = ROWCOLOR
            # 2) bold individual cell
            cell.Style.Font = BOLD 
            # 3) color individual cell green
            # but skip if it is a selected cell
            if cell.Style.BackColor != SELECTCOLOR:
                cell.Style.BackColor = MOUSEOVERCOLOR
            row = self.dgv.Rows[rowidx]
            for cellx in row.Cells:
                # 4) color each cell in row except 
                #    green cell yellow
                if cellx.ColumnIndex != colidx:
                    # skip selected cell
                    if cellx.Style.BackColor != SELECTCOLOR:
                        cellx.Style.BackColor = ROWCOLOR
                        # 5) add bold to row
                        cellx.Style.Font = BOLD
            # highlighting a column is harder
            #    have to cycle through all cells
            for num in xrange(self.numrows):
                for num2 in xrange(self.numcols):
                    # want to skip single highlighted cell
                    if num != rowidx:
                        if num2 == colidx:
                            row = self.dgv.Rows[num]
                            cell = row.Cells[num2]
                            # 6) color all other cells in column cyan
                            # skip selected cells
                            if cell.Style.BackColor != SELECTCOLOR:
                                cell.Style.BackColor = COLUMNCOLOR
    def onmouseleavingcell(self, sender, event):
        """
        Change format of data cells
        back to "normal" when mouse passes 
        out of the cell.
        """
        cell = self.getcell(event)
        rowidx, colidx = getcellidxs(event)
        if cell:
            # need to cycle through all cells individually
            #    to reset them
            resetcellfmts(self.dgv, self.numrows)
            resetheaderfmts(self.dgv, rowidx, colidx)
            clearselection(self.dgv)
    def onmouseclickcell(self, sender, event):
        """
        Attempt to override selection.
        """
        # get selected cells' data onto clipboard
        selected = []
        mockclearselection(self.dgv)
        # had to separate these into two functions
        #    problems with tuple unpacking
        cell = self.getcell(event)
        rowidx, colidx = getcellidxs(event)
        # overrides Windows selection color
        #   sometimes flashes blue for a split second
        clearselection(self.dgv)
        # if dealing with one valid data cell (not header)
        if cell:
            cell.Style.Font = BOLD 
            cell.Style.BackColor = SELECTCOLOR
            selected.append(cell.Value)
        # if a row header is clicked, select row
        if colidx == INDEXERROR:
            # need to make sure that upper left header is not clicked
            if rowidx != INDEXERROR:
                row = self.dgv.Rows[rowidx]
                cells = row.Cells
                # highlight all the cells in the row
                for cell in cells:
                    cell.Style.Font = BOLD
                    cell.Style.BackColor = SELECTCOLOR
                    selected.append(cell.Value)
                # highlight the row header
                row.HeaderCell.Style.BackColor = SELECTCOLOR
        # get mouseover coloration reset
        self.onmouseovercell(sender, event)
        copytoclipboard(selected)
            
DGF = DataGridForm(NUMCOLS, NUMROWS)
Application.Run(DGF)