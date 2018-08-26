import clr
clr.AddReference("System.Xml")
clr.AddReference("PresentationFramework, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
clr.AddReference("PresentationCore, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
from System.IO import StringReader
from System.Xml import XmlReader
from System.Windows.Markup import XamlReader, XamlWriter
from System.Windows import Window, Application
from System.Windows.Controls import Button, Canvas

xaml = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Enterprisey-FireDrop Express" Width="640" Height="480"
    >

<DockPanel>
       <StackPanel DockPanel.Dock ="Left">
         <StackPanel.Margin>
    <Thickness Left="10" Top="10" Right="10" Bottom="10"/>
  </StackPanel.Margin>
        <StackPanel.Background>
            <RadialGradientBrush 
                RadiusX="{Binding ElementName=RadiusX, Path=Value}" 
                RadiusY="{Binding ElementName=RadiusY, Path=Value}"> 
                    <GradientStop Color="#FAFAFAFA" Offset="0"/>
                    <GradientStop Color="#B2171F6E" Offset="0.971"/>
            </RadialGradientBrush>
        </StackPanel.Background>
        
        <Expander IsExpanded ="True" Header ="Site">
        <StackPanel x:Name="N">
        <Button Content="Open">
        
        </Button>
        <Button Content="New" x:Name="NewSite">
        </Button>
        <Button Content="Build">
        </Button>
        <Button Content="Publish">
        </Button>
        <Button Content="ViewOnline">
        </Button>
        
        </StackPanel>
        </Expander>
        <Expander IsExpanded ="True" Header ="Entry">
        <StackPanel>
        <Button Content="Open">
        
        </Button>
        <Button Content="New">
        </Button>
        <Button Content="Save">
        </Button>
        <Button Content="Preview">
        </Button>
        
        </StackPanel>
        </Expander>
        <Expander IsExpanded ="True" Header ="Plugins">
        <StackPanel>
        <Button Content="ScratchPad" x:Name="ScratchPad">
        </Button>
        <Button Content="Photo" x:Name="Photo">
        </Button>
        <Button Content="GR Shared Items" x:Name="Shared">
        </Button>
        
        </StackPanel>
        </Expander>
       </StackPanel>

    <Grid>
        <ListBox Margin="3,3,0,3" Padding="2" HorizontalAlignment="Left" Width="90">
            <ListBoxItem>Document 1</ListBoxItem>
            <ListBoxItem>Document 2</ListBoxItem>
            <ListBoxItem>Document 3</ListBoxItem>
        </ListBox>
        <RichTextBox Padding="2" Background="Thistle" Margin="97,3,3,3" />
    </Grid>

</DockPanel>

</Window>"""


#
# Waddle returns a dictionary of Control types e.g. listbox, Button.
# Each Entry is a dictionary of Control Instance Names i.e.
# controls['Button']['NewSite'] returns the button control named NewSite
# Controls should have Unique names and only those with a Name attrib set
# will be included.
#
def Waddle(c, d):
    s = str(c.__class__)
    if "System.Windows.Controls." in str(c) and hasattr(c,"Name") and c.Name.Length>0:
        ControlType = s[s.find("'")+1:s.rfind("'")]
        if ControlType not in d:
            d[ControlType] = {}
        d[ControlType][c.Name] = c
    if hasattr(c,"Children"):
        for cc in c.Children:
            Waddle(cc, d)
    elif hasattr(c,"Child"):
        Waddle(c.Child, d)
    elif hasattr(c,"Content"):
        Waddle(c.Content, d)

# Test Functions.
def sayhello(s,e):
    print "sayhello"
def sayhello2(s,e):
    print "sayhello2"

if __name__ == "__main__":
    xr = XmlReader.Create(StringReader(xaml))
    win = XamlReader.Load(xr)
    
    controls = {}
    Waddle(win, controls)
    
    #Make all Named buttons do something!
    for butt in controls['Button']:
        controls['Button'][butt].Click += sayhello
    
    #Make one button do something.
    controls['Button']['NewSite'].Click += sayhello2
    Application().Run(win)
	
xr = XmlReader.Create(StringReader(xaml))
win = XamlReader.Load(xr)

controls = {}
Waddle(win, controls)

#Make all Named buttons do something!
for butt in controls['Button']:
	controls['Button'][butt].Click += sayhello

#Make one button do something.
controls['Button']['NewSite'].Click += sayhello2
Application().Run(win)	