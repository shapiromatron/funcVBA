'''A self-logging program for examining simple wxPython behavior
   during GUI/program execution. Also introduces the wx.Notebook,
   a handy class for creating tabs in the GUI (e.g., Excel sheets).'''

import wx
import os
import sys

class ExamplePanel(wx.Panel):

    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        
        # CREATE SOME SIZERS
            # the BoxSizer places objects
            # either vertically or horizontally
        mainSizer = wx.BoxSizer(wx.VERTICAL)
            # the grid sizer uses a (row, column)
            # layout
        grid = wx.GridBagSizer(hgap=5, vgap=5)
        hSizer = wx.BoxSizer(wx.HORIZONTAL)
        
        self.quote = wx.StaticText(self, label="Your quote: ")
        # pos=(0,0) corresponds to 'x' in the diagram:
        #    x   ...   ___
        #   ...  ...   ...
        #   ___  ...   ___
        grid.Add(self.quote, pos=(0,0))
        
        # A multiline TextCtrl - This is here to show how the events work in this program, don't pay too much attention to it
        self.logger = wx.TextCtrl(self, size=(200,300), style=wx.TE_MULTILINE | wx.TE_READONLY)
        
        # A Save button
        self.button =wx.Button(self, label="Save")
        self.Bind(wx.EVT_BUTTON, self.OnClick, self.button)
        
        # the edit control - one line version.
        self.lblname = wx.StaticText(self, label="Your name :")
        grid.Add(self.lblname, pos=(1,0))
        self.editname = wx.TextCtrl(self, value="Enter here your name", size=(140,-1))
        grid.Add(self.editname, pos=(1,1))
        self.Bind(wx.EVT_TEXT, self.EvtText, self.editname)
        self.Bind(wx.EVT_CHAR, self.EvtChar, self.editname)
        
        # the combobox Control
        self.sampleList = ['friends', 'advertising', 'web search', 'Yellow Pages']
        self.lblhear = wx.StaticText(self, label="How did you hear from us ?")
        grid.Add(self.lblhear, pos=(3,0))
        self.edithear = wx.ComboBox(self, size=(95, -1), choices=self.sampleList, style=wx.CB_DROPDOWN)
        grid.Add(self.edithear, pos=(3,1))
        self.Bind(wx.EVT_COMBOBOX, self.EvtComboBox, self.edithear)
        self.Bind(wx.EVT_TEXT, self.EvtText,self.edithear)
        
        # add a spacer to the sizer
        grid.Add((10, 40), pos=(2,0))
        
        # Checkbox
        self.insure = wx.CheckBox(self, label="Do you want Insured Shipment ?")
        grid.Add(self.insure, pos=(4,0), span=(1,2), flag=wx.BOTTOM, border=5)
        self.Bind(wx.EVT_CHECKBOX, self.EvtCheckBox, self.insure)
        
        # Radio Boxes
        radioList = ['blue', 'red', 'yellow', 'orange', 'green', 'purple', 'navy blue', 'black', 'gray']
        rb = wx.RadioBox(self, label="What color would you like ?", pos=(20, 210), choices=radioList,  majorDimension=3,
                        style=wx.RA_SPECIFY_COLS)
        grid.Add(rb, pos=(5,0), span=(1,2))
        self.Bind(wx.EVT_RADIOBOX, self.EvtRadioBox, rb)
        
        hSizer.Add(grid, 0, wx.ALL, 5)
        hSizer.Add(self.logger)
        mainSizer.Add(hSizer, 0, wx.ALL, 5)
        mainSizer.Add(self.button, 0, wx.CENTER)
        self.SetSizerAndFit(mainSizer)
        
    # Bind controls to correct events
    def EvtRadioBox(self, event):
        self.logger.AppendText('EvtRadioBox: %d\n' % event.GetInt())
    def EvtComboBox(self, event):
        self.logger.AppendText('EvtComboBox: %s\n' % event.GetString())
    def OnClick(self,event):
        self.logger.AppendText(" Click on object with Id %d\n" %event.GetId())
    def EvtText(self, event):
        self.logger.AppendText('EvtText: %s\n' % event.GetString())
    def EvtChar(self, event):
        self.logger.AppendText('EvtChar: %d\n' % event.GetKeyCode())
        event.Skip()
    def EvtCheckBox(self, event):
        self.logger.AppendText('EvtCheckBox: %d\n' % event.Checked())
        
        
app = wx.App(False)
frame = wx.Frame(None, title="Demo with Noteboox")
nb = wx.Notebook(frame)

nb.AddPage(ExamplePanel(nb), "Page 1")
nb.AddPage(ExamplePanel(nb), "Page 2")
nb.AddPage(ExamplePanel(nb), "Page 3")

frame.Show()
app.MainLoop()