'''Demonstrate using a class to define a wxPython app's main visual element.'''

import wx
import os

# Create a more sophisticated Frame by adding
# features to wx.Frame via subclassing.
class ExampleFrame(wx.Frame):
    
    # define built-in properties
    def __init__(self, parent, label):
        wx.Frame.__init__(self, parent) # override wx.Frame's init with some new features
        
        # wx.StaticText is one of wx's many
        # `Control` objects. Using a few simple
        # inputs, you can specify labels, position
        # and size for a great variety of controls.
        self.quote = wx.StaticText(self, label="Your quote: %s" % label, pos=(20,30), size=(200,-1))

        self.Show()
        
        
app = wx.App(False)

# Create an ExampleFrame object.
#
# No need to save the created object
# to a variable and call that object's
# "Show" method because self.Show()
# is present in ExampleFrame's __init__ 
# function.
ExampleFrame(None, "LABEL")

# Start application
app.MainLoop()