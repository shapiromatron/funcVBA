'''wxPython's "Hello, World!" example.'''



# To find this and other introductory examples, go to:
#       http://wiki.wxpython.org/Getting%20Started
#
# Extremely useful core documentation:
#       http://www.wxpython.org/onlinedocs.php
#      --> the 'Alphabetical class reference' is
#          a must-have guide for exploring wx objects



# Import the wx module, which wraps
# the reliable C++ GUI library, wx
import wx

# Create a new app; 'False' means don't redirect stdout/stderr to a window.
#                   'True' means redirect to a window, or you can optionally
#                          provide the argument `filename` to output to a file.
app = wx.App(False)  

# A `Frame` object is a top-level window.
#   A Frame object can contain other wx display objects
#   in this (suggested) hierarchy:
#         1. Frame (must be highest level)
#         2. Window(s)
#         3. Menu bar(s)
#         4. Panel(s)
#         5. Widgets/Controls
#       wx.Frame(Window parent, int id, string title)
frame = wx.Frame(None, wx.ID_ANY, "Hello World")

# Show the frame using True. (Use 'False' to hide it.)
frame.Show(True)     

# Construct the application's GUI elements
# in the appropriate order and await user
# interaction.
app.MainLoop()

