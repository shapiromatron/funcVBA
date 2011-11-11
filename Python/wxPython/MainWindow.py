'''Create a very simple text-editor.'''

import wx
import os

# Subclass wx.Frame to add new
# functionality for our app
class MainWindow(wx.Frame):

    def __init__(self, parent, title):
        self.dirname=''
    
        # A '-1' in the size parameter instructs wxWidgets to use the default size.
        # In this case, we select 200px width and the default height.
        wx.Frame.__init__(self, parent, title=title, size=(200,-1))
        
        # Create a simple text box as the text entry
        # control for the user
        self.control = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        
        # status bar in bottom of the window
        self.CreateStatusBar() 
        
        # Set up the menu
        filemenu = wx.Menu()
            # wx.ID_ABOUT and wx.ID_EXIT are standard IDs provided by wxWidgets.
        filemenuOpen  = filemenu.Append(wx.ID_OPEN, 
                                    "Open", "Open a file")
        filemenuAbout = filemenu.Append(wx.ID_ABOUT, 
                                    "About", " Information about this program")
        filemenuExit  = filemenu.Append(wx.ID_EXIT, 
                                    "Exit", " Terminate the program")
            # Create the menubar
        menuBar = wx.MenuBar()
        menuBar.Append(filemenu, "File") # adding the 'filemenu' to the MenuBar
        self.SetMenuBar(menuBar)  # Adding the MenuBar to the Frame content
        
        # Set events
        # Bind(event code, callback function, GUI control)
        self.Bind(wx.EVT_MENU, self.OnOpen, filemenuOpen)
        self.Bind(wx.EVT_MENU, self.OnAbout, filemenuAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, filemenuExit)
        
        # We will use some wx.Sizer objects
        # to layout our controls in a more
        # sophisticated way that doing raw
        # coordinate positioning. This is
        # more portable and readable, also.
        self.sizer2 = wx.BoxSizer(wx.HORIZONTAL)
        self.buttons = []
        
        # Make 6 buttons
        for i in range(6):
            self.buttons.append(wx.Button(self, -1, "Button %r" % i))
            self.sizer2.Add(self.buttons[i], 1, wx.EXPAND)
            
        # Use some sizers to see layout options
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        # sizer.Add takes 3 arguments:
            # arg 1:    the control to include in the sizer
            # arg 2:    the relative weight factor, which sizes
            #           the control relative to others. A value
            #           of 0 means that the control or sizer
            #           will NOT grow!
            # arg 3:    resizing behavior; wx.EXPAND and wx.GROW
            #           resize the controls when necessary.
            #           wx.SHAPED keeps aspect ratios the same.
            #           If arg 2 is 0, this arg can be a
            #           wx.ALIGN* argument.
        # items appear in `sizer` in the order they're Add-ed
        self.sizer.Add(self.sizer2, 0, wx.EXPAND)
        self.sizer.Add(self.control, 1, wx.EXPAND)
        
        
        # Layout sizers
            # THIS MUST BE DONE TO HOOK THE SIZER TO THE WINDOW/FRAME!
        self.SetSizerAndFit(self.sizer) 
        self.SetAutoLayout(True) # automatically re-draw layout when the window is resized
        
        # Show the frame
        self.Show()
        
    def OnAbout(self, e):
        # A message dialog box with an OK button; 
        # wx.OK is a standard ID in wxWidgets.
        dialog = wx.MessageDialog(self, "A small text editor", "About Sample Editor", wx.OK)
        dialog.ShowModal() # show the Dialog
        dialog.Destroy()   # when user closes Dialog, be SURE to Destroy it!
        
    def OnExit(self, e):
        '''Close the program.'''
        self.Close(True)
        
    def OnOpen(self, e):
        '''Open a file.'''
        
        # Very useful dialog.
        #     wx.FileDialog(paent, window title, directory to open in,  file names to view, frame ID)
        dlg = wx.FileDialog(self, "Choose a file", self.dirname, "", "*.*", wx.OPEN)
        
        # dlg.ShowModal() returns a code based on
        # how the user closes it/responds to it.
        # If they hit the OK button, open the
        # chosen file.
        if dlg.ShowModal() == wx.ID_OK:
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            with open(os.path.join(self.dirname, self.filename), 'r') as infile:
                # set the text control's value to the string
                # represented in the file
                self.control.SetValue(infile.read())
        dlg.Destroy()
        
app = wx.App(False) # do not divert text to stdout
frame = MainWindow(None, "Sample editor")
app.MainLoop()