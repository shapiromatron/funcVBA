"""
A simplified template/guide for creating a Python program with
a wxPython GUI.

USEFUL WEB LINKS:
    wxPython home page:         http://www.wxpython.org/
    Very brief 'about' page:    http://www.wxpython.org/what.php
    How to learn wxPython:      http://wiki.wxpython.org/How%20to%20Learn%20wxPython
    online documentation:       http://www.wxpython.org/onlinedocs.php


GUIDE:

(1) Read this guide in full before beginning.

    Decide what your program needs to do. Outline the
    program's functionality from beginning to end,
    conceptually. You can add to this (and any step)
    as you go along.
    
(2) Decide which to write first: backend functions, or
    the GUI. Depending on the level of complexity of
    your program, things may change as your program
    begins to take shape. Writing the GUI code and
    backend functions separately may help you simplify
    the program for the user and for you.
    
    For complex projects, you will probably need to
    develop the backend and the GUI in a more
    simultaneous manner.
    
(3) Draw your desired layout for the GUI, and label 
    controls/frames etc. Compare your drawing to the 
    conceptual layout from step (1), and make sure your 
    ideas for the GUI accomplish what's required.
    
(4) Write down how the GUI elements are connected to
    each other and to the backend functions, and what
    their purpose is.
    
(5) Code what you chose to code first in step (2).
    Test your code and make any necessary changes.

(6) Make sure you have a good idea of how to bind the
    code you wrote in step (5) to the GUI or to the
    backend, depending on which one you wrote first.
    Make any corrections necessary. Re-test your code
    if you made changes.
    
(7) Code the other portion of the program, i.e., the
    component (backend or GUI) that you did not write
    in step (5). Test your code and make any necessary
    changes.
    
(8) Attempt to connect the code from steps (5) and (7).
    This usually is just a matter of binding functions,
    creating unforseen code, and testing your results.
    
(9) Use the program from beginning to end, and try to
    test every significant combination of events. Seemingly
    small events are often hard to anticipate, but can crash
    your program, so be thorough and squash those bugs!
    This also gives you a chance to start revising your
    'rough draft' program by seeing what features should
    be added/removed/tweaked. Be sure to take notes on
    your observations.
    
(10) Revise your draft code if necessary.
"""


import wx
# PLACE ADDITIONAL IMPORTS HERE
    
    
#
#   PLACE CORE PROCESSING ('BACKEND') FUNCTIONS HERE
#
#     (note that these could live just about anywhere,
#    such as in a Python module or the class(es) below!)
#
def YourBackendFunction(optionalArg='MESSAGE GOES HERE'):
    return 'This function does something, like saying %s' % optionalArg
    
    
#
#
#   DEFINE GUI CLASS(ES)
#   
#
#
class YourGUIframe(wx.Frame):
    # you can also subclass wx.Panel and 
    # wx.Window -- just be sure to add them 
    # to a wx.Frame object later!
    
    # Initialize this object with properties
    # that always need to be defined, such
    # as menu bars, status bars, the GUI's
    # frame/window element(s), etc.
    def __init__(self, parent, title):
        '''
        Initialize this object with properties that always
        need to be defined for this object, such as menu
        bars, status bars, the GUI's frame/window element(s),
        etc.
        '''
        
        # override the Frame's init with a custom title and size
        wx.Frame.__init__(self, parent, title=title, size=(300,300))
        
        # Add some Frame elements 
            # status bar
        self.CreateStatusBar()
            # menu bar with menus
        menuBar = wx.MenuBar()
        fileMenu = wx.Menu()
        fileMenu_About = fileMenu.Append(wx.ID_ABOUT,  # 'About' event ID
                                         "About",      # menu item label
                                         "About this program") # status bar text
        fileMenu_Exit = fileMenu.Append(wx.ID_EXIT,
                                        "Exit",
                                        "Exit this program")
        menuBar.Append(fileMenu, "File")
        self.SetMenuBar(menuBar)
        
        # Add some controls! see the 'Alphabetical class reference'
        # on http://www.wxpython.org/onlinedocs.php
        # self.someControl = wx.YourControlOfChoice(args)
        # just a FEW useful controls:
        # wx.CheckBox
        # wx.CheckListBox
        # wx.ComboBox
        # wx.ColourPickerCtrl
        # wx.TextCtrl
        
            # Example controls
        self.DemoButton  = wx.Button(self, wx.ID_ANY, 
                                     label='Push for Demo',
                                     pos=(90,50))
        self.DemoMessage = wx.TextCtrl(self, wx.ID_ANY, 
                                       value='DEMO MESSAGE HERE',
                                       size=(200,-1), # '-1' is default height
                                       pos=(40,100))   # raw coords is bad form - use Sizers (see documentation)
        
        # Bind GUI elements to backend functions!
            # Note that your __init__ function can call
            # any function of this class, so you can place
            # your binding functions in a separate function
            # and just call it. This can help clean up messy code.
        self.Bind(wx.EVT_BUTTON, self.onDemoButtonClick, self.DemoButton)
        self.Bind(wx.EVT_MENU,   self.onAbout,           fileMenu_About )
        self.Bind(wx.EVT_MENU,   self.onExit,            fileMenu_Exit  )
        
        # Show the Frame
        self.Show()
        
        
    #
    #   PLACE GUI EVENT FUNCTIONS HERE
    #
    def onExit(self, event):
        '''An important function for closing properly.'''
        
        # Make sure user wants to quit
        message = 'Are you sure you want to exit?'
        caption = 'Exit'
        
        dialog = wx.MessageDialog(self, message, caption=caption)
        userChoice = dialog.ShowModal() # save Cancel or OK response from user
        dialog.Destroy()

        if userChoice == wx.ID_OK: # If user says 'OK' to exiting, exit program
            self.Close(True)
        
    def onAbout(self, event):
        '''Display a message describing this program.'''
        
        message = ('This is a template for a basic wxPython '
                   'GUI program.')
        caption = 'About this Program'
        style   = wx.OK # no 'cancel' button; just 'OK'
        
        dialog  = wx.MessageDialog(self, message, caption=caption, style=style)
        dialog.ShowModal()
        dialog.Destroy()
    
    def onDemoButtonClick(self, event):
        '''Called when self.DemoButton is clicked.'''
        
        substring = self.DemoMessage.GetValue()
        message = YourBackendFunction(optionalArg=substring)
        
        # display a MessageDialog telling
        # the user something
        dialog = wx.MessageDialog(self, message, caption='Demo', style=wx.OK)
        dialog.ShowModal() # display dialog until user closes it
        dialog.Destroy()   # ALWAYS destroy dialogs!!!
        

#
#
#   CREATE APPLICATION / START APPLICATION
#        (wxPython-specific code)
#
#
app = wx.App(False) # False instructs wx to not redirect stdout to a window or file
frame = YourGUIframe(None, "Your Program's Frame Title")
app.MainLoop() 