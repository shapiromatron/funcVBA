Option Explicit

Const MenuName = "&Commenter Tools"

Private Sub Workbook_Open()
   CreateMenu
   Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DeleteMenu
End Sub

Private Sub CreateMenu()
    Dim MenuObject As CommandBarPopup
    Dim SubMenuItem As CommandBarButton
    Dim MenuItem As Object
    Dim ToolbarLocation As Integer

    '****************************************
    '***** DELETE MENU IF ALREADY OPEN ******
    '****************************************
     DeleteMenu
    
    '*******************************
    '***** Add Top Level Menu ******
    '*******************************
    'decide where to add the menu bar at the end
    ToolbarLocation = Application.CommandBars("Worksheet Menu Bar").Controls.Count + 1
		
    'add menu
    Set MenuObject = Application.CommandBars("Worksheet Menu Bar").Controls.Add( _
        Type:=msoControlPopup, _
        Before:=ToolbarLocation, _
        temporary:=True)
    MenuObject.Caption = MenuName
    
    '*********************************
    '***** Add Commands to Menu ******
    '*********************************
    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "a_Main.ViewSubmissions"
    MenuItem.FaceId = 156
    MenuItem.Caption = "View &Comments"

    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "a_Main.AddExcerpt"
    MenuItem.FaceId = 156
    MenuItem.Caption = "Add &Excerpt"

    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "a_Main.ViewExcerpt"
    MenuItem.FaceId = 156
    MenuItem.Caption = "&View Excerpts"

    Set MenuItem = MenuObject.Controls.Add(Type:=msoControlButton)
    MenuItem.OnAction = "a_Main.ViewEPAresponse"
    MenuItem.FaceId = 156
    MenuItem.Caption = "Edit &Response"
End Sub

Private Sub DeleteMenu()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(MenuName).Delete
End Sub
