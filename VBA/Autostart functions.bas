Option Explicit

Sub SC_Merge()
' Keyboard Shortcut: Ctrl+m
    On Error GoTo errexit
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
		.ShrinkToFit = False
		.ReadingOrder = xlContext
		.MergeCells = True
    End With
errexit:
    Exit Sub
End Sub

Sub SC_CopyVisible()
    ' Copies only visble cells
    ' Keyboard shortuct: Ctrl+q
    On Error GoTo 0
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
0     Exit Sub
End Sub

Sub SC_SelectVisible()
' Keyboard Shortcut: Ctrl+w
    On Error Resume Next
    Selection.SpecialCells(xlCellTypeVisible).Select
    On Error GoTo 0
End Sub

