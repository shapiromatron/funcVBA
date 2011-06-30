Option Explicit

'/-----------------------------------------------\
'| FOR EACH commands for Pivot Table componenets |
'\-----------------------------------------------/
Private Sub EachPivotTable()
    Dim SheetName As String
    Dim PivField As PivotField
    Dim PivItem As PivotItem
    Dim PivTable As PivotTable
    Dim wks As Worksheet
    
    Sheets(SheetName).Activate
    'for each pivot table on a worksheet
    For Each PivTable In Sheets(SheetName).PivotTables
        
        'for each page field in a pivot table
        For Each PivField In PivTable.PageFields
            PivField.CurrentPage = "(All)"  'select all pivot fields to be presented
        Next PivField
        
        'for each pivot field in a Pivot Table
        For Each PivField In PivTable.PivotFields
            'For each Item in a Pivot Field
            For Each PivItem In PivField.PivotItems
                PivItem.Visible = True
            Next PivItem
        Next PivField
    Next PivTable
    'Reset for each variables
    Set PivItem = Nothing
    Set PivField = Nothing
    Set PivTable = Nothing
End Sub

'/---------------------------------------------------\
'| Fill in value in a cell from field in pivot table |
'\---------------------------------------------------/
Function PivotFillDown(TargetCell As Range)
    
    Dim OffSet As Long
    Dim Value
    Dim StartRow As Long, StartCol As Long
    
    If TargetCell.Value = "" Then
        StartRow = TargetCell.Row
        StartCol = TargetCell.Column
        OffSet = 0
        Do While TargetCell.Row - OffSet >= 1
            OffSet = OffSet - 1
            Value = ActiveSheet.Cells(StartRow + OffSet, StartCol)
            If IsEmpty(Value) = False Then
                PivotFillDown = Value
                Exit Do
            End If
        Loop
    Else
        PivotFillDown = TargetCell.Value
    End If
End Function

'/---------------------------------------------------\
'| See if a value in a field exists in a pivot table |
'\---------------------------------------------------/
Function TextExistsInPivot(SearchText As String, SheetName As String, PivTblName As String, PivFieldName As String) As Boolean
    Dim PivItem As PivotItem
    TextExistsInPivot = False
    For Each PivItem In Sheets(SheetName).PivotTables(PivTblName).PivotFields(PivFieldName).PivotItems
        If SearchText = PivItem.Name Then
            'change the pivot field to this value
            Sheets(SheetName).PivotTables(PivTblName).PivotFields(PivFieldName).CurrentPage = PivItem.Name
            Sheets(SheetName).PivotTables(PivTblName).PivotCache.Refresh
            'return true
            TextExistsInPivot = True
            Exit For
        End If
    Next PivItem
End Function
