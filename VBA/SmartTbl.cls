VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SmartTbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Collection of commonly used functions when dealing with standard data tables in Excel.
'  This assumes that the data table has a data structure where the first row is the header,
'  and each column name in the header is unique. Each row is a new data field.
'
'Every time a method is executed, the table range is updated. Depending on the table extension
'  type, the table range is updated in various manners:
'  1) Fixed = table range is fixed and never updated
'  2) Extend = table range is extended down and to the right from teh top
'
'This class module requires the following function module:
'   z_Excel.bas
'
' ajs to add:
' smart extend, where it does a standard extright, and then goes through each column to find the max row, in counta
'   this should replace extdown, so the first row doesn't need to be complete, and it'll work if rows are hidden
'   and/or filtered
' add row, where the inputs are a dictionary, with value and column name, appended to end of table
' find row, based on matching criteria, may return first or multiple as options
' get dictionary, creates a dictionay for each row; or may create multiple dictionaries if mutliple rows are specified

'###################
'# CLASS VARIABLES #
'###################
Enum TableExtensionType:
    ExtendAll = 1
    Extend = 2
    Fixed = 3
End Enum

'REQUIRED INPUTS (AND EXAMPLES IN COMMENTS)
Private pTblRange As Range
Private pExtendType As TableExtensionType

'##############
'# INITIALIZE #
'##############
Private Sub Class_Initialize()
    Set pTblRange = Nothing
End Sub

'##############################
'# LET FUNCTIONS (SET VALUES) #
'##############################
Public Property Let SetTbl_ExtendAll(TopLeft As Range)
    Set pTblRange = TopLeft
    pExtendType = ExtendAll
    UpdateTblRange
End Property
Public Property Let SetTbl_Extend(TopLeft As Range)
    Set pTblRange = TopLeft
    pExtendType = Extend
    UpdateTblRange
End Property
Public Property Let SetTbl_Exact(TableRange As Range)
    Set pTblRange = TableRange
    pExtendType = Fixed
End Property

'##################
'# PUBLIC METHODS #
'##################
Public Function get_tbl_range() As Range
'Return the range of the smart table
    UpdateTblRange
    Set get_tbl_range = pTblRange
End Function

Public Function get_max_row() As Long
'Return max row as long, or -999 if error
    On Error GoTo IsError
    UpdateTblRange
    get_max_row = pTblRange.Row + pTblRange.Rows.Count - 1
    Exit Function
IsError:
    get_max_row = -999
End Function

Public Function get_max_column() As Long
'Return max column as long, or -999 if error
    On Error GoTo IsError
    UpdateTblRange
    get_max_column = pTblRange.Column + pTblRange.Columns.Count - 1
    Exit Function
IsError:
    get_max_column = -999
End Function

Public Function get_header_column(HeaderName As String) As Long
'Return absolute reference column of header if name found, or -999 if error
    On Error GoTo IsError
    UpdateTblRange
    get_header_column = pTblRange.Column - 1 + _
                        WorksheetFunction.Match(HeaderName, pTblRange.Rows(1), False)
    Exit Function
IsError:
    get_header_column = -999
End Function

Public Function get_unique_values_in_column(HeaderName As String) As Collection
'Return all the unique values in the header range in a Collection
    Dim ColumnHeader As Long
    Dim SearchRange As Range
    UpdateTblRange
    ColumnHeader = get_header_column(HeaderName)
    If ColumnHeader = -999 Then
        Set SearchRange = Nothing
    Else
        Set SearchRange = pTblRange.Worksheet.Range( _
                            Cells(pTblRange.Row, ColumnHeader), _
                            Cells(get_max_row, ColumnHeader))
    End If
    Set get_unique_values_in_column = ReturnUnique(SearchRange)
End Function

Public Function delete_all()
'Delete everything, including the first row
    UpdateTblRange
    pTblRange.Delete shift:=xlUp
    UpdateTblRange
End Function

Public Function delete_all_data_rows()
'Delete everything except the first row
    UpdateTblRange
    pTblRange.Offset(1, 0).Delete shift:=xlUp
    UpdateTblRange
End Function

'###################
'# PRIVATE METHODS #
'###################
Private Sub UpdateTblRange()
'Update the table range to make sure the current table is updated
    Select Case pExtendType:
        Case ExtendAll:
            Set pTblRange = MaxRange(pTblRange.Cells(1, 1))
        Case Extend:
            Set pTblRange = z_Excel.ExtTbl(pTblRange.Cells(1, 1), 0, 0)
        Case Fixed:
            Set pTblRange = pTblRange
    End Select
End Sub

Private Function MaxRange(TopLeft As Range) As Range
'Returns a range with the top-left cell in range, and the bottom right cell
' equal to the maximum range used on the current worksheet.
' Note that special care must be used with this function if multiple tables
' exist on the same worksheet
    Dim FullRange As Range
    Set FullRange = TopLeft.Worksheet.UsedRange
    Set MaxRange = TopLeft.Worksheet.Range( _
                    Cells(TopLeft.Row, TopLeft.Column), _
                    Cells(FullRange.Row + FullRange.Rows.Count - 1, _
                          FullRange.Column + FullRange.Columns.Count - 1))
End Function

Private Function ReturnUnique(InputRange As Range) As Collection
'Returns a collection of unique values from an input range.
' Note that the range is offset by one row, assuming that the first
' row in the table is the header, and therefore is not included
    Dim eachCell As Range, cUnique As New Collection
    On Error Resume Next   'required so error doesn't pop-up when adding duplicate value
    If InputRange.Rows.Count > 1 Then
        For Each eachCell In InputRange.Offset(1, 0)
            If eachCell.Value <> "" Then cUnique.Add eachCell.Value, CStr(eachCell.Value)
        Next eachCell
    End If
    Set ReturnUnique = cUnique
    On Error GoTo 0
End Function

Function ReturnDict(Row As Long) As Variant
'Returns a Dictionary with the key being the header name,
' and value equaling the value in a row. If Row is -999 error code,
' returns an empty dictionary with header keys but empty values.
' All dictionaries include a key value "_Row_" that includes the row number
    Dim TempDict As Variant
    Dim eachRng As Range
    Dim Value As Variant
    Set TempDict = CreateObject("Scripting.Dictionary")
    TempDict.RemoveAll
    TempDict.Add "_Row_", Row
    For Each eachRng In pTblRange.Rows(1).Cells
        If Row > 0 Then
            Value = eachRng.Offset(Row - pTblRange.Row, 0).Value
        Else
            Value = Empty
        End If
        TempDict.Add eachRng.Value, Value
    Next
    Set ReturnDict = TempDict
End Function

Public Function Lookup_ReturnFirstMatch(ReturnHeaderName As String, _
                        ParamArray SearchCriteria() As Variant) As Variant
'Returns the value in ReturnColumnName for the first row which matches all
'   search criteria.  If no match found, returns empty string
    Dim ReturnColumn As Long, eachRow As Variant, Criteria As Variant
    UpdateTblRange
    ReturnColumn = get_header_column(ReturnHeaderName)
    If ReturnColumn = -999 Then GoTo IsError
    If UBound(SearchCriteria) = -1 Then GoTo IsError
    Criteria = ParseCriteria(SearchCriteria)
    If Criteria(0)(0) = -999 Then GoTo IsError
    'For each row
    For Each eachRow In pTblRange.Offset(1, 0).Rows
        If CheckForMatch(eachRow.Row, Criteria) = True Then
          Lookup_ReturnFirstMatch = pTblRange.Worksheet.Cells(eachRow.Row, ReturnColumn)
          Exit For
        End If
    Next
    Exit Function
IsError:
    Lookup_ReturnFirstMatch = Empty
End Function

Public Function Lookup_ReturnAllMatches(ReturnHeaderName As String, _
                        ParamArray SearchCriteria() As Variant) As Collection
'Returns a collection of values in ReturnColumnName for the first row which matches all
'   search criteria.  If no match found, returns empty collection
    Dim ReturnColumn As Long, eachRow As Variant, Criteria As Variant
    Dim thisCollection As New Collection
    UpdateTblRange
    ReturnColumn = get_header_column(ReturnHeaderName)
    If ReturnColumn = -999 Then GoTo IsError
    If UBound(SearchCriteria) = -1 Then GoTo IsError
    Criteria = ParseCriteria(SearchCriteria)
    If Criteria(0)(0) = -999 Then GoTo IsError
    'For each row
    For Each eachRow In pTblRange.Offset(1, 0).Rows
        If CheckForMatch(eachRow.Row, Criteria) = True Then
            thisCollection.Add pTblRange.Worksheet.Cells(eachRow.Row, ReturnColumn)
        End If
    Next
    Set Lookup_ReturnAllMatches = thisCollection
    Exit Function
IsError:
    Set Lookup_ReturnAllMatches = thisCollection    'empty collection
End Function

Public Function Lookup_ReturnFirstRow(ParamArray SearchCriteria() As Variant) As Long
'Returns the row number for the first row which matches all search criteria.
' If no match found, returns -999
    Dim eachRow As Variant, Criteria As Variant
    UpdateTblRange
    If UBound(SearchCriteria) = -1 Then GoTo IsError
    Criteria = ParseCriteria(SearchCriteria)
    If Criteria(0)(0) = -999 Then GoTo IsError
    'For each row
    For Each eachRow In pTblRange.Offset(1, 0).Rows
        If CheckForMatch(eachRow.Row, Criteria) = True Then
            Lookup_ReturnFirstRow = eachRow.Row
            Exit For
        End If
    Next
    Exit Function
IsError:
    Lookup_ReturnFirstRow = -999
End Function

Public Function Lookup_ReturnAllRows(ParamArray SearchCriteria() As Variant) As Collection
'Returns a collection of row numbers which matches all search criteria.
' If no match found, returns an empty collection
    Dim eachRow As Variant, Criteria As Variant
    Dim thisCollection As New Collection
    UpdateTblRange
    If UBound(SearchCriteria) = -1 Then GoTo IsError
    Criteria = ParseCriteria(SearchCriteria)
    If Criteria(0)(0) = -999 Then GoTo IsError
    'For each row
    For Each eachRow In pTblRange.Offset(1, 0).Rows
        If CheckForMatch(eachRow.Row, Criteria) = True Then
            thisCollection.Add eachRow.Row
        End If
    Next
    Set Lookup_ReturnAllRows = thisCollection
    Exit Function
IsError:
    Set Lookup_ReturnAllRows = thisCollection    'empty collection
End Function

Public Function Lookup_ReturnFirstDict(ParamArray SearchCriteria() As Variant) As Variant
'Returns the dictionary of values for the first row which matches all criteria;
' key values are each column name, values are the values with the row.
' If no match found, returns a dictionary with keys but no values
    Dim eachRow As Variant, Criteria As Variant
    Dim Dict As Variant
    UpdateTblRange
    If UBound(SearchCriteria) = -1 Then GoTo IsError
    Criteria = ParseCriteria(SearchCriteria)
    If Criteria(0)(0) = -999 Then GoTo IsError
    Set Lookup_ReturnFirstDict = ReturnDict(-999)   'set empty dictionary
    'For each row
    For Each eachRow In pTblRange.Offset(1, 0).Rows
        If CheckForMatch(eachRow.Row, Criteria) = True Then
            Set Lookup_ReturnFirstDict = ReturnDict(eachRow.Row)
            Exit For
        End If
    Next
    Exit Function
IsError:
    'return empty dictionary
End Function

Public Function Lookup_ReturnAllDicts(ParamArray SearchCriteria() As Variant) As Variant
'Returns an array of dictionaries of values for the first row which matches all criteria;
' key values are each column name, values are the values with the row.
' If no matches are found, returns an array of one dictionary with keys but no values
    Dim eachRow As Variant, Criteria As Variant
    Dim eachDict As Variant, allDicts() As Variant
    Dim NumMatch As Integer
    UpdateTblRange
    If UBound(SearchCriteria) = -1 Then GoTo IsError
    Criteria = ParseCriteria(SearchCriteria)
    If Criteria(0)(0) = -999 Then GoTo IsError
    NumMatch = 0
    For Each eachRow In pTblRange.Offset(1, 0).Rows
        If CheckForMatch(eachRow.Row, Criteria) = True Then
            NumMatch = NumMatch + 1
            Set eachDict = ReturnDict(eachRow.Row)
            If NumMatch = 1 Then
                ReDim allDicts(0 To 0)
            Else
                ReDim Preserve allDicts(LBound(allDicts) To UBound(allDicts) + 1)
            End If
            Set allDicts(UBound(allDicts)) = eachDict
        End If
    Next
    If NumMatch = 0 Then GoTo IsError
    Lookup_ReturnAllDicts = allDicts
    Exit Function
IsError:
    'return empty dictionary
    ReDim allDicts(0 To 0)
    Set eachDict = ReturnDict(-999)   'set empty dictionary
    Set allDicts(UBound(allDicts)) = eachDict
    Lookup_ReturnAllDicts = allDicts
End Function

Private Function CheckForMatch(SearchRow As Long, MatchCriteria As Variant) As Boolean
'Given a search row and the match criteria, returns true if all criteria are matched, false if otherwise
    Dim i As Integer
    CheckForMatch = True
    For i = 0 To UBound(MatchCriteria(0))
        If pTblRange.Worksheet.Cells(SearchRow, MatchCriteria(0)(i)) <> MatchCriteria(1)(i) Then
            CheckForMatch = False
            Exit For
        End If
    Next i
End Function
Private Function ParseCriteria(ParamArray SearchCriteria() As Variant) As Variant
'Breaks up parse criteria to match row and search string
    Dim MatchCriteria() As String, ColNum() As Long, SplitVal2() As String, i As Integer
    On Error GoTo IsError
    ReDim MatchCriteria(0 To UBound(SearchCriteria(0)))
    ReDim ColNum(0 To UBound(SearchCriteria(0)))
    For i = 0 To UBound(SearchCriteria(0))
        SplitVal2 = Split(SearchCriteria(0)(i), "=")
        ColNum(i) = get_header_column((SplitVal2(0)))
        MatchCriteria(i) = SplitVal2(1)
        If ColNum(i) = -999 Then GoTo IsError
    Next i
    ParseCriteria = Array(ColNum, MatchCriteria)
    Exit Function
IsError:
    MatchCriteria(0) = -999
    ColNum(0) = -999
    ParseCriteria = Array(ColNum, MatchCriteria)
End Function




