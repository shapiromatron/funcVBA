Attribute VB_Name = "z_Excel"
Option Explicit

Enum Tbl_LookupReturn
        '----------------------------------------------------------------
        ' Tbl_LookupReturn    - Used with Function Tbl_Lookup
        '----------------------------------------------------------------
        FirstMatch = 1
        FirstMatchRowInTbl = 2
        AllMatch = 3
        AllMatchRowsInTbl = 4
End Enum

'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|        WORKBOOK FUNCTIONS               |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************

Public Function WB_OpenOrSelect(WBDir As String, WBName As String) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' WB_OpenOrSelect    - Opens workbook if not already open, or selects open workbook
        '                    - In : WBDir As String, WBName As String (include extension)
        '                    - Out: selected workbook if avaliable, error if not available
        '                    - Last Updated: 7/4/11 by AJS
        '---------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        If WB_IsOpen(WBName) = True Then
                Set WB_OpenOrSelect = Workbooks(WBName)
        Else
                If Right(WBDir, 1) <> "\" Then WBDir = WBDir & "\"
                Set WB_OpenOrSelect = Workbooks.Open(WBDir & WBName)
        End If
IsError:
        WB_OpenOrSelect = CVErr(xlErrNA)
        Debug.Print "Error in WB_OpenOrSelect: " & Err.Number & ": " & Err.Description
End Function

Public Function WB_IsOpen(WBName As String) As Boolean
        '---------------------------------------------------------------------------------------------------------
        ' WB_IsOpen          - Check to see if workbook is open
        '                    - In : WBName As String (include ".xls" extension)
        '                    - Out: true if worbook is open, false if workbook is not open
        '                    - Last Updated: 3/6/11 by AJS
        '---------------------------------------------------------------------------------------------------------
        Dim wBook As Workbook
        On Error Resume Next
        Set wBook = Workbooks(WBName)
        If wBook Is Nothing Then 'Not open
                Set wBook = Nothing
                WB_IsOpen = False
        Else 'It is open
                WB_IsOpen = True
        End If
End Function

Public Function ColumnLetter(ColumnNumber As Variant) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' ColumnLetter  - Returns column letter of input column number, for up to 16348 columns
        '                 Tested 3/25/11 - significantly quicker than function ColumnLetter2; validated same results either way
        '                 Alternate methods: ColumnLetter2 = Application.ConvertFormula("R1C" & ColumnNumber, xlR1C1, xlA1)
        '                                    ColumnLetter2 = Mid(ColumnLetter, 2, Len(ColumnLetter) - 3)  ' - Tested 6/30/11 - Select Case is slightly quicker than using a Nested If function
        '               - In : ColumnNumber As Integer
        '               - Out: ColumnLetter as String
        '               - Last Updated: 6/30/11 by AJS
        '---------------------------------------------------------------------------------------------------------
         On Error GoTo IsError
         Select Case ColumnNumber
                Case Is > 1378
                        'special case, the first 26 column set should be subtracted , 26*26 = 676
                        ColumnLetter = Chr(Int((ColumnNumber - 26 - 1) / 676) + 64) & _
                                                        Chr(Int(((ColumnNumber - 1 - 26) Mod 676) / 26) + 65) & _
                                                        Chr(((ColumnNumber - 1) Mod 26) + 65)
                Case Is > 702   '(703-1377)
                        'includes first column, 26*26 + 26=702
                        ColumnLetter = Chr(Int(ColumnNumber / 702) + 64) & _
                                                        Chr(Int(((ColumnNumber - 1) Mod 702) / 26) + 65) & _
                                                        Chr(((ColumnNumber - 1) Mod 26) + 65)
                Case Is > 26    '(27-702)
                        ColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & _
                                                        Chr(((ColumnNumber - 1) Mod 26) + 65)
                Case Else       '(1-26)
                        ColumnLetter = Chr(ColumnNumber + 64)
        End Select
        Exit Function
IsError:
        ColumnLetter = CVErr(xlErrNA)
        Debug.Print "Error in ColumnLetter: " & Err.Number & ": " & Err.Description
End Function

Public Function Picture_AddFromFile(FN As String, ImageName As String, _
                                                                        PasteRange As Range, WidthInches As Single, _
                                                                        WidthHeight As Single) As Variant
        '----------------------------------------------------------------
        ' Picture_AddFromFile   - Adds a picture as a shape to a worksheet
        '                       - In : FN As String, ImageName As String, PasteRange As Range, WidthInches As Single, WidthHeight As Single
        '                       - Out: Boolean true if succesfully completed
        '                       - Last Updated: 5/31/11 by AJS
        '----------------------------------------------------------------
        Dim ThisShape As Shape
        On Error GoTo IsError
        Set ThisShape = PasteRange.Worksheet.Shapes.AddPicture(FN, msoFalse, msoTrue, _
                                                                                                                        PasteRange.Left, PasteRange.Top, _
                                                                                                                        Application.InchesToPoints(Width), _
                                                                                                                        Application.InchesToPoints(Height))
        ThisShape.Name = ImageName
        Picture_AddFromFile = True
        Exit Function
IsError:
        Picture_AddFromFile = CVErr(xlErrNA)
        Debug.Print "Error in Picture_AddFromFile: " & Err.Number & ": " & Err.Description
End Function

'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|        CELL COMMENT FUNCTIONS           |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************
Public Function Comment_AddPicture(CommentCell As Range, _
                                                                        PictureFN As String, _
                                                                        Optional ScaleFactor As Double = 1) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' Comment_AddPicture - Adds a picture into a comment for an Excel cell; deletes current comment
        '                    - In : Comment Cell as Range, PictureFN as string, Optional ScaleFactor as Double
        '                    - Out: Boolean true if picture comment succesfully added
        '                    - Requires: UserForm "Frm_Image" to be in the workbook in order to determine image dimensions
        '                    - Last Updated: 7/4/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        
        'DELETE EXISTING COMMENT
        On Error Resume Next
                CommentCell.Comment.Delete
        On Error GoTo IsError
        
        'CHECK TO SEE IF FILE EXISTS
        If z_Files.GetFileInfo(PictureFN, FileExists) = False Then GoTo FileNotFound
        If FileLen(PictureFN) = 0 Then GoTo FileNotFound
        
        'LOAD PICTURE INTO COMMENT
        Frm_Image.Image1.Picture = LoadPicture(PictureFN)
        CommentCell.AddComment Text:=" "
        CommentCell.Comment.Visible = False
        CommentCell.Comment.Shape.Fill.UserPicture PictureFN
        CommentCell.Comment.Shape.Height = Frm_Image.Image1.Height * ScaleFactor
        CommentCell.Comment.Shape.Width = Frm_Image.Image1.Width * ScaleFactor
        Comment_AddPicture = True
        Exit Function
FileNotFound:
        CommentCell.AddComment Text:="Image not found:" & vbNewLine & vbNewLine & PictureFN
        CommentCell.Comment.Visible = False
        Comment_AddPicture = False
        Exit Function
IsError:
        Comment_AddPicture = CVErr(xlErrNA)
        Debug.Print "Error in Comment_AddPicture: " & Err.Number & ": " & Err.Description
End Function

Public Function Comment_AddText(CommentCell As Range, _
                                                                        StringText As String, _
                                                                        Optional CommentHeight As Integer = 100, _
                                                                        Optional CommentWidth As Integer = 300) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' Comment_AddText    - Adds a comment with the text specified by the user; deletes current comment
        '                    - In : CommentCell As Range
        '                           StringText As String
        '                           Optional CommentHeight As Integer = 100
        '                           Optional CommentWidth As Integer = 300
        '                    - Out: Boolean true if comment succesfully added
        '                    - Last Updated: 3/6/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        On Error Resume Next
        CommentCell.Comment.Delete
        On Error GoTo IsError
        CommentCell.AddComment StringText
        CommentCell.Comment.Visible = False
        CommentCell.Comment.Shape.Height = CommentHeight
        CommentCell.Comment.Shape.Width = CommentWidth
        Comment_AddText = True
        Exit Function
IsError:
        Comment_AddPicture = CVErr(xlErrNA)
        Debug.Print "Error in Comment_AddText: " & Err.Number & ": " & Err.Description
End Function

'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|   EXCEL RANGE VALIDATION FUNCTIONS      |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************

Public Function Validation_DoesItExist(cellRange As Range) As Boolean
        '----------------------------------------------------------------
        ' Validation_DoesItExist    - Tests to determine if validation exists on a range
        '                           - In : CellRange As Range
        '                           - Out: Boolean TRUE if validation exists, FALSE otherwise
        '                           - Last Updated: 7/3/11 by AJS
        '----------------------------------------------------------------
        On Error GoTo IsError
                If IsNumeric(cellRange.Validation.Type) Then Validation_DoesItExist = True
        Exit Function
IsError:
        Validation_DoesItExist = False
End Function

Public Function Validation_AddList(RangeToAddValidation As Range, _
                                                                        NamedRangeName As String, _
                                                                        Optional InputTitle As String, _
                                                                        Optional InputMessage As String) As Variant
        '---------------------------------------------------------------------------------
        ' Validation_AddList    - Adds a validation list to the selected cell
        '                       - In : RangeToAddValidation As Range, NamedRange As Range
        '                       - Out: Boolean true if validation succesfully added
        '                       - Created: 3/6/11 by AJS
        '                       - Modified: 6/1/11 by AJS
        '---------------------------------------------------------------------------------
        On Error Resume Next
        RangeToAddValidation.Validation.Delete
        On Error GoTo IsError
        With RangeToAddValidation.Validation
                .Add Type:=xlValidateList, _
                                        AlertStyle:=xlValidAlertStop, _
                                        Operator:=xlBetween, _
                                        Formula1:="=" & NamedRangeName
                .InputMessage = InputMessage
                .InputTitle = InputTitle
        End With
        Validation_AddList = True
        Exit Function
IsError:
        Validation_AddList = CVErr(xlErrNA)
        Debug.Print "Error in Validation_AddList: " & Err.Number & ": " & Err.Description
End Function

Public Function Validiation_DeleteAll(cellRange As Range) As Variant
        '---------------------------------------------------------------------------------
        ' Validiation_DeleteAll - Deletes all validation in selected range
        '                       - In : CellRange As Range
        '                       - Out: Boolean true if validation succesfully deleted
        '                       - Last Updated: 5/2/11 by AJS
        '---------------------------------------------------------------------------------
        On Error GoTo IsError
        cellRange.Validation.Delete
        Validiation_DeleteAll = True
        Exit Function
IsError:
        Validiation_DeleteAll = CVErr(xlErrNA)
        Debug.Print "Error in Validiation_DeleteAll: " & Err.Number & ": " & Err.Description
End Function

Public Function Validation_WholeNumber(cellRange As Range, _
                                                                        Min As Long, Max As Long, _
                                                                        Optional InputTitle As String = "Whole Numbers Only", _
                                                                        Optional InputInstructions As String = "") As Variant
        '---------------------------------------------------------------------------------
        ' Validation_WholeNumber - Adds validation for any whole number within specified range
        '                        - In : CellRange As Range
        '                               Min As Long
        '                               Max As Long
        '                               InputTitle As String
        '                               InputInstructions As String
        '                        - Out: Boolean true if validation added, error if error
        '                        - Last Updated: 5/2/11 by AJS
        '---------------------------------------------------------------------------------
        On Error GoTo IsError
        With cellRange.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Min, Formula2:=Max
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = InputTitle
                .ErrorTitle = ""
                .InputMessage = InputInstructions
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = True
        End With
        Validation_WholeNumber = True
        Exit Function
IsError:
        Validation_WholeNumber = CVErr(xlErrNA)
        Debug.Print "Error in Validation_WholeNumber: " & Err.Number & ": " & Err.Description
End Function

Public Function Validation_FreeText(cellRange As Range, _
                                                                        InputTitle As String, _
                                                                        InputInstructions As String) As Variant
        '---------------------------------------------------------------------------------
        ' Validation_FreeText    - Adds validation for any text, but includes input instructions
        '                        - In : CellRange As Range, InputTitle As String, InputInstructions As String
        '                        - Out: Boolean true if validation added
        '                        - Last Updated: 5/2/11 by AJS
        '---------------------------------------------------------------------------------
        On Error GoTo IsError
        With cellRange.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = InputTitle
                .ErrorTitle = ""
                .InputMessage = InputInstructions
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = True
        End With
        Validation_FreeText = True
        Exit Function
IsError:
        Validation_FreeText = CVErr(xlErrNA)
        Debug.Print "Error in Validation_FreeText: " & Err.Number & ": " & Err.Description
End Function

'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|   EXCEL BORDER CELL FUNCTIONS           |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************
Public Function Borders_AddStandard(TableRange As Range) As Variant
        '----------------------------------------------------------------
        ' Borders_AddStandard   - Adds standard thin line borders around range
        '                       - In : TableRange As Range
        '                       - Out: Boolean true if borders succesfully added
        '                       - Last Updated: 3/6/11 by AJS
        '----------------------------------------------------------------
        On Error GoTo IsError
        With TableRange
                'no diagonals
           .Borders(xlDiagonalDown).LineStyle = xlNone
           .Borders(xlDiagonalUp).LineStyle = xlNone
           'left border
           .Borders(xlEdgeLeft).LineStyle = xlContinuous
           .Borders(xlEdgeLeft).Weight = xlThin
           .Borders(xlEdgeLeft).ColorIndex = xlAutomatic
           'top border
           .Borders(xlEdgeTop).LineStyle = xlContinuous
           .Borders(xlEdgeTop).Weight = xlThin
           .Borders(xlEdgeTop).ColorIndex = xlAutomatic
           'bottom border
           .Borders(xlEdgeBottom).LineStyle = xlContinuous
           .Borders(xlEdgeBottom).Weight = xlThin
           .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
           'right border
           .Borders(xlEdgeRight).LineStyle = xlContinuous
           .Borders(xlEdgeRight).Weight = xlThin
           .Borders(xlEdgeRight).ColorIndex = xlAutomatic
           'inside vertical
           If .Columns.Count > 1 Then
                   .Borders(xlInsideVertical).LineStyle = xlContinuous
                   .Borders(xlInsideVertical).Weight = xlThin
                   .Borders(xlInsideVertical).ColorIndex = xlAutomatic
           End If
           'inside horizontal
           If .Rows.Count > 1 Then
                   .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                   .Borders(xlInsideHorizontal).Weight = xlThin
                   .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
           End If
        End With
        Borders_AddStandard = True
        Exit Function
IsError:
        Borders_AddStandard = CVErr(xlErrNA)
        Debug.Print "Error in Borders_AddStandard: " & Err.Number & ": " & Err.Description
End Function

Public Function Borders_AddDblOutside(TableRange As Range) As Variant
        '----------------------------------------------------------------
        ' Borders_AddDblOutside - Adds double-line exterior border, and thin-line interior borders
        '                       - In : TableRange As Range
        '                       - Out: Boolean true if borders succesfully added
        '                       - Last Updated: 5/25/11 by AJS
        '----------------------------------------------------------------
        On Error GoTo IsError
        With TableRange
                'no diagonals
           .Borders(xlDiagonalDown).LineStyle = xlNone
           .Borders(xlDiagonalUp).LineStyle = xlNone
           'left border
           .Borders(xlEdgeLeft).LineStyle = xlDouble
           .Borders(xlEdgeLeft).Weight = xlThick
           .Borders(xlEdgeLeft).ColorIndex = xlAutomatic
           'top border
           .Borders(xlEdgeTop).LineStyle = xlDouble
           .Borders(xlEdgeTop).Weight = xlThick
           .Borders(xlEdgeTop).ColorIndex = xlAutomatic
           'bottom border
           .Borders(xlEdgeBottom).LineStyle = xlDouble
           .Borders(xlEdgeBottom).Weight = xlThick
           .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
           'right border
           .Borders(xlEdgeRight).LineStyle = xlDouble
           .Borders(xlEdgeRight).Weight = xlThick
           .Borders(xlEdgeRight).ColorIndex = xlAutomatic
           'inside vertical
           If .Columns.Count > 1 Then
                   .Borders(xlInsideVertical).LineStyle = xlContinuous
                   .Borders(xlInsideVertical).Weight = xlThin
                   .Borders(xlInsideVertical).ColorIndex = xlAutomatic
           End If
           'inside horizontal
           If .Rows.Count > 1 Then
                   .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                   .Borders(xlInsideHorizontal).Weight = xlThin
                   .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
           End If
        End With
        Borders_AddDblOutside = True
        Exit Function
IsError:
        Borders_AddDblOutside = CVErr(xlErrNA)
        Debug.Print "Error in Borders_AddDblOutside: " & Err.Number & ": " & Err.Description
End Function

'************************************************
'*/--------------------------------------------\*
'*|                                            |*
'*|  EXCEL NAMED RANGE MANIPULATION FUNCTIONS  |*
'*|                                            |*
'*\--------------------------------------------/*
'************************************************

Public Function NamedRange_Add(NamedRange As Range, _
                                                                NamedRangeName As String, _
                                                                Optional WorkbookRange As Boolean = True) As Variant
        '----------------------------------------------------------------
        ' NamedRange_Add     - Add named range to Worbook or Worksheet
        '                    - In : NamedRange As Range
        '                           NamedRangeName As String
        '                           WorkbookRange As boolean [toggles workbook or worksheet range, by default true=workbook]
        '                    - Out: Boolean true/false if succesfully completed
        '                    - Last Updated: 3/6/11 by AJS
        '----------------------------------------------------------------
        On Error GoTo IsError
        If WorkbookRange = True Then
                ActiveWorkbook.Names.Add Name:=NamedRangeName, RefersTo:="='" & NamedRange.Worksheet.Name & "'!" & NamedRange.Address
        Else
                Sheets(NamedRange.Worksheet.Name).Names.Add Name:=NamedRangeName, RefersTo:="='" & NamedRange.Worksheet.Name & "'!" & NamedRange.Address
        End If
        On Error GoTo 0
        NamedRange_Add = True
        Exit Function
IsError:
        NamedRange_Add = CVErr(xlErrNA)
        Debug.Print "Error in NamedRange_Add: " & Err.Number & ": " & Err.Description
End Function

Public Function NamedRange_AddValueIfUnqiue(Value As String, _
                                                                                        NamedRangeName As String, _
                                                                                        RangeIncludesHeader As Boolean) As Variant
        '----------------------------------------------------------------
        ' NamedRange_AddValueIfUnqiue   - Adds value to named range if it doesn't already exist
        '                               - In : Value As String
        '                                      NamedRangeName As String
        '                               - Out: Boolean TRUE if value added, FALSE if already exists, error if error
        '                               - Requires: Range_FindMatch
        '                               - Last Updated: 3/6/11 by AJS
        '----------------------------------------------------------------
        Dim NamedRange As Range
        Set NamedRange = Range(NamedRangeName)
        If IsError(Range_FindMatch(Value, NamedRange)) = True Then
                NamedRange.Worksheet.Cells(NamedRange.Row + NamedRange.Rows.Count, NamedRange.Column) = Value
                NamedRange_Add z_Excel.ExtDown(NamedRange), NamedRangeName
                Range_Sort Range(NamedRangeName), RangeIncludesHeader
                NamedRange_AddValueIfUnqiue = True
        Else
                NamedRange_AddValueIfUnqiue = False
        End If
        Exit Function
IsError:
        NamedRange_AddValueIfUnqiue = CVErr(xlErrNA)
        Debug.Print "Error in NamedRange_AddValueIfUnqiue: " & Err.Number & ": " & Err.Description
End Function

Public Function NamedRange_Replace(NamedRangeName As String, ReplaceCollection As Collection) As Boolean
    '----------------------------------------------------------------
    ' NamedRange_Replace   - Replace all fields currently listed in named range with new fields in a collection
    '                      - In : NamedRangeName As String, ReplaceCollection as Collection
    '                      - Out: Boolean TRUE if sucesfully completed, FALSE if unsuccesfull
    '                      - Last Updated: 9/28/11 by AJS
    '----------------------------------------------------------------
    On Error GoTo IsError
    Dim FirstCell As String, eachCell As Variant, Row As Long
    Dim Col As String, WS As String, FirstRow As Long
    'get information on first cell
    WS = Range(NamedRangeName).Worksheet.Name
    Col = ColumnLetter(Range(NamedRangeName).Column)
    FirstRow = Range(NamedRangeName).Row
    'clear existing cells
    For Each eachCell In Range(NamedRangeName)
        eachCell.Value = ""
    Next
    'update with new values
    If ReplaceCollection.Count > 0 Then
        Row = FirstRow
        For Each eachCell In ReplaceCollection
            Sheets(WS).Range(Col & Row) = CStr(eachCell)
            Row = Row + 1
        Next
        z_Excel.NamedRange_Add ExtDown(Sheets(WS).Range(Col & FirstRow)), NamedRangeName
    Else
        z_Excel.NamedRange_Add Sheets(WS).Range(Col & FirstRow), NamedRangeName
    End If
    NamedRange_Replace = True
    Exit Function
IsError:
    NamedRange_Replace = False
    Debug.Print "Error in NamedRange_Replace: " & Err.Number & ": " & Err.Description
End Function

'************************************************
'*/--------------------------------------------\*
'*|                                            |*
'*|  EXCEL HYPERLINK FUNCTIONS                 |*
'*|                                            |*
'*\--------------------------------------------/*
'************************************************

Public Function Hyperlink_Add(AnchorRange As Range, _
                                                                HyperlinkAddress As String, _
                                                                TextToDisplay As String) As Variant
        '----------------------------------------------------------------
        ' Hyperlink_Add          - Adds a hyperlink to a cell range
        '                        - In : AnchorRange As Range, HyperlinkAddress As String, TextToDisplay As String
        '                        - Out: Boolean true if hyperlink succesfully added
        '                        - Last Updated: 3/6/11 by AJS
        '----------------------------------------------------------------
        On Error GoTo IsError
        AnchorRange.Worksheet.Hyperlinks.Add Anchor:=AnchorRange, _
                                                                                        Address:=HyperlinkAddress, _
                                                                                        TextToDisplay:=TextToDisplay
        Hyperlink_Add = True
        Exit Function
IsError:
        Hyperlink_Add = CVErr(xlErrNA)
        Debug.Print "Error in Hyperlink_Add: " & Err.Number & ": " & Err.Description
End Function


'************************************************
'*/--------------------------------------------\*
'*|                                            |*
'*|  EXCEL LIST BOX ACTIVE OBJECT FUNCTIONS    |*
'*|                                            |*
'*\--------------------------------------------/*
'************************************************

Public Function ListBox_ReturnSelected(List_Box_OLE_Object As Variant) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' ListBox_ReturnSelected    - Returns selected fields from a multiselect list box
    '                           - In : List_Box_OLE_Object as ListBox Object
    '                           - Out: Collection of selected text or empty collection
    '                           - Last Updated: 9/28/11 by AJS
    ' Example function call:
    '       Set NewColl = ListBox_ReturnSelected(Sheet1.OLEObjects("List Box 1"))
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim ListCount As Integer
    Dim ReturnCol As New Collection
    On Error GoTo IsError
    With List_Box_OLE_Object
        ListCount = .Object.ListCount - 1
        For i = 0 To ListCount
            If .Object.Selected(i) = True Then
                ReturnCol.Add .Object.List(i)
            End If
        Next i
    End With
    Set ListBox_ReturnSelected = ReturnCol
    Exit Function
IsError:
    ListBox_ReturnSelected = CVErr(xlErrNA)
    Debug.Print "Error in ListBox_ReturnSelected: " & Err.Number & ": " & Err.Description
End Function

'************************************************
'*/--------------------------------------------\*
'*|                                            |*
'*|  EXCEL EMBEDDED OBJECT FUNCTIONS           |*
'*|                                            |*
'*\--------------------------------------------/*
'************************************************
Public Function EmbededObject_Add(FullFileName As String, _
                                                                        SheetRange As Range, _
                                                                        Optional NameInExcel As String) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' EmbededObject_Add   - embed an object (such as a text file or picture) to a worksheet
        '                     - In : FullFilename As String
        '                             SheetRange As String
        '                             Optional NameInExcel As String
        '                     - Out: Boolean true/false if succesfully completed
        '                     - Last Updated: 4/22/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim OBJ As Variant
        On Error GoTo IsError
        
        'CHECK TO SEE IF FILE EXISTS
        If z_Files.GetFileInfo(FullFileName, FileExists) = False Then GoTo FileNotFound
        If FileLen(FullFileName) = 0 Then GoTo FileNotFound

        If SheetRange.Address <> "" Then
                SheetRange.Worksheet.Activate
                SheetRange.Select
        End If
        Set OBJ = ActiveSheet.OLEObjects.Add(FileName:=FullFileName, Link:=False, DisplayAsIcon:=False)
        If NameInExcel <> "" Then OBJ.Name = NameInExcel

        EmbededObject_Add = True
        Exit Function
FileNotFound:
        EmbededObject_Add = False
        Exit Function
IsError:
        EmbededObject_Add = CVErr(xlErrNA)
        Debug.Print "Error in EmbededObject_Add: " & Err.Number & ": " & Err.Description
End Function

Public Function EmbededObject_ClearAllInWB() As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' EmbededObject_ClearAllInWB    - embed an object (such as a text file or picture) to a worksheet
        '                               - Out: Boolean true/false if succesfully completed
        '                               - Last Updated: 7/4/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Application.ScreenUpdating = False
        Dim thisObj As Object
        Dim thisWS As Worksheet
        On Error GoTo IsError
        For Each thisWS In ThisWorkbook
                For Each thisObj In thisWS.OLEObjects
                        thisObj.Delete
                Next
        Next
        EmbededObject_ClearAllInWB = True
        Exit Function
IsError:
        EmbededObject_ClearAllInWB = CVErr(xlErrNA)
        Debug.Print "Error in EmbededObject_ClearAllInWB: " & Err.Number & ": " & Err.Description
End Function

'****************************************************
'*/------------------------------------------------\*
'*|                                                |*
'*|  TABLE MANIPULATION FUNCTIONS                  |*
'*|    (standard manipulation of data stored       |*
'*|     in an Excel-based table, with              |*
'*|     one header row at top and multiple rows)   |*
'*|                                                |*
'*\------------------------------------------------/*
'****************************************************

Public Function Tbl_Lookup(Tbl As Range, _
                           ReturnColumnName As String, _
                           LookupReturn As Tbl_LookupReturn, _
                           ParamArray SearchCriteria() As Variant) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' Tbl_Lookup         - Returns a collection or a single value of matches that equal each SearchCritiera
        '                       This is a substitute to the INDEX/MATCH or SUMPRODUCT method of match multiple criteria
        '                    - In : Tbl As Range [Table Range to Search, Including Headers]
        '                           ReturnColumnName As String [Name of Return Column in Header]
        '                           LookupReturn As Tbl_LookupReturn [FirstMatch, FirstMatchRow, AllMatch, AllMatchRows]
        '                           ParamArray SearchCriteria() As Variant [Match criteria, in this format: Array("Name=Bob","Age=13")]
        '                    - Out: FirstMatch will be string of match
        '                           FirstMatchRow will be the relative row references in the table
        '                           AllMatch will be a collection of ALL matches
        '                           AllMatchRows will be a collection of all relative row references in the table
        '                           IF NO MATCH, WILL RETURN AN ERROR!
        '                    - Last Updated: 7/4/11
        '---------------------------------------------------------------------------------------------------------
        '   EXAMPLE FUNCTION CALLS:
        '   ------------------------
        '    Debug.Print "FirstMatch: " & Tbl_Lookup(Range("Tbl") "Hair Color", FirstMatch, Array("Name=Andy", "Age=2"))
        '    For Each eachItem In Tbl_Lookup(Range("Tbl"), "Hair Color", AllMatch, Array("Name=Andy", "Age=2"))
        '        Debug.Print "AllMatch: " & CStr(eachItem)
        '    Next
        '---------------------------------------------------------------------------------------------------------
        Dim eachItem As Variant, EachRow As Variant
        Dim ColName() As String, MatchCriteria() As String, ColNum() As Integer
        Dim SplitVal2() As String, i As Integer, ReturnCol As Integer
        Dim MatchFound As Boolean
        Dim ReturnCollection As New Collection
                                
        'BREAK UP ALL MATCH CRITERIA
        On Error GoTo IsError
        ReDim ColName(0 To UBound(SearchCriteria(0)))
        ReDim MatchCriteria(0 To UBound(SearchCriteria(0)))
        ReDim ColNum(0 To UBound(SearchCriteria(0)))
        For i = 0 To UBound(SearchCriteria(0))
                SplitVal2 = Split(SearchCriteria(0)(i), "=")
                ColName(i) = SplitVal2(0)
                MatchCriteria(i) = SplitVal2(1)
                ColNum(i) = Range_FindMatch(SplitVal2(0), Tbl.Rows(1).Cells)
        Next i
        
        'FIND RETURN COLUMN
        ReturnCol = Range_FindMatch(ReturnColumnName, Tbl.Rows(1).Cells)
        
        'FIND MATCHES IN TABLE
        For Each EachRow In Tbl.Columns(1).Cells    'Search each row in table
                MatchFound = True
                For i = 0 To UBound(SearchCriteria(0))  'Search each match criteria
                        If MatchCriteria(i) <> EachRow.Offset(0, ColNum(0) - 1).Value Then
                                MatchFound = False
                                Exit For
                        End If
                Next i
                If MatchFound = True Then
                        ' FirstMatch = 1, FirstMatchRow = 2, AllMatch = 3, AllMatchRows = 4
                        Select Case LookupReturn
                                Case 1
                                        Tbl_Lookup = EachRow.Offset(0, ReturnCol - 1).Value
                                        Exit Function
                                Case 2
                                        Tbl_Lookup = EachRow.Row
                                        Exit Function
                                Case 3
                                        ReturnCollection.Add EachRow.Offset(0, ReturnCol - 1).Value
                                Case 4
                                        ReturnCollection.Add EachRow.Row
                        End Select
                End If
        Next
        'return collection or error, depending on if collection is available
        If ReturnCollection.Count > 0 Then
                Set Tbl_Lookup = ReturnCollection
        Else
                Tbl_Lookup = CVErr(xlErrNA)
        End If
        Exit Function
IsError:
        Tbl_Lookup = CVErr(xlErrNA)
        Debug.Print "Error in Tbl_Lookup: " & Err.Number & ": " & Err.Description
End Function

Public Function Tbl_ReturnUniqueList(SearchRange As Range) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' Tbl_ReturnUniqueList  - Returns a collection of unique values in the specified range
        '                       - In : SearchRange As Range
        '                       - Out: A string collection of unique values
        '                       - Last Updated: 6/24/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        ' Example call:
        ' Dim ThisCollection as New Collection
        ' Set ThisCollection = Tbl_ReturnUniqueList(SearchRange)
        '-----------------------------------------------------------------------------------------------------------
        Dim eachRng As Range
        Dim UniqueCollection As New Collection
        Dim CollectionItem As Variant
        Dim Unique As Boolean
        
        On Error GoTo IsError
        For Each eachRng In SearchRange
                Unique = True
                For Each CollectionItem In UniqueCollection
                        If eachRng.Value = CollectionItem Then
                                Unique = False
                                Exit For
                        End If
                Next
                If Unique = True Then UniqueCollection.Add eachRng.Value
        Next
        Set Tbl_ReturnUniqueList = UniqueCollection
        Exit Function
IsError:
        Tbl_ReturnUniqueList = CVErr(xlErrNA)
        Debug.Print "Error in Tbl_ReturnUniqueList: " & Err.Number & ": " & Err.Description
End Function

Public Function Tbl_GetHeaderColumn(SearchField As String, _
                                    SearchRange As Range, _
                                    Optional ReturnNumeric As Boolean = False) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' Tbl_GetHeaderColumn   - Returns the column of the header, either numeric or column letter
        '                       - In : SearchField As String
        '                               SearchRange As Range
        '                               Optional ReturnNumeric As Boolean = False [by default returns Column Letter]
        '                       - Out: Column Letter or Number of match, or error if not found
        '                       - Last Updated: 6/24/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim ReturnColumn As Variant
        
        On Error GoTo IsError
        ReturnColumn = SearchRange.Column + WorksheetFunction.Match(SearchField, SearchRange.Rows(1).Cells, False) - 1
        If ReturnNumeric = False Then ReturnColumn = z_Excel.ColumnLetter(CLng(ReturnColumn))
        Tbl_GetHeaderColumn = ReturnColumn
        Exit Function
IsError:
        Tbl_GetHeaderColumn = CVErr(xlErrNA)
        Debug.Print "Error in Tbl_GetHeaderColumn: " & Err.Number & ": " & Err.Description
End Function

'****************************************************
'*/------------------------------------------------\*
'*|                                                |*
'*|  EXTEND RANGE FUNCTIONS                        |*
'*|                                                |*
'*\------------------------------------------------/*
'****************************************************

Public Function ExtTbl(Rng As Range, _
                       Optional RowOffset As Long = 0, _
                       Optional ColOffset As Long = 0) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' ExtTbl             - Extends the table down to the first blank at bottom of top right row/column
        '                           will stop at the first blank row
        '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
        '                    - Out: Tbl_ExtTbl as Range
        '                    - Last Updated: 4/7/11 by AJS
        '---------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        Set ExtTbl = z_Excel.ExtRight(z_Excel.ExtDown(Rng.Offset(RowOffset, ColOffset), 0, 0), 0, 0)
        Exit Function
IsError:
        ExtTbl = CVErr(xlErrNA)
        Debug.Print "Error in ExtTbl: " & Err.Number & ": " & Err.Description
End Function

Public Function ExtDown(Rng As Range, _
                        Optional RowOffset As Long = 0, _
                        Optional ColOffset As Long = 0) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' ExtDown            - Extends the selected range down to the final non-blank row in current table;
        '                           will stop at the first blank row
        '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
        '                    - Out: Tbl_ExtDown as Range
        '                    - Last Updated: 4/7/11 by AJS
        '---------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        Set Rng = Rng.Offset(RowOffset, ColOffset)
        If IsEmpty(Rng.Offset(1, 0)) Then
                Set ExtDown = Rng
        Else
                Set ExtDown = Range(Rng, Rng.End(xlDown))
        End If
        Exit Function
IsError:
        ExtDown = CVErr(xlErrNA)
        Debug.Print "Error in ExtDown: " & Err.Number & ": " & Err.Description
End Function

Public Function ExtRight(Rng As Range, _
                         Optional RowOffset As Long = 0, _
                         Optional ColOffset As Long = 0) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' ExtRight           - Extends the selected range down to the final non-blank column in current table;
        '                           will stop at the first blank column
        '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
        '                    - Out: ExtRight as Range
        '                    - Last Updated: 4/7/11 by AJS
        '---------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        Set Rng = Rng.Offset(RowOffset, ColOffset)
        If IsEmpty(Rng.Offset(0, 1)) Then
                Set ExtRight = Rng
        Else
                Set ExtRight = Range(Rng, Rng.End(xlToRight))
        End If
        Exit Function
IsError:
        ExtRight = CVErr(xlErrNA)
        Debug.Print "Error in ExtRight: " & Err.Number & ": " & Err.Description
End Function

Public Function ExtDownNonBlank(Rng As Range, _
                                Optional RowOffset As Long = 0, _
                                Optional ColOffset As Long = 0) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' ExtDownNonBlank    - Extends the range down to the first non-blank at bottom-left of selected range;
        '                           will stop at the first blank row where a formula evaluates to a value
        '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
        '                    - Out: ExtDownNonBlank as Range
        '                    - Created: 5/15/11 by GH
        '                    - Last Updated: 6/1/11 by AJS
        '---------------------------------------------------------------------------------------------------------
        Dim NewRng As Range
        Dim LastRow As Long
        
        On Error GoTo IsError
        Set NewRng = ExtDown(Rng, RowOffset, ColOffset)
        For LastRow = NewRng.Rows.Count To 1 Step -1
                If Not NewRng.Cells(LastRow, 1) = "" Then
                  Exit For
                End If
        Next LastRow
        Set ExtDownNonBlank = NewRng.Resize(LastRow, NewRng.Columns.Count)
        Exit Function
IsError:
        ExtDownNonBlank = CVErr(xlErrNA)
        Debug.Print "Error in ExtDownNonBlank: " & Err.Number & ": " & Err.Description
End Function

Public Function ExtTblNonBlank(Rng As Range, _
                               Optional RowOffset As Long = 0, _
                               Optional ColOffset As Long = 0) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' ExtTblNonBlank     - Extends the range down and right to the first non-blank at bottom-left of selected range;
        '                           will stop at the first blank row where a formula evaluates to a value
        '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
        '                    - Out: ExtTblNonBlank as Range
        '                    - Created: 5/15/11 by GH
        '                    - Last Updated: 6/1/11 by AJS
        '---------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        Set ExtTblNonBlank = z_Excel.ExtRight(ExtDownNonBlank(Rng.Offset(RowOffset, ColOffset), 0, 0), 0, 0)
        Exit Function
IsError:
        ExtTblNonBlank = CVErr(xlErrNA)
        Debug.Print "Error in ExtTblNonBlank: " & Err.Number & ": " & Err.Description
End Function

Public Function ExtAllTbl(ByRef Rng As Range, _
                          Optional RowOffset As Long = 0, _
                          Optional ColOffset As Long = 0) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' ExtAllTbl          - Extends the selected range right and down to the final non-blank row if
        '                           leftmost column and the final non-blank row in topmost row
        '                    - In : ByRef Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
        '                    - Out: ExtAllTbl as Range
        '                    - Last Updated: 3/9/11 by AJS (originally from GH)
        '---------------------------------------------------------------------------------------------------------
        Dim RightmostColumn As Long
        Dim BottomRow As Long
        
        On Error GoTo IsError
        Set Rng = Rng.Offset(RowOffset, ColOffset)
        BottomRow = Application.Max(LastRow(Rng.Worksheet, Rng.Column), Rng.Row)
        RightmostColumn = Application.Max(LastColumn(Rng.Worksheet, Rng.Row), Rng.Column)
        Set ExtAllTbl = Rng.Resize(RowSize:=(BottomRow - Rng.Rows.Count - Rng.Row + 2), _
                                                                ColumnSize:=(RightmostColumn - Rng.Columns.Count - Rng.Column + 2))
        Exit Function
IsError:
        ExtAllTbl = CVErr(xlErrNA)
        Debug.Print "Error in ExtAllTbl: " & Err.Number & ": " & Err.Description
End Function

Public Function ExtAllDown(ByRef Rng As Range, _
                           Optional RowOffset As Long = 0, _
                           Optional ColOffset As Long = 0) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' ExtAllDown         - Extends the selected range down to the final non-blank row in leftmost column
        '                    - In : ByRef Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
        '                    - Out: ExtAllDown as Range
        '                    - Last Updated: 3/9/11 by AJS (originally from GH)
        '---------------------------------------------------------------------------------------------------------
        Dim BottomRow As Long
        Set Rng = Rng.Offset(RowOffset, ColOffset)
        BottomRow = Application.Max(LastRow(Rng.Worksheet, Rng.Column), Rng.Row)
        Set ExtAllDown = Rng.Resize(RowSize:=(BottomRow - Rng.Rows.Count - Rng.Row + 2))
        Exit Function
IsError:
        ExtAllDown = CVErr(xlErrNA)
        Debug.Print "Error in ExtAllDown: " & Err.Number & ": " & Err.Description
End Function

Public Function ExtAllRight(ByRef Rng As Range, _
                            Optional RowOffset As Long = 0, _
                            Optional ColOffset As Long = 0) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' ExtAllRight        - Extends the selected range right to the final non-blank row in topmost row
        '                    - In : ByRef Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
        '                    - Out: ExtAllRight as Range
        '                    - Last Updated: 3/9/11 by AJS (originally from GH)
        '---------------------------------------------------------------------------------------------------------
        Dim RightmostColumn As Long
        Set Rng = Rng.Offset(RowOffset, ColOffset)
        RightmostColumn = Application.Max(LastColumn(Rng.Worksheet, Rng.Row), Rng.Column)
        Set ExtAllRight = Rng.Resize(ColumnSize:=(RightmostColumn - Rng.Columns.Count - Rng.Column + 2))
        Exit Function
IsError:
        ExtAllRight = CVErr(xlErrNA)
        Debug.Print "Error in ExtAllRight: " & Err.Number & ": " & Err.Description
End Function

Private Function LastRow(ByVal OfSheet As Worksheet, Optional ByVal InColumn As Long = 0) As Long
        '---------------------------------------------------------------------------------------------------------
        ' LastRow            - Returns the number of the last used row in the specified sheet [and column]
        '                    - In : ByVal OfSheet As Worksheet, Optional ByVal InColumn As Long = 0
        '                    - Out: LastRow as Range
        '                    - Last Updated: 3/9/11 by AJS (originally from GH)
        '---------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        If InColumn = 0 Then
                LastRow = OfSheet.UsedRange.Row + OfSheet.UsedRange.Rows.Count - 1
        Else
                LastRow = OfSheet.Cells(Application.Rows.Count, InColumn).End(xlUp).Row
        End If
        Exit Function
IsError:
        LastRow = CVErr(xlErrNA)
        Debug.Print "Error in LastRow: " & Err.Number & ": " & Err.Description
End Function

Private Function LastColumn(ByVal OfSheet As Worksheet, Optional ByVal InRow As Long = 0) As Integer
        '---------------------------------------------------------------------------------------------------------
        ' LastColumn         - Returns the number of the last used column in the specified sheet [and row]
        '                    - In : ByVal OfSheet As Worksheet, Optional ByVal InRow As Long = 0
        '                    - Out: LastColumn as Range
        '                    - Last Updated: 3/9/11 by AJS (originally from GH)
        '---------------------------------------------------------------------------------------------------------
        Dim i As Integer, letter As String
        On Error GoTo IsError
        If InRow = 0 Then
                i = OfSheet.UsedRange.Columns.Count + 1
                Do
                        i = i - 1
                        LastColumn = OfSheet.UsedRange.Columns(i).Cells(1, 1).Column
                        letter = z_Excel.ColumnLetter(LastColumn)
                Loop Until (Application.WorksheetFunction.CountA(OfSheet.Range(letter & ":" & letter)) > 0 Or i < 2)
        Else
                LastColumn = OfSheet.Cells(InRow, Application.Columns.Count).End(xlToLeft).Column
        End If
        Exit Function
IsError:
        LastColumn = CVErr(xlErrNA)
        Debug.Print "Error in LastColumn: " & Err.Number & ": " & Err.Description
End Function

'****************************************************
'*/------------------------------------------------\*
'*|                                                |*
'*|  TABLE MANIPULATION FUNCTIONS                  |*
'*|    (standard manipulation of data stored       |*
'*|     in an Excel-based table, with              |*
'*|     one header row at top and multiple rows)   |*
'*|                                                |*
'*\------------------------------------------------/*
'****************************************************

Public Function Range_Set(WSName As String, LowCol As Variant, LowRow As Variant, Optional HighCol As Variant, Optional HighRow As Variant) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Range_Set             - Returns a range set using various permutations of inputs
    '                       -     Range_Set("Main", 2, 3) = 'Main'!$B$3
    '                       -     Range_Set("Main", "B", 3) = 'Main'!$B$3
    '                       -     Range_Set("Main", 2, 3, 4, 5) = 'Main'!$B$3:$D$5
    '                       -     Range_Set("Main", "B", 3, "D", 5) = 'Main'!$B$3:$D$5
    '                       - In : WSName As String, LowCol As Variant, LowRow As Variant,
    '                               Optional HighCol As Variant, Optional HighRow As Variant
    '                       - Out: Range Object if succesful, error if otherwise
    '                       - Last Updated: 8/22/11
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo IsError:
    If IsMissing(HighRow) And IsMissing(HighCol) Then
        Set Range_Set = Sheets(WSName).Cells(LowRow, LowCol)
    ElseIf IsNumeric(LowCol) And IsNumeric(HighCol) Then
        Set Range_Set = Sheets(WSName).Range( _
                            Sheets(WSName).Cells(LowRow, LowCol), _
                            Sheets(WSName).Cells(HighRow, HighCol))
    Else
        Set Range_Set = Sheets(WSName).Range( _
                            LowCol & LowRow & ":" & _
                            HighCol & HighRow)
    End If
    Exit Function
IsError:
    Range_Set = CVErr(xlErrNA)
    Debug.Print "Error in Range_Set: " & Err.Number & ": " & Err.Description
End Function

Public Function Range_CopyPasteValues(CopyRange As Range, PasteRange As Range) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Range_CopyPasteValues - Copies and pastes values only
    '                       - In : CopyRange As Range
    '                               PasteRange As Range
    '                       - Last Updated: 7/4/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
    PasteRange.Value = CopyRange.Value
    Range_CopyPasteValues = True
    Exit Function
IsError:
    Range_CopyPasteValues = CVErr(xlErrNA)
    Debug.Print "Error in Range_CopyPasteValues: " & Err.Number & ": " & Err.Description
End Function

Public Function Range_CopyPasteAll(CopyRange As Range, PasteRange As Range) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Range_CopyPasteAll    - Copies and pastes everything, including formatting
    '                           Known bug- only copies and pastes visible cells
    '                       - In : CopyRange As Range
    '                               PasteRange As Range
    '                       - Last Updated: 7/4/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
    CopyRange.Copy
    PasteRange.PasteSpecial (xlPasteAll)
    Application.CutCopyMode = False
    Range_CopyPasteAll = True
    Exit Function
IsError:
    Range_CopyPasteAll = CVErr(xlErrNA)
    Debug.Print "Error in Range_CopyPasteAll: " & Err.Number & ": " & Err.Description
End Function

Public Function Range_Sort(wsRange As Range, RangeIncludesHeader As Boolean) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Range_Sort         - Sorts the selected range, w/ or w/o header
    '                    - In : wsRange As Range, Header As Boolean
    '                    - Last Updated: 3/9/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
    Dim HeaderType As Integer
    Select Case RangeIncludesHeader
            Case True
                    HeaderType = 1
            Case False
                    HeaderType = 2
    End Select
    wsRange.Sort Key1:=wsRange.Cells(1), _
                    Order1:=xlAscending, _
                    Header:=HeaderType, _
                    MatchCase:=False, _
                    Orientation:=xlTopToBottom, _
                    DataOption1:=xlSortNormal
    Range_Sort = True
    Exit Function
IsError:
    Range_Sort = False
    Debug.Print "Error in Range_Sort: " & Err.Number & ": " & Err.Description
End Function

Public Function Range_GetMaxRow(FirstCellInRange As Range) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Range_GetMaxRow      Returns the maximum row down using ExtDown from a Range
    '                      In : FirstCellInRange As Range
    '                      Out: Maximum Row ID or an error
    '                      Last Updated: 8/18/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo IsError:
    Range_GetMaxRow = FirstCellInRange.Row + z_Excel.ExtDown(FirstCellInRange).Rows.Count - 1
    Exit Function
IsError:
    Range_GetMaxRow = CVErr(xlErrNA)
    Debug.Print "Error in Function Range_GetMaxRow: " & Err.Number & ": " & Err.Description
End Function

Public Function Range_ConvertTo1DArray(InputRange As Range) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Range_ConvertTo1DArray      Converts a range of values into a 1-D array
    '                             In : InputRange As Range
    '                             Last Updated: 7/28/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
    Dim eachCell As Range
    Dim OutputArray() As Variant
    Dim Counter As Long
    
    ReDim OutputArray(0 To InputRange.Cells.Count - 1)
    Counter = 0
    For Each eachCell In InputRange
        OutputArray(Counter) = eachCell.Value
        Counter = Counter + 1
    Next
    Range_ConvertTo1DArray = OutputArray
    Exit Function
IsError:
        Range_ConvertTo1DArray = CVErr(xlna)
        Debug.Print "Error in Range_ConvertTo1DArray: " & Err.Number & ": " & Err.Description
End Function

Private Function Range_ConvertToArray(InputRange As Range) As Variant
    '-----------------------------------------------------------------------------------------------------------
    ' Range_ConvertToArray      Converts a range of values into an array, size number of rows x number of columns
    '                           In : InputRange As Range
    '                           Last Updated: 7/28/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
    Range_ConvertToArray = InputRange.Value2
    Exit Function
IsError:
        Range_ConvertToArray = CVErr(xlna)
        Debug.Print "Error in Range_ConvertToArray: " & Err.Number & ": " & Err.Description
End Function

Public Function Range_FindMatch(SearchString As String, SearchRange As Range) As Variant
        '----------------------------------------------------------------
        ' Range_FindMatch       - Searches named range to see if string matches
        '                       - In : SearchString As String, SearchRange As Range
        '                       - Out: Index of matched string, if found, FALSE if not match
        '                       - Last Updated: 3/24/11 by AJS
        '----------------------------------------------------------------
        On Error GoTo IsError
            Range_FindMatch = WorksheetFunction.Match(SearchString, SearchRange, False)
        Exit Function
IsError:
        Range_FindMatch = CVErr(xlErrNA)
End Function

Private Function DoesFileExist(FN As String) As Variant
        Dim fso As Object
        On Error GoTo IsError
        Set fso = CreateObject("Scripting.FileSystemObject")
        DoesFileExist = fso.FileExists(FN)
        Exit Function
IsError:
        DoesFileExist = CVErr(xlErrNA)
        Debug.Print "Error in Private Function DoesFileExist: " & Err.Number & ": " & Err.Description
End Function

'****************************************************
'*/------------------------------------------------\*
'*|                                                |*
'*|  ERROR TRAPPING FUNCTIONS                      |*
'*|    (traps errors that would cause runtime      |*
'*|     errors unless otherwise caught, such as    |*
'*|     does a worksheet exist or a chart exist)   |*
'*|                                                |*
'*\------------------------------------------------/*
'****************************************************

Public Function ErrorTrap_WSExists(WSName As String) As Boolean
    '----------------------------------------------------------------
    ' ErrorTrap_WSExists    - Returns TRUE if worksheet exists, FALSE if it doesn't
    '                       - Last Updated: 9/1/11 by AJS
    '----------------------------------------------------------------
On Error GoTo IsError:
    If Sheets(WSName).Name <> "" Then
        ErrorTrap_WSExists = True
    End If
    Exit Function
IsError:
    ErrorTrap_WSExists = False
End Function

Public Function ErrorTrap_ChartExists(ChartName As String) As Boolean
    '----------------------------------------------------------------
    ' ErrorTrap_ChartExists - Returns TRUE if chart exists, FALSE if it doesn't
    '                       - Last Updated: 9/1/11 by AJS
    '----------------------------------------------------------------
On Error GoTo IsError:
    If Charts(ChartName).Name <> "" Then
        ErrorTrap_ChartExists = True
    End If
    Exit Function
IsError:
    ErrorTrap_ChartExists = False
End Function
