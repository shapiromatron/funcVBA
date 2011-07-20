Option Explicit

Enum SaveAsDialogType
    ExcelFiles = 1
    WordFiles = 3
    AnyExtension = 5
    Custom = 99
End Enum
'   /-------------------------------------------------------------------------------\
'   |   NEW FUNCTIONS TO BE PROPERLY QA/QC AND CATEGORIZED                          |
'   \-------------------------------------------------------------------------------/

Function FormatNumberToPrint(ByVal Number As Double) As String
    Dim NumAfterDecimal As Double
    Dim ConvertedText As String
    Dim ShowDec As String
    
	'DON'T ADD PRECISION WHEN THERE ISN'T ANY
    NumAfterDecimal = InStr(1, CStr(Number), ".")
    If NumAfterDecimal > 0 Then
        NumAfterDecimal = Len(CStr(Number)) - NumAfterDecimal
        ShowDec = "."
    Else
        ShowDec = ""
    End If
    
    Select Case Number
		Case 0
            ConvertedText = "0"
        Case Is > 100000
            ConvertedText = Format(Number, "0.0E+00")
        Case Is > 10000
            ConvertedText = Format((Number), "0,000")
        Case Is > 1000  'and < 10,000
            ConvertedText = Format((Number), "0,000")
        Case Is > 100   'and < 1,000
            ConvertedText = Format(Number, "0")
        Case Is > 10    'and < 100
            ConvertedText = Format(Number, "0" & ShowDec & WorksheetFunction.Rept("0", WorksheetFunction.Min(1, NumAfterDecimal)))
        Case Is > 1     'and < 10
            ConvertedText = Format(Number, "0" & ShowDec & WorksheetFunction.Rept("0", WorksheetFunction.Min(2, NumAfterDecimal)))
        Case Is > 0.1   'and < 1
            ConvertedText = Format(Number, "0." & WorksheetFunction.Rept("0", WorksheetFunction.Min(3, NumAfterDecimal)))
        Case Is > 0.01  'and < 0.1
            ConvertedText = Format(Number, "0." & WorksheetFunction.Rept("0", WorksheetFunction.Min(4, NumAfterDecimal)))
        Case Is > 0.001 'and < 0.01
            ConvertedText = Format(Number, "0." & WorksheetFunction.Rept("0", WorksheetFunction.Min(5, NumAfterDecimal)))
        Case Else   ' <= 0.001
            ConvertedText = Format(Number, "0.00E+00")
    End Select
    FormatNumberToPrint = ConvertedText
End Function

Function DateLastModified(ByVal FN As String) As String
    '-----------------------------------------------------------------------------------------------------------
    ' DateLastModified   - Returns the last date an item was modified
    '                    - In : Filename as String
    '                    - Out: A date in string form
    '                    - Last Updated: 6/30/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
	Dim oFS As Object, f As Variant
	on error goto IsError:
    Set oFS = CreateObject("Scripting.FileSystemObject")
    Set f = oFS.getfile(FN)
    DateLastModified = CStr(f.DateLastModified)
	Exit Function
isError:
	DateLastModified = CVErr(xlErrNA)
    Debug.Print "Error " & Err.Number & ": " & Err.Description & ": " Err.Location
End Function

Function IsValidFileName(sFileName As String) As Boolean
'http://www.exceltoolset.com/a-simple-vba-function-for-testing-a-filename-is-valid/
    Dim IllegalChar As Variant
    Dim i As Long
    Dim result As Boolean
 
    IllegalChar = Array("/", ":", "*", "?", "<", ">", "|", "")
    result = True
 
    For i = LBound(IllegalChar) To UBound(IllegalChar)
        If InStr(1, sFileName, IllegalChar(i)) > 0 Then
            result = False
            Exit Function
        End If
    Next i
    IsValidFileName = result
End Function

Function MakeDirFullPath(Path As String) As Boolean
    Dim UncreatedPaths As Collection, EachPath As Variant
    Set UncreatedPaths = New Collection
    Dim NewPath As String
    
    On Error GoTo PathNotCreated
    NewPath = Path
    Do While ICF.DoesFolderExist(NewPath) = False
        UncreatedPaths.Add NewPath
        NewPath = GetParentFolder(NewPath)
    Loop
    Do While UncreatedPaths.Count > 0
        MkDir UncreatedPaths(UncreatedPaths.Count)
        UncreatedPaths.Remove UncreatedPaths.Count
    Loop
    MakeDirFullPath = True
    Exit Function
PathNotCreated:
    MakeDirFullPath = False
End Function

Function ReturnUniqueList(SearchRange As Range) As Collection
    '-----------------------------------------------------------------------------------------------------------
    ' ReturnUniqueList   - Returns a collection of unique values in the specified range
    '                    - In : SearchRange As Range
    '                    - Out: A string collection of unique values
    '                    - Last Updated: 6/24/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    Dim EachRng As Range
    Dim UniqueCollection As New Collection
    Dim CollectionItem As Variant
    Dim Unique As Boolean
    
    For Each EachRng In SearchRange
        Unique = True
        For Each CollectionItem In UniqueCollection
            If EachRng.Value = CollectionItem Then
                Unique = False
                Exit For
            End If
        Next
        If Unique = True Then UniqueCollection.Add EachRng.Value
    Next
    Set ReturnUniqueList = UniqueCollection
End Function

Function SplitTextReturnOne(InputString As String, ReturnValue As Integer) As String
'Parse text string and return desired word
    Dim StringVariant As Variant
    On Error GoTo isErr:
    
    StringVariant = SplitText(InputString, " ")
    SplitTextReturnOne = StringVariant(ReturnValue)
    Exit Function
isErr:
    SplitTextReturnOne = "SplitTextError"
End Function

Function GetParentFolder(ByVal FN As String)
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    GetParentFolder = oFS.GetParentFolderName(FN)
End Function

Function CopyPasteValuesExcel(CopyRange As Range, PasteRange As Range) As Variant
    On Error GoTo IsError:
    PasteRange.Value = CopyRange.Value
    Exit Function
IsError:
    CopyPasteValuesExcel = CVErr(xlErrNA)
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Function

Function CopyPasteExcel(CopyRange As Range, PasteRange As Range) As Variant
    On Error GoTo IsError:
    CopyRange.Copy
    PasteRange.PasteSpecial (xlPasteAll)
    Application.CutCopyMode = False
    Exit Function
IsError:
    CopyPasteExcel = CVErr(xlErrNA)
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Function

Function FileList(ByVal PathName As String, Optional ByVal FileFilter As String = "*.*") As Collection
    '-----------------------------------------------------------------------------------------------------------
    ' FileList           - Returns a collection of files in a given foldre with the specified filter
    '                    - In : PathName As String, Optional FileFilter As String
    '                    - Out: A string collection of file names in the specified folder
    '                    - Created: Greg Haskins
    '                    - Last Updated: 6/17/11 by AJS
    '-----------------------------------------------------------------------------------------------------------

    'Replacement for Application.FileSearch
    Dim sTemp As String, sHldr As String
    Dim RetVal As New Collection
    If Right$(PathName, 1) <> "\" Then PathName = PathName & "\"
    sTemp = Dir(PathName & FileFilter)
    If sTemp = "" Then
        Set FileList = RetVal
        Exit Function
    Else
        RetVal.Add sTemp
    End If
    Do
        sHldr = Dir
        If sHldr = "" Then Exit Do
        'sTemp = sTemp & "|" & sHldr
        RetVal.Add sHldr
    Loop
    'FileList = Split(sTemp, "|")
    Set FileList = RetVal
End Function

Function SortRange(wsRange As Range, Header As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    ' SortRange          - Sorts the selected range, w/ or w/o header
    '                    - In : wsRange As Range, Header As Boolean
    '                    - Last Updated: 3/9/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
    Dim HeaderType As Integer
    Select Case Header
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
    SortRange = True
    Exit Function
IsError:
    SortRange = False
End Function

Function Regex(SearchString As String, RegExPattern As String, Optional CaseSensitive As Boolean = False) As String
'http://www.regular-expressions.info/dotnet.html
'http://www.tmehta.com/regexp/
'http://www.ozgrid.com/forum/showthread.php?t=37624&page=1
'
'Example function call would return "ty1234"
'MsgBox RegEx("qwerty123456uiops123456", "[a-z][A-Z][0-9][0-9][0-9][0-9]", False)
'
    Dim RE As Object, REMatches As Object
    Set RE = CreateObject("vbscript.regexp")
    With RE
        .MultiLine = False
        .Global = False
        .IgnoreCase = Not (CaseSensitive)
        .Pattern = RegExPattern
    End With
    Set REMatches = RE.Execute(SearchString)
    If REMatches.Count > 0 Then
        Regex = REMatches(0)
    Else
        Regex = False
    End If
End Function

Function Printf(ByVal FormatWithPercentSign As String, ParamArray InsertArray()) As String
'http://www.freevbcode.com/ShowCode.asp?ID=9342
    Dim ResultString As String
    Dim Element As Variant
    Dim FormatLocation As Long

    If IsMissingValue(InsertArray()) Then
        'raise an error
    End If
    
    ResultString = FormatWithPercentSign
    For Each Element In InsertArray
        FormatLocation = InStr(ResultString, "%")
        ResultString = Left$(ResultString, FormatLocation - 1) & Element & Right$(ResultString, Len(ResultString) - FormatLocation - 1)
    Next
    Printf = ResultString
End Function

Public Function AddToArrayIfUnique(ByVal NewString As String, ArrayName() As Variant) As Variant

'    Dim ArrayName() As Variant
'    Let ArrayName = [{"Andy", "Cara", "Josh"}]
'    ArrayName() = AddToArrayIfUnique("Bill", ArrayName())
'    ArrayName() = AddToArrayIfUnique("Andy", ArrayName())

    Dim EachValue As Variant
    Dim Duplicate As Boolean
    Duplicate = False
    For Each EachValue In ArrayName
        If NewString = EachValue Then
            Duplicate = True
            Exit For
        End If
    Next
    If Duplicate = False Then
        ReDim Preserve ArrayName(LBound(ArrayName()) To UBound(ArrayName()) + 1)
        ArrayName(UBound(ArrayName)) = NewString
    End If
    Let AddToArrayIfUnique = ArrayName
End Function

Function DoesFileExist(ByVal FN As String) As Boolean
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    DoesFileExist = oFS.FileExists(FN)
End Function

Function DoesFolderExist(ByVal FN As String) As Boolean
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    DoesFolderExist = oFS.FolderExists(FN)
End Function

Function DoesValidationExist(CellRange As Range) As Boolean
    '----------------------------------------------------------------
    ' DoesValidationExist - Tests to determine if validation exists on a range
    '                     - In : CellRange As Range
    '                     - Out: Boolean true if validation exists, false otherwise
    '                     - Created: 6/1/11 by AJS
    '                     - Last Updated: 6/1/11 by AJS
    '----------------------------------------------------------------
    On Error GoTo IsError:
        If IsNumeric(CellRange.Validation.Type) Then DoesValidationExist = True
    Exit Function
IsError:
    DoesValidationExist = False
End Function

Function GetBaseName(ByVal FN As String) As String
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    GetBaseName = oFS.GetBaseName(FN)
End Function

Function GetFileName(ByVal FN As String) As String
    Dim oFS
    Set oFS = CreateObject("Scripting.FileSystemObject")
    GetFileName = oFS.GetFileName(FN)
End Function

Function GetExtension(ByVal FN As String) As String
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    GetExtension = oFS.GetExtensionName(FN)
End Function

Function GetPath(ByVal FN As String) As String
    Dim oFS As Object
    Set oFS = CreateObject("Scripting.FileSystemObject")
    GetPath = oFS.GetParentFolderName(FN) & "\"
End Function

Function IsTextFound(ByVal FindText As String, ByVal WithinText As String) As Boolean
    '----------------------------------------------------------------
    ' IsTextFound       - Returns true if text is found, false if otherwise
    '                   - In : ByVal FindText As String, ByVal WithinText As String
    '                   - Out: Boolean true if found, false if not
    '                   - Last Updated: 4/12/11 by AJS
    '----------------------------------------------------------------
    If InStr(1, WithinText, FindText, vbTextCompare) > 0 Then
        IsTextFound = True
    Else
        IsTextFound = False
    End If
End Function

Public Function R_MinReturn(SearchString As Range) As Long
    Dim LowVal As Double
    Dim sncell As Range
    On Error GoTo IsError
    LowVal = WorksheetFunction.Min(SearchString)
    For Each sncell In SearchString
        If sncell.Value = LowVal Then
            R_MinReturn = sncell.Row
            Exit Function
        End If
    Next
IsError:
    R_MinReturn = -999
End Function
Function R_MaxReturn(SearchString As Range) As Long
    Dim LowVal As Double
    Dim sncell As Range
    On Error GoTo IsError
    LowVal = WorksheetFunction.Max(SearchString)
    For Each sncell In SearchString
        If sncell.Value = LowVal Then
            R_MaxReturn = sncell.Row
            Exit Function
        End If
    Next
IsError:
    R_MaxReturn = -999
End Function

Public Function GetFolder(ByVal Title As String, Optional ByVal strPath As String) As String
    'http://www.mrexcel.com/forum/showthread.php?t=294728
    Dim fldr As FileDialog
    Dim sItem As String
    If IsEmpty(strPath) Then strPath = "C:\"
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .InitialView = msoFileDialogViewDetails
        .Title = Title
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Public Function GetWordFile(Optional TitleName As String, Optional strPath As String) As String
    If strPath = "" Then strPath = "C:\"
    If TitleName = "" Then TitleName = "Select Word File"
    GetWordFile = Application.GetOpenFilename(FileFilter:="Word Files (*.docx; *.docm; *.doc; *.dot; *.dotx; *.dotm), *.docx *.docm *.xlsb *.doc *.dot; *.dotx; *.dotm", Title:=TitleName, MultiSelect:=False)
End Function

Public Function GetExcelFile(strPath As String) As String
    GetExcelFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls; *.xlsx; *.xlsm; *.xlsb), *.xls *.xlsm *.xlsb *.xlsx", Title:="Select Excel file", MultiSelect:=False)
End Function
 
Public Function SaveAsDialog(FileFilter As SaveAsDialogType, ByVal Title As String, Optional ByVal InitialDirectory As String, Optional ByVal CustomFilter As String) As String
    'http://msdn.microsoft.com/en-us/library/bb209903(v=office.12).aspx
    'new to office 2007
    Dim Looping As Boolean
    Select Case FileFilter
        Case 1
            CustomFilter = "Excel 2007 File (*.xlsx),*.xlsx,Macro-enabled Excel 2007 File (*.xlsm),*.xlsm,Excel 2003 File (*.xls),*.xls"
        Case 3
            CustomFilter = "Word 2007 File (*.docx),*.docx,Macro-enabled Word 2007 File (*.docm),*.docm,Excel 2003 File (*.doc),*.doc"
        Case 5
            CustomFilter = "Any File (*.*),*.*"
        Case 99
            CustomFilter = CustomFilter
    End Select
    If InitialDirectory = "" Then InitialDirectory = ThisWorkbook.Path
    Do
        SaveAsDialog = Application.GetSaveAsFilename(InitialDirectory, _
                                            CustomFilter, _
                                            1, _
                                            Title)
        If DoesFileExist(SaveAsDialog) = True Then
            If vbYes = MsgBox("File already exists, replace existing file?" & vbNewLine & vbNewLine & SaveAsDialog, vbYesNo, "Replace existing file?") Then
                Looping = False
            Else
                Looping = True
            End If
        Else
            Looping = False
        End If
    Loop Until Looping = False
End Function
 
Public Function InputBox_GetFile(FileFilter As String, Title As String, MultiSelect As Boolean) As String
    'FileFilter:="Excel Files (*.xls; *.xlsx; *.xlsm; *.xlsb), *.xls *.xlsx *.xlsm *xlsb
    'Title:="Select Excel file"
    InputBox_GetFile = Application.GetOpenFilename(FileFilter:=FileFilter, Title:=Title, MultiSelect:=MultiSelect)
End Function
 

Public Function GetFileExtension(ByVal FileName As String) As String
    '----------------------------------------------------------------
    ' GetFileExtension      - Gets the extension of any filename, or returns -999 if file not found
    '                       - In : ByVal FileName As String
    '                       - Out: Extension for file, or -999 if not found
    '                       - Last Updated: 3/23/11 by AJS
    '----------------------------------------------------------------
    On Error GoTo IsError
    Dim fso As Object
    If DoesFileExist(FileName) = False Then GoTo IsError
    On Error GoTo IsError
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileExtension = (fso.GetExtensionName(FileName))
    Set fso = Nothing
    Exit Function
IsError:
    GetFileExtension = -999
End Function
Public Function BuildXMLText(ByVal FieldName As String, ByVal Value As String, Optional NumTabs As Integer = 0) As String
    '----------------------------------------------------------------
    ' BuildXMLText          - Builds XML text string
    '                       - In : ByVal FieldName As String, ByVal Value As String, Optional NumTabs As Integer = 0
    '                       - Out: XML test string for a single line:   <FieldName>Value</Field>
    '                       - Last Updated: 3/23/11 by AJS
    '----------------------------------------------------------------
    BuildXMLText = AddTab(NumTabs) & "<" & FieldName & ">" & Value & "</" & FieldName & ">"
End Function
'
'Public Function DirExists(ByVal DirectoryName As String) As Boolean
'    '----------------------------------------------------------------
'    ' DirExists             - Searches named range to see if string matches
'    '                       - In : ByVal SearchString As String, SearchRange As Range
'    '                       - Out: Index of matched string, if found, -999 if no match
'    '                       - Last Updated: 3/23/11 by AJS
'    '----------------------------------------------------------------
'    If Len(Dir(DirectoryName)) = 0 Then
'        DirExists = False
'    Else
'        DirExists = True
'    End If
'End Function

Public Function FindMatch(SearchString As String, SearchRange As Range) As Long
    '----------------------------------------------------------------
    ' FindMatch             - Searches named range to see if string matches
    '                       - In : SearchString As String, SearchRange As Range
    '                       - Out: Index of matched string, if found, -999 if no match
    '                       - Last Updated: 3/24/11 by AJS
    '----------------------------------------------------------------
    On Error GoTo IsError:
        FindMatch = WorksheetFunction.Match(SearchString, SearchRange, False)
    Exit Function
IsError:
    FindMatch = -999
End Function

Public Function FileCopy2(ByVal SourceFile As String, ByVal DestinationFile As String) As Boolean
    '----------------------------------------------------------------
    ' FileCopy2             - Revised version of FileCopy that will return TRUE when file is actually copied
    '                       - In : SourceFile As String, DestinationFile As String
    '                       - Out: Boolean true if file is succesfully copied; false otherwise
    '                       - Last Updated: 3/23/11 by AJS
    '----------------------------------------------------------------
    If DoesFileExist(SourceFile) = False Then
        MsgBox "Error- file does not exist and cannot be copied:" & ICF.AddNewLine(2) & _
            SourceFile, vbCritical, "Error in FileCopy2"
        GoTo IsError
    End If
    If DoesFileExist(DestinationFile) = True Then Kill2 DestinationFile
    Do While DoesFileExist(DestinationFile) = False
        On Error Resume Next
        FileCopy SourceFile, DestinationFile
        On Error GoTo 0
    Loop
    FileCopy2 = True
    Exit Function
IsError:
    FileCopy2 = False
End Function

Public Function Kill2(ByVal PathName As String) As Boolean
    '----------------------------------------------------------------
    ' Kill2             - Deletes file; continues until succesfullly deleted
    '                   - In : ByVal PathName As String
    '                   - Out: Boolean true if file is succesfully removed
    '                   - Last Updated: 3/23/11 by AJS
    '----------------------------------------------------------------
    On Error Resume Next
        Do While DoesFileExist(PathName) = True
            Kill PathName
        Loop
    On Error GoTo 0
    Kill2 = True
End Function

'
'   /-------------------------------------------------------------------------------\
'   |   ADD EXCEL OBJECTS FUNCTION LIBRARY                                          |
'   |-------------------------------------------------------------------------------|
'   |                       |                                                       |
'   | AddNamedRange         |   Adds a named range (worksheet or workbook level)    |
'   | AddHyperlink          |   Adds a hyperlink to the address of the cell range   |
'   | AddStandardBorders    |   Adds standard thin solid black borders around Range |
'   | AddValidationList     |   Adds a Validation list to the selected cell         |
'   | AddCommentPicture     |   Adds comment with picture in background             |
'   | AddPictureToSheet     |   Adds picture as a shape to a worksheet              |
'   | AddCommentText        |   Adds comment with text                              |
'   | AddEmbededObject      |   Adds embedded object to workbook                    |
'   |                       |                                                       |
'   \-------------------------------------------------------------------------------/
'
'   /-------------------------------------------------------------------------------\
'   |   RETURN RANGE FUNCTIONS                                                      |
'   |-------------------------------------------------------------------------------|
'   |                       |                                                       |
'   | ExtTbl                | Extends selected table to the bottom row/right column |
'   |                       |   (stops at first blank)                              |
'   | ExtDown               | Extends selected column to the bottom row in tbl      |
'   |                       |   (stops at first blank)                              |
'   | ExtRight              | Extends selected row to the rightmost column in tbl   |
'   |                       |   (stops at first blank)                              |
'   | ExtDownNonBlank       | Extends selected column to the bottom row in tbl      |
'   |                       |   (stops at first non-blank formula)                  |
'   | ExtTblNonBlank        | Extends selected table to the bottom row/right column |
'   |                       |   (stops at first non-blank formula)                  |
'   | ExtAllTbl             | Extends table to bottom row/right column              |
'   |                       |   (stops at bottom/rightmost regardless of blanks)    |
'   | ExtAllDown            | Extends row to final non-blank cell in column         |
'   |                       |   (stops at bottom/rightmost regardless of blanks)    |
'   | ExtAllRight           | Extends columns to final non-blank cell in row        |
'   |                       |   (stops at bottom/rightmost regardless of blanks)    |
'   |                       |                                                       |
'   \-------------------------------------------------------------------------------/
'
'   /-------------------------------------------------------------------------------\
'   |   RETURN BOOLEAN FUNCTIONS                                                    |
'   |-------------------------------------------------------------------------------|
'   |                       |                                                       |
'   | DoesFileExist         |  Check to see if a file exists                        |
'   | IsWBOpen              |  Check to see if workbook is open                     |
'   |                       |                                                       |
'   \-------------------------------------------------------------------------------/
'
'   /-------------------------------------------------------------------------------\
'   |   RETURN STRING FUNCTIONS                                                     |
'   |-------------------------------------------------------------------------------|
'   |                       |                                                       |
'   | ColumnLetter          |  Returns column letter of input column number         |
'   | MakeDirString         |  Adds a "\" to a directory string if not found        |
'   | SplitText             |  Returns a string array of delimited values;          |
'   |                       |     removes extra spaces in splits                    |                                    |
'   | AddNewLine            |   Adds a new line (i.e. a hard return, or Chr(13)     |
'   | AddQuotes             |   Adds quotations around text                         |
'   | ReturnTextBetween     |   Returns string b/w starting & ending search strings |                                                   |
'   |                       |                                                       |
'   \-------------------------------------------------------------------------------/
'
'   /-------------------------------------------------------------------------------\
'   |   RETURN NUMERIC FUNCTIONS                                                    |
'   |-------------------------------------------------------------------------------|
'   |                           |                                                   |
'   | Log10                     |  Calculates log10 of any number                   |
'   | LogX                      |  Calculates the log of any base                   |
'   | LogXFactorial             |  Returns the factorial of any number in baseX     |
'   | SigFigs                   |  Return number with specified significant digits  |
'   |                           |      (needs work!)                                |
'   | StudentTText_EqualVar     |  Pairwise student's t-test for two normal         |
'   |                           |       distributions with equal variance           |
'   | StudentTText_UnequalVar   |  Pairwise student's t-test for two normal         |
'   |                           |       distributions with unequal variance         |
'   | FishersExactTest          |  Pairwise test for dichotomous data               |
'   | LinInterpolate            |  Linearly interpolates a value between given two  |
'   |                           |       distributions                               |
'   |                           |                                                   |
'   \-------------------------------------------------------------------------------/

'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|                                         |*
'*|   ADD EXCEL OBJECTS FUNCTION LIBRARY    |*
'*|               AJS LIBRARY               |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************
Function AddPictureToSheet(FN As String, ImageName As String, PasteRange As Range, WidthInches As Single, WidthHeight As Single) As Boolean
    '----------------------------------------------------------------
    ' AddPictureToSheet  - Adds a picture as a shape to a worksheet
    '                    - In : FN As String, ImageName As String, PasteRange As Range, WidthInches As Single, WidthHeight As Single
    '                    - Out: Boolean true if succesfully completed
    '                    - Last Updated: 5/31/11 by AJS
    '----------------------------------------------------------------
    Dim ThisShape As Shape
    On Error GoTo IsError
    Set ThisShape = PasteRange.Worksheet.Shapes.AddPicture(FN, msoFalse, msoTrue, _
                                                            PasteRange.Left, PasteRange.Top, _
                                                            Application.InchesToPoints(Width), Application.InchesToPoints(Height))
    ThisShape.Name = ImageName
    AddPictureToSheet = True
IsError:
    Debug.Print Err.Number & ": " & Err.Description
    AddPictureToSheet = False
End Function

Public Function AddNamedRange(NamedRange As Range, NamedRangeName As String, Optional WorkbookRange As Boolean = True) As Boolean
    '----------------------------------------------------------------
    ' AddNamedRange      - Add named range to Worbook or Worksheet
    '                    - In : NamedRange As Range, NamedRangeName As String, WorkbookRange As boolean [toggles workbook or worksheet range, by default true=workbook]
    '                    - Out: Boolean true/false if succesfully completed
    '                    - Last Updated: 3/6/11 by AJS
    '----------------------------------------------------------------
    On Error GoTo isErr
    If WorkbookRange = True Then
        ActiveWorkbook.Names.Add Name:=NamedRangeName, RefersTo:="='" & NamedRange.Worksheet.Name & "'!" & NamedRange.Address
    Else
        Sheets(NamedRange.Worksheet.Name).Names.Add Name:=NamedRangeName, RefersTo:="='" & NamedRange.Worksheet.Name & "'!" & NamedRange.Address
    End If
    On Error GoTo 0
    AddNamedRange = True
    Exit Function
isErr:
    AddNamedRange = False
    On Error GoTo 0
End Function

Public Function AddHyperlink(AnchorRange As Range, HyperlinkAddress As String, TextToDisplay As String) As Boolean
    '----------------------------------------------------------------
    ' AddHyperlink          - Adds a hyperlink to a cell range
    '                       - In : AnchorRange As Range, HyperlinkAddress As String, TextToDisplay As String
    '                       - Out: Boolean true if hyperlink succesfully added
    '                       - Last Updated: 3/6/11 by AJS
    '----------------------------------------------------------------
On Error GoTo IsError
    AnchorRange.Worksheet.Hyperlinks.Add Anchor:=AnchorRange, Address:=HyperlinkAddress, TextToDisplay:=TextToDisplay
    AddHyperlink = True
    On Error GoTo 0
    Exit Function
IsError:
    AddHyperlink = False
    On Error GoTo 0
End Function

Public Function AddStandardBorders(TableRange As Range) As Boolean
    '----------------------------------------------------------------
    ' AddStandardBorders - Adds standard thin line borders around range
    '                    - In : TableRange As Range
    '                    - Out: Boolean true if borders succesfully added
    '                    - Last Updated: 3/6/11 by AJS
    '----------------------------------------------------------------
    On Error GoTo isErr
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
    AddStandardBorders = True
    Exit Function
isErr:
    AddStandardBorders = False
End Function

Function AddDblOutSideBorder(TableRange As Range) As Boolean
    '----------------------------------------------------------------
    ' AddDblOutSideBorder - Adds double-line exterior border, and thin-line interior borders
    '                     - In : TableRange As Range
    '                     - Out: Boolean true if borders succesfully added
    '                     - Last Updated: 5/25/11 by AJS
    '----------------------------------------------------------------
    On Error GoTo isErr
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
    AddDblOutSideBorder = True
    Exit Function
isErr:
    AddDblOutSideBorder = False
End Function

Public Function AddValidationList(RangeToAddValidation As Range, NamedRangeName As Range, Optional InputTitle As String, Optional InputMessage As String) As Boolean
    '---------------------------------------------------------------------------------
    ' AddValidationList     - Adds a validation list to the selected cell
    '                       - In : RangeToAddValidation As Range, NamedRange As Range
    '                       - Out: Boolean true if validation succesfully added
    '                       - Created: 3/6/11 by AJS
    '                       - Modified: 6/1/11 by AJS
    '---------------------------------------------------------------------------------
    On Error Resume Next
    RangeToAddValidation.Validation.Delete
    On Error GoTo ValidationFailed
    With RangeToAddValidation.Validation
        .Add Type:=xlValidateList, _
                    AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, _
                    Formula1:="=" & NamedRangeName.Name.Name
        .InputMessage = InputMessage
        .InputTitle = InputTitle
    End With
    On Error GoTo 0
    AddValidationList = True
    Exit Function
ValidationFailed:
    AddValidationList = False
End Function

Function Validiation_DeleteAll(CellRange As Range) As Boolean
    '---------------------------------------------------------------------------------
    ' Validiation_DeleteAll - Deletes all validation in selected range
    '                       - In : CellRange As Range
    '                       - Out: Boolean true if validation succesfully deleted
    '                       - Last Updated: 5/2/11 by AJS
    '---------------------------------------------------------------------------------
    On Error GoTo IsError
    With CellRange.Validation
        .Delete
'        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = ""
'        .InputMessage = ""
'        .ShowInput = True
'        .ShowError = True
    End With
    Validiation_DeleteAll = True
    Exit Function
IsError:
    Validiation_DeleteAll = False
End Function

Function Validation_WholeNumber(CellRange As Range, Min As Long, Max As Long, InputTitle As String, InputInstructions As String) As Boolean
    '---------------------------------------------------------------------------------
    ' Validation_WholeNumber - Adds validation for any whole number within specified range
    '                        - In : CellRange As Range, Min As Long, Max As Long, InputTitle As String, InputInstructions As String
    '                        - Out: Boolean true if validation added
    '                        - Last Updated: 5/2/11 by AJS
    '---------------------------------------------------------------------------------
    On Error GoTo IsError
    With CellRange.Validation
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
    Validation_WholeNumber = False
End Function

Function Validation_FreeText(CellRange As Range, InputTitle As String, InputInstructions As String) As Boolean
    '---------------------------------------------------------------------------------
    ' Validation_FreeText    - Adds validation for any text, but includes input instructions
    '                        - In : CellRange As Range, InputTitle As String, InputInstructions As String
    '                        - Out: Boolean true if validation added
    '                        - Last Updated: 5/2/11 by AJS
    '---------------------------------------------------------------------------------
    On Error GoTo IsError
    With CellRange.Validation
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
    Validation_FreeText = False
End Function

Public Function AddCommentPicture(CommentCell As Range, PictureFN As String, Optional ScaleFactor As Double) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    ' AddCommentPicture  - Adds a picture into a comment for an Excel cell; deletes current comment
    '                    - In : Comment Cell as Range, PictureFN as string, Optional ScaleFactor as Double
    '                    - Out: Boolean true if picture comment succesfully added
    '                    - Note: Requires the "Frm_Image" to be in the workbook in order to determine image dimensions
    '                    - Last Updated: 4/7/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    If Len(Dir(PictureFN)) = 0 Then GoTo ReadError
    Frm_Image.Image1.Picture = LoadPicture(PictureFN)
    On Error Resume Next
    CommentCell.Comment.Delete
    On Error GoTo 0
    CommentCell.AddComment Text:=" "
    CommentCell.Comment.Visible = False
    CommentCell.Comment.Shape.Fill.UserPicture PictureFN
    If ScaleFactor = 0 Then ScaleFactor = 1
    CommentCell.Comment.Shape.Height = Frm_Image.Image1.Height * ScaleFactor
    CommentCell.Comment.Shape.Width = Frm_Image.Image1.Width * ScaleFactor
    AddCommentPicture = True
    Exit Function
ReadError:
    CommentCell.AddComment Text:="Image not found:" & vbNewLine & vbNewLine & PictureFN
    CommentCell.Comment.Visible = False
    AddCommentPicture = False
End Function

Public Function AddCommentText(CommentCell As Range, StringText As String, Optional CommentHeight As Integer = 100, Optional CommentWidth As Integer = 300) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    ' AddCommentText     - Adds a comment with the text specified by the user; deletes current comment
    '                    - In : CommentCell As Range, StringText As String, Optional CommentHeight As Integer = 100, Optional CommentWidth As Integer = 300
    '                    - Out: Boolean true if comment succesfully added
    '                    - Last Updated: 3/6/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    On Error Resume Next
    CommentCell.Comment.Delete
    On Error GoTo isErr
    CommentCell.AddComment StringText
    CommentCell.Comment.Visible = False
    CommentCell.Comment.Shape.Height = CommentHeight
    CommentCell.Comment.Shape.Width = CommentWidth
    AddCommentText = True
    On Error GoTo 0
    Exit Function
isErr:
    MsgBox "Comment Text failed to be added. " & vbNewLine & vbNewLine & _
        "Cell Address: " & CommentCell.Worksheet & "!" & CommentCell.Address & _
        "Text: " & StringText, vbCritical, "AddCommentText Function Failed"
    AddCommentText = False
    On Error GoTo 0
End Function

Function AddEmbededObject(FullFileName As String, SheetRange As Range, Optional NameInExcel As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    ' AddEmbededObject   - embed an object (such as a text file or picture) to a worksheet
    '                    - In : FullFilename As String, ExcelName As String, Optional SheetName As String, Optional SheetRange As String
    '                    - Out: Boolean true/false if succesfully completed
    '                    - Last Updated: 4/22/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    Dim OBJ As Variant
    On Error GoTo isErr
    If DoesFileExist(FullFileName) = True Then
        If SheetRange.Address <> "" Then
            SheetRange.Worksheet.Activate
            SheetRange.Select
        End If
        Set OBJ = ActiveSheet.OLEObjects.Add(FileName:=FullFileName, Link:=False, DisplayAsIcon:=False)
        If NameInExcel <> "" Then
            OBJ.Name = NameInExcel
        End If
        AddEmbededObject = True
        Exit Function
    End If
        GoTo isErr
isErr:
    AddEmbededObject = False
End Function

'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|                                         |*
'*|       RETURN BOOLEAN FUNCTIONS          |*
'*|             AJS LIBRARY                 |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************

'Public Function DoesFileExist(FullFileName As String) As Boolean
'    '---------------------------------------------------------------------------------------------------------
'    ' DoesFileExist      - Check to see if a file exists
'    '                    - In : FullFilename As String
'    '                    - Out: true if file exists, false if file does not exist
'    '                    - Last Updated: 3/6/11 by AJS
'    '---------------------------------------------------------------------------------------------------------
'    If Len(Dir(FullFileName)) > 0 Then
'            DoesFileExist = True
'        Else
'            DoesFileExist = False
'    End If
'End Function

Public Function IsWBOpen(WBName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------
    ' IsWBOpen           - Check to see if workbook is open
    '                    - In : WBName As String (include ".xls" extension)
    '                    - Out: true if worbook is open, false if workbook is not open
    '                    - Last Updated: 3/6/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim wBook As Workbook
    On Error Resume Next
    Set wBook = Workbooks(WBName)
    If wBook Is Nothing Then 'Not open
        Set wBook = Nothing
        IsWBOpen = False
    Else 'It is open
        IsWBOpen = True
    End If
    Set wBook = Nothing
    On Error GoTo 0
End Function

Public Sub ClearEmbeddedObjects()
    Application.ScreenUpdating = False
    Dim thisObj As Object
    Dim thisWS As Worksheet
    For Each thisWS In ThisWorkbook
        For Each thisObj In thisWS.OLEObjects
            thisObj.Delete
        Next
    Next
End Sub

'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|                                         |*
'*|        RETURN STRING FUNCTIONS          |*
'*|             AJS LIBRARY                 |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************

Public Function ColumnLetter(ColumnNumber As Variant) As String
    '---------------------------------------------------------------------------------------------------------
    ' ColumnLetter - Returns column letter of input column number, for up to 16348 columns
    ' - Tested 3/25/11 - significantly quicker than function ColumnLetter2; validated same results either way
	'    Alternate method:      ColumnLetter2 = Application.ConvertFormula("R1C" & ColumnNumber, xlR1C1, xlA1)
	' 							ColumnLetter2 = Mid(ColumnLetter, 2, Len(ColumnLetter) - 3)	 ' - Tested 6/30/11 - Select Case is slightly quicker than using a Nested If function
    ' - In : ColumnNumber As Integer
    ' - Out: ColumnLetter as String
    ' - Last Updated: 6/30/11 by AJS
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
    ColumnLetter = -999
End Function

Function MakeDirString(DirString As String) As String
    '---------------------------------------------------------------------------------------------------------
    ' MakeDirString      - Returns column letter of input column number
    '                    - In : DirString As String
    '                    - Out: MakeDirString as String
    '                    - Last Updated: 3/9/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    If Right(DirString, 1) <> "\" Then
        MakeDirString = DirString & "\"
    Else
        MakeDirString = DirString
    End If
End Function

Public Function SplitText(InTextLine As String, Delimeter As String) As Variant
    '---------------------------------------------------------------------------------------------------------
    ' SplitText          - Returns a string array of delimited values; removes extra spaces in splits
    '                    - In : InTextLine As String, Delimeter As String
    '                    - Out: SplitText as String()
    '                    - Last Updated: 3/9/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim k As Long, StringCount As Integer
    Dim TempString() As String
    Dim ThisChar As String, LastChar As String
    
    StringCount = 1
    ReDim TempString(1 To StringCount)
    LastChar = Delimeter
    
    For k = 1 To Len(InTextLine)
        ThisChar = Mid(InTextLine, k, 1)
        If ThisChar = Delimeter Then
            If LastChar <> Delimeter Then
                StringCount = StringCount + 1
                ReDim Preserve TempString(1 To StringCount)
                LastChar = ThisChar
            End If
        Else
            TempString(StringCount) = TempString(StringCount) & ThisChar
            LastChar = ThisChar
        End If
    Next k
    SplitText = TempString
End Function

Function SplitTextReturn(InTextLine As String, Delimeter As String, ReturnID As Integer) As String
    '---------------------------------------------------------------------------------------------------------
    ' SplitTextReturn    - Returns a field of a delimited text string
    '                    - In : InTextLine As String, Delimeter As String, ReturnID as Integer
    '                    - Out: SplitTextReturn as String
    '                    - Last Updated: 5/2/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim SplitString As Variant
    On Error GoTo isErr
    SplitString = SplitText(InTextLine, Delimeter)
    SplitTextReturn = SplitString(ReturnID)
    Exit Function
isErr:
    SplitTextReturn = Null
End Function

Public Function ReturnTextBetween(SearchText As String, StartField As String, EndField As String) As String
    '---------------------------------------------------------------------------------------------------------
    ' ReturnTextBetween  - Returns string between starting and ending search strings
    '                    - In : SearchText As String, StartField As String, EndField As String
    '                    - Out: ReturnTextBetween as String
    '                    - Last Updated: 3/9/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim CropLeft As String
    If InStr(1, SearchText, EndField, vbTextCompare) = 0 Then
        FindTextBetween = "ERROR- End field not found (" & """" & EndField & """" & " not not found in " & """" & SearchText & """" & ")"
        MsgBox FindTextBetween
    ElseIf InStr(1, SearchText, StartField, vbTextCompare) = 0 Then
        MsgBox FindTextBetween
        FindTextBetween = "ERROR- Start field not found (" & """" & StartField & """" & " not not found in " & """" & SearchText & """" & ")"
    Else
        CropLeft = Left(SearchText, InStr(1, SearchText, EndField, vbTextCompare) - 1)
        ReturnTextBetween = Right(CropLeft, Len(CropLeft) - (InStr(1, SearchText, StartField, vbTextCompare) + Len(StartField) - 1))
    End If
End Function

Public Function AddNewLine(Optional Repeat As Integer = 1) As String
    '----------------------------------------------------
    ' AddNewLine         - Prints a new line Chr(10), can be repeated
    '                    - In : <none>
    '                    - Out: Chr(10)
    '                    - Last Updated: 3/15/11 by AJS
    '----------------------------------------------------
     AddNewLine = WorksheetFunction.Rept(Chr(10), Repeat)
End Function

Public Function AddQuotes(ByVal TextInQuotes As String) As String
    '----------------------------------------------------
    '  AddNewLine         - Surrounds text in quotations
    '                    - In : TextInQuotes as String
    '                    - Out: "TextInQuotes" as String
    '                    - Last Updated: 3/6/11 by AJS
    '----------------------------------------------------
    AddQuotes = Chr(34) & TextInQuotes & Chr(34)
End Function
Public Function AddTab(Optional Repeat As Integer = 1) As String
    '----------------------------------------------------
    '  AddTab            - Adds a tab
    '                    - In : Repeat As Integer
    '                    - Out: Tabs in string
    '                    - Last Updated: 6/17/11 by AJS
    '----------------------------------------------------
    AddTab = WorksheetFunction.Rept(Chr(9), Repeat)
End Function
'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|                                         |*
'*|        RETURN NUMERIC FUNCTIONS         |*
'*|             AJS LIBRARY                 |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************

Public Function SigFigs(Value As Double, NumDigits As Integer) As String
    '---------------------------------------------------------------------------------------------------------
    ' SigFigs            - Returns the value with the specified significant digits
    '                    - In : Value As Double, NumDigits As Integer
    '                    - Out: SigFigs as String
    '                    - Last Updated: 3/9/11 by AJS
    '                    - Notes: close, but still needs to add zeroes if number is right of decimal, for example 0.0000100005, if sig figs = 3 then you would want to get 0.0000100
    '---------------------------------------------------------------------------------------------------------
    Dim MaxDigits As Integer
    On Error GoTo isErr
    If InStr(1, ".", SigFigs, vbTextCompare) = 0 Then
        MaxDigits = WorksheetFunction.Min(Len(CStr(Value)), NumDigits)
    Else
        MaxDigits = WorksheetFunction.Min(Len(CStr(Value)) - 1, NumDigits)
    End If
    SigFigs = WorksheetFunction.Round(Value, NumDigits - (1 + Int(WorksheetFunction.Log10(Abs(Value)))))
    If Len(SigFigs) >= MaxDigits Then Exit Function
    If InStr(1, ".", SigFigs, vbTextCompare) = 0 Then
        SigFigs = SigFigs & "."
    End If
    SigFigs = SigFigs & WorksheetFunction.Rept("0", MaxDigits - (Len(SigFigs) - 1))
    Exit Function
isErr:
    SigFigs = -999
End Function

Function LogX(Number As Double, Optional Base As Double = 10) As Double
    '---------------------------------------------------------------------------------------------------------
    '   LogX               - Converts a number to LogX form, Log10 by default
    '                      - In : Number as Double
    '                      - Out: LogX as Double
    '                      - Last Updated: 5/31/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
    LogX = Log(Number) / Log(Base)
    Exit Function
IsError:
    Debug.Print Err.Number & ": " & Err.Description
    LogX = CVErr(xlErrNA)
End Function

Function Log10(Number As Double) As Double
    '---------------------------------------------------------------------------------------------------------
    '   Log10              - Converts a number to Log10
    '                      - In : Number as Double
    '                      - Out: Log10 as Double
    '                      - Last Updated: 5/31/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
    LogX = Log(Number) / Log(10)
    Exit Function
IsError:
    Debug.Print Err.Number & ": " & Err.Description
    LogX = CVErr(xlErrNA)
End Function

Function StudentTText_UnequalVar(ByVal ControlMean As Double, ByVal ControlSD As Double, ByVal ControlN As Integer, _
                                 ByVal DoseMean As Double, ByVal DoseSD As Double, ByVal DoseN As Integer) As Double
    '----------------------------------------------------------------
    ' StudentTText_UnequalVar   - Calculates a student t-test with unequal sample size and
    '                             unequal variance, using a mean and standard deviation for
    '                             two distributions
    '                           - In : ControlMean As Double, ControlSD As Double, ControlN As Integer,
    '                                  DoseMean As Double, DoseSD As Double, DoseN As Integer
    '                           - Out: Double T-test p-value, or -999 if error
    '                           - Created On  : 5/24/11 by KEM
    '                           - Last Updated: 5/24/11 by AJS
    '----------------------------------------------------------------
    Dim Sx1x2, SQSx1x2, t, DFN, DFD, DF As Double
    On Error GoTo IsError
    Sx1x2 = ((ControlSD ^ 2) / ControlN) + ((DoseSD ^ 2) / DoseN)
    SQSx1x2 = Sqr(Sx1x2)
    t = Abs((ControlMean - DoseMean)) / SQSx1x2
    DFN = (Sx1x2) ^ 2
    DFD = (((ControlSD ^ 2 / ControlN) ^ 2) / (ControlN - 1)) + (((DoseSD ^ 2 / DoseN) ^ 2) / (DoseN - 1))
    DF = DFN / DFD
    StudentTText_UnequalVar = Application.TDist(t, DF, 2)
    Exit Function
IsError:
    Debug.Print Err.Number & ": " & Err.Description
    StudentTText_UnequalVar = -999
End Function

Function StudentTText_EqualVar(ByVal ControlMean As Double, ByVal ControlSD As Double, ByVal ControlN As Integer, _
                               ByVal DoseMean As Double, ByVal DoseSD As Double, ByVal DoseN As Integer) As Double
    '----------------------------------------------------------------
    ' StudentTText_UnequalVar   - Calculates a student t-test with equal sample size and
    '                             equal variance, using a mean and standard deviation for
    '                             two distributions
    '                           - In : ControlMean As Double, ControlSD As Double, ControlN As Integer,
    '                                  DoseMean As Double, DoseSD As Double, DoseN As Integer
    '                           - Out: Double T-test p-value, or -999 if error
    '                           - Created On  : 5/24/11 by KEM
    '                           - Last Updated: 5/24/11 by AJS
    '----------------------------------------------------------------
    Dim Sx1x2N, Sx1x2D, Sx1x2, t, DF As Double
    On Error GoTo IsError
    Sx1x2N = ((ControlN - 1) * ControlSD ^ 2) + ((DoseN - 1) * DoseSD ^ 2)
    Sx1x2D = ControlN + DoseN - 2
    Sx1x2 = Sqr(Sx1x2N / Sx1x2D)
    t = Abs(ControlMean - DoseMean) / (Sx1x2 * Sqr((1 / ControlN) + (1 / DoseN)))
    DF = ControlN + DoseN - 2
    StudentTText_EqualVar = Application.TDist(t, DF, 2)
    Exit Function
IsError:
    Debug.Print Err.Number & ": " & Err.Description
    StudentTText_EqualVar = -999
End Function

Function FishersExactText(ByVal A1 As Long, ByVal B1 As Long, ByVal A2 As Long, ByVal B2 As Long) As Double
    '---------------------------------------------------------------------------------------------------------
    '   FishersExactTest   - Calculates a pair-wise significance test using fisher's exact test method
    '                        Modified to calculate in log-space to allow for much larger matrices (tested w/ integer value of 12,000+)
    '                        Adapted from http://mathworld.wolfram.com/FishersExactTest.html
    '                      - In : A1 As Long, B1 As Long, A2 As Long, B2 As Long
    '                      - Out: FishersExactTest as Double
    '                      - Last Updated: 5/31/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim LogMatrix(1 To 3, 1 To 3) As Double
    LogMatrix(1, 1) = LogXFactorial(A1, 10)
    LogMatrix(1, 2) = LogXFactorial(B1, 10)
    LogMatrix(1, 3) = LogXFactorial(A1 + B1, 10)
    LogMatrix(2, 1) = LogXFactorial(A2, 10)
    LogMatrix(2, 2) = LogXFactorial(B2, 10)
    LogMatrix(2, 3) = LogXFactorial(A2 + B2, 10)
    LogMatrix(3, 1) = LogXFactorial(A1 + A2, 10)
    LogMatrix(3, 2) = LogXFactorial(B1 + B2, 10)
    LogMatrix(3, 3) = LogXFactorial(A1 + A2 + B1 + B2, 10)
    ' added/subtracted rather than multiplied/divided becase in logspace
    FishersExactText = 10 ^ (LogMatrix(1, 3) + LogMatrix(2, 3) + (LogMatrix(3, 1) + LogMatrix(3, 2)) - _
                    (LogMatrix(3, 3) + LogMatrix(1, 1) + LogMatrix(1, 2) + LogMatrix(2, 1) + LogMatrix(2, 2)))
End Function

Function LogXFactorial(ByVal Value As Long, Optional Base As Integer = 10) As Double
    '---------------------------------------------------------------------------------------------------------
    '   LogXFactorial     - Calculates the factorial of any number, using log10 space by default
    '                      - Returns the result in the specified log
    '                      - In : Value as Long, Optional Base as Integer = 10
    '                      - Out: LogXFactorial as Double
    '                      - Last Updated: 6/15/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim i As Long
    LogXFactorial = 0
    For i = 1 To Value
        LogXFactorial = LogXFactorial + LogX(CDbl(i), CDbl(Base))
    Next i
End Function

Public Function LinInterpolate(XValue As Range, XRange As Range, YRange As Range) As String
    '----------------------------------------------------------------
    ' LinInterpolate        - Linearly interpolates between two ranges of values
    '                       - In : ByVal XValue As String, XRange As Range, YRange As Range
    '                       - Out: Linear interpolation as string, may include < or > if greater than bounds of range
    '                       - Last Updated: 3/24/11 by AJS
    '----------------------------------------------------------------
    Dim sncell As Range, XValueDbl As Double
    Dim X1 As Double, X2 As Double, Y1 As Double, Y2 As Double
    On Error GoTo IsError
    ' error checking
    If IsNumeric(XValue) = False Then
        GoTo IsError
    Else
        XValueDbl = CDbl(XValue)
    End If
    If XRange.Columns.Count <> 1 Then
        MsgBox "Error- XRange should only be one column"
        Exit Function
    End If
    If YRange.Columns.Count <> 1 Then
        MsgBox "Error- YRange should only be one column"
        Exit Function
    End If
    If XRange.Cells.Count <> YRange.Cells.Count Then
        MsgBox "Error- XRange does not have the same rows as YRange"
        Exit Function
    End If
    If XRange.Cells(1).Row <> YRange.Cells(1).Row Then
        MsgBox "Error- XRange and YRange must start on the same row"
        Exit Function
    End If
    If XValueDbl < WorksheetFunction.Min(XRange) Then
        LinInterpolate = "<" & WorksheetFunction.Min(XRange)
        Exit Function
    End If
    If XValueDbl > WorksheetFunction.Max(XRange) Then
        LinInterpolate = ">" & WorksheetFunction.Max(XRange)
        Exit Function
    End If
    If FindMatch(XValue.Value, XRange) > 0 Then
        LinInterpolate = YRange(FindMatch(XValue.Value, XRange))
        Exit Function
    End If
    For Each sncell In XRange
        If IsNumeric(sncell.Value) Then
            If XValueDbl < sncell.Value Then
                X1 = Sheets(YRange.Worksheet.Name).Cells(sncell.Row - 1, XRange.Column)
                X2 = Sheets(YRange.Worksheet.Name).Cells(sncell.Row, XRange.Column)
                Y1 = Sheets(YRange.Worksheet.Name).Cells(sncell.Row - 1, YRange.Column)
                Y2 = Sheets(YRange.Worksheet.Name).Cells(sncell.Row, YRange.Column)
                LinInterpolate = Y1 + (Y2 - Y1) * ((XValueDbl - X1) / (X2 - X1))
                Exit Function
            End If
        End If
    Next
        LinInterpolate = "-"
    Exit Function
IsError:
    LinInterpolate = CVErr(xlErrNA)
End Function

'*********************************************
'*/-----------------------------------------\*
'*|                                         |*
'*|                                         |*
'*|        RETURN RANGE FUNCTIONS           |*
'*|             AJS LIBRARY                 |*
'*|                                         |*
'*\-----------------------------------------/*
'*********************************************
Public Function ExtTbl(Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0) As Range
    '---------------------------------------------------------------------------------------------------------
    ' ExtTbl             - Entends the table down to the first blank at bottom of top right row/column
    '                           will stop at the first blank row
    '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
    '                    - Out: ExtTbl as Range
    '                    - Last Updated: 4/7/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    On Error GoTo isErr
    Set ExtTbl = ExtRight(ExtDown(Rng.Offset(RowOffset, ColOffset), 0, 0), 0, 0)
    Exit Function
isErr:
    Set ExtTbl = Rng
End Function

Public Function ExtDown(Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0) As Range
    '---------------------------------------------------------------------------------------------------------
    ' ExtDown            - Extends the selected range down to the final non-blank row in current table;
    '                           will stop at the first blank row
    '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
    '                    - Out: ExtDown as Range
    '                    - Last Updated: 4/7/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    On Error GoTo isErr
    Set Rng = Rng.Offset(RowOffset, ColOffset)
    If IsEmpty(Rng.Offset(1, 0)) Then
        Set ExtDown = Rng
    Else
        Set ExtDown = Range(Rng, Rng.End(xlDown))
    End If
    Exit Function
isErr:
    ExtDown = Rng
End Function

Public Function ExtRight(Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0) As Range
    '---------------------------------------------------------------------------------------------------------
    ' ExtRight           - Extends the selected range down to the final non-blank column in current table;
    '                           will stop at the first blank column
    '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
    '                    - Out: ExtRight as Range
    '                    - Last Updated: 4/7/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    On Error GoTo isErr
    Set Rng = Rng.Offset(RowOffset, ColOffset)
    If IsEmpty(Rng.Offset(0, 1)) Then
        Set ExtRight = Rng
    Else
        Set ExtRight = Range(Rng, Rng.End(xlToRight))
    End If
    Exit Function
isErr:
    ExtRight = Rng
End Function

Public Function ExtDownNonBlank(Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0) As Range
    '---------------------------------------------------------------------------------------------------------
    ' ExtDownNonBlank    - Extends the range down to the first non-blank at bottom-left of selected range;
    '                           will stop at the first blank row where a formula evaluates to a value
    '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
    '                    - Out: ExtDownNonBlank as Range
    '                    - Created: 5/15/11 by GH
    '                    - Last Updated: 6/1/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Application.Volatile True
    Dim NewRng As Range
    Dim LastRow As Long
    Set NewRng = ExtDown(Rng, RowOffset, ColOffset)
    For LastRow = NewRng.Rows.Count To 1 Step -1
        'check if blank
        If Not NewRng.Cells(LastRow, 1) = "" Then
            GoTo exitloop:
        End If
    Next LastRow
exitloop:
    Set ExtDownNonBlank = NewRng.Resize(LastRow, NewRng.Columns.Count)
End Function

Public Function ExtTblNonBlank(Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0) As Range
    '---------------------------------------------------------------------------------------------------------
    ' ExtTblNonBlank     - Extends the range down and right to the first non-blank at bottom-left of selected range;
    '                           will stop at the first blank row where a formula evaluates to a value
    '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
    '                    - Out: ExtTblNonBlank as Range
    '                    - Created: 5/15/11 by GH
    '                    - Last Updated: 6/1/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Application.Volatile (True)
    On Error GoTo isErr
    Set ExtTblNonBlank = ExtRight(ExtDownNonBlank(Rng.Offset(RowOffset, ColOffset), 0, 0), 0, 0)
    Exit Function
isErr:
    Set ExtTblNonBlank = Rng
End Function

Public Function ExtAllTbl(ByRef Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0) As Range
    '---------------------------------------------------------------------------------------------------------
    ' ExtAllTbl          - Extends the selected range right and down to the final non-blank row if
    '                           leftmost column and the final non-blank row in topmost row
    '                    - In : ByRef Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
    '                    - Out: ExtAllTbl as Range
    '                    - Last Updated: 3/9/11 by AJS (originally from GH)
    '---------------------------------------------------------------------------------------------------------
    Dim RightmostColumn As Long
    Dim BottomRow As Long
    Set Rng = Rng.Offset(RowOffset, ColOffset)
    BottomRow = Application.Max(LastRow(Rng.Worksheet, Rng.Column), Rng.Row)
    RightmostColumn = Application.Max(LastColumn(Rng.Worksheet, Rng.Row), Rng.Column)
    Set ExtAllTbl = Rng.Resize(rowsize:=(BottomRow - Rng.Rows.Count - Rng.Row + 2), ColumnSize:=(RightmostColumn - Rng.Columns.Count - Rng.Column + 2))
End Function

Public Function ExtAllDown(ByRef Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0) As Range
    '---------------------------------------------------------------------------------------------------------
    ' ExtAllDown         - Extends the selected range down to the final non-blank row in leftmost column
    '                    - In : ByRef Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
    '                    - Out: ExtAllDown as Range
    '                    - Last Updated: 3/9/11 by AJS (originally from GH)
    '---------------------------------------------------------------------------------------------------------
    Dim BottomRow As Long
    Set Rng = Rng.Offset(RowOffset, ColOffset)
    BottomRow = Application.Max(LastRow(Rng.Worksheet, Rng.Column), Rng.Row)
    Set ExtAllDown = Rng.Resize(rowsize:=(BottomRow - Rng.Rows.Count - Rng.Row + 2))
End Function

Function ExtAllRight(ByRef Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0) As Range
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
End Function

Public Function LastRow(ByVal OfSheet As Worksheet, Optional ByVal InColumn As Long = 0) As Long
    '---------------------------------------------------------------------------------------------------------
    ' LastRow            - Returns the number of the last used row in the specified sheet [and column]
    '                    - In : ByVal OfSheet As Worksheet, Optional ByVal InColumn As Long = 0
    '                    - Out: LastRow as Range
    '                    - Last Updated: 3/9/11 by AJS (originally from GH)
    '---------------------------------------------------------------------------------------------------------
    If InColumn = 0 Then
        LastRow = OfSheet.UsedRange.Row + OfSheet.UsedRange.Rows.Count - 1
    Else
        LastRow = OfSheet.Cells(Application.Rows.Count, InColumn).End(xlUp).Row
    End If
End Function

Public Function LastColumn(ByVal OfSheet As Worksheet, Optional ByVal InRow As Long = 0) As Integer
    '---------------------------------------------------------------------------------------------------------
    ' LastColumn         - Returns the number of the last used column in the specified sheet [and row]
    '                    - In : ByVal OfSheet As Worksheet, Optional ByVal InRow As Long = 0
    '                    - Out: LastColumn as Range
    '                    - Last Updated: 3/9/11 by AJS (originally from GH)
    '---------------------------------------------------------------------------------------------------------
    Dim i As Integer, letter As String
    If InRow = 0 Then
        i = OfSheet.UsedRange.Columns.Count + 1
        Do
            i = i - 1
            LastColumn = OfSheet.UsedRange.Columns(i).Cells(1, 1).Column
            letter = ColumnLetter(LastColumn)
        Loop Until (Application.WorksheetFunction.CountA(OfSheet.Range(letter & ":" & letter)) > 0 Or i < 2)
    Else
        LastColumn = OfSheet.Cells(InRow, Application.Columns.Count).End(xlToLeft).Column
    End If
End Function