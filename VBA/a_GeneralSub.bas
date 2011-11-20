Option Explicit


'EXAMPLE SETUP FOR DEFINING A WORKSHEET
Global Const SN_Main As String = "Company"
Public Enum SMain
    R_Start = 13
    C_Start = 2
    C_FN = 3
End Enum

'
'   /-------------------------------------------------------------------------------\
'   |   GENERAL COMPUTATIONAL SUBS                                                  |
'   |-------------------------------------------------------------------------------|
'   |                       |                                                       |
'   | AppWait               |   Makes the workbook wait for a period of time        |
'   | SortRange             |   Sorts the selected range, w/ or w/o header          |
'   |                       |                                                       |
'   \-------------------------------------------------------------------------------/
'

'Evaluate an array function in VBA
InputTitle = ActiveSheet.Evaluate("=Index(Tbl_ValidationInputNotes, Match(" & _
                         Chr(34) & eachTarget.Name.Name & eachTarget.Value & Chr(34) & _
                         ",C_ValidInputNotes_NameRange & C_ValidInputNotes_OutputValue, False), 4)")

'Wait for a minimum of one second
Public Sub AppWait(Hrs As Double, Min As Double, Sec As Double)
    '-----------------------------------------------------------------------------------------------------------
    ' AppWait            - Makes the workbook wait for a period of time
    '                    - In : Hrs As Double, Min As Double, Sec As Double
    '                    - Last Updated: 3/9/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    Application.Wait Now + TimeValue(Hrs & ":" & Min & ":" & Sec)
End Sub

'Example commonly used coding terms:
Sub ExampleCode()

    'SCREEN UPDATING
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True
    
    'AUTOMATIC CALCULATION FOR SPREADSHEETS
    Application.Calculation = xlManual
    Application.Calculation = xlAutomatic
    
    'STATUS BAR AT BOTTOM
    Application.StatusBar = "Add text here..."
    Application.StatusBar = False
    
    'DISPLAY ALERTS
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    
    'COPY AND PASTE (VALUES AND FORMATTING)
    Set CopyRange = Sheets("Inputs").Range("A1:D5")
    Set PasteRange = Sheets("Outputs").Range("A1:D5")
    CopyRange.Copy PasteRange           '(paste all)
    PasteRange.Value = CopyRange.Value  '(values only)
                
    'For loop examples
    For i = LBound(StringArray) To UBound(StringArray)  'to the lower/upper bound of an array
    Next i
    For i = 1 To 10 Step 2  'every 2
    Next i
    For i = 10 To 1 Step -1 'backwards
    Next i
        
    'SEE IF VALUE FOUND IN RANGE
    If IsError(Application.Match(SearchString, Range("SearchList"), False)) Then
        'value not found
    Else
        'value found
    End If
    
    'HOW TO HIDE A GROUP BOX USING FORM CONTROLS
    Sheet1.Shapes("Group Box 3").Visible = True
    
    'INPUT TEXT FILES LINE BY LINE
    Open InFileDir & InFileName For Input As #1
         Do Until EOF(1)
		     Line Input #1, StringVariable
             'split comma delimited
             StringVariable() = Split(StringVariable, ",")
             'remove excess spaces from each side of string
             For j = 0 To UBound(StringVariable())
                 StringVariable(j) = Trim(StringVariable(j))
             Next j
             'test to see if something exists in the code
             If InStr(1, "Example string to search", "string") > 0 Then
                 MsgBox "string found"
             Else
                 MsgBox "string not found"
             End If
         Loop
     Close #1
	 
	'INPUT FULL TEXT FILE INTO ONE STRING
	Open OutFN For Input As #1
		OutText = Input(LOF(1), #1)
		Sheets(SN_Output).Cells(OutRow, C_Output_OutText) = OutText
	Close #1
    
    'OUTPUT TEXT FILES
    Open InFileDir & InFileName For Output As #2
        Print #2, StringVariable    'print line; go to next line
        Print #2, StringVariable;   'print line; don't go to next line
        Write #2, StringVariable    'Writes a series of values, with text enclosed in quotes, separated by commas
    Close #2
    
    'CHECKS FOR KEYSTROKES ON ANYTHING ELSE WHEN WORKING
    DoEvents
    
    'CONVERTS FROM R1C1 TO A1 RANGE
    x = Application.ConvertFormula("R1C1", xlR1C1, xlA1)
    
    'CONVERTS FROM RANGE to R1C1
    y = Application.ConvertFormula("$A$1:$C$10", xlA1, xlR1C1) ' must include $ signs!
    
    'File scripting objects:
    'http://msdn.microsoft.com/en-us/library/2z9ffy99(v=vs.85).aspx#Y800
    
    'MAXIMUM NON-BLANK ROW: array formula (ctrl-shift-enter)
        '=MAX(IF(NOT(ISBLANK(RowRange)),ROW(RowRange),"N/A"))
        '=MAX(IF(RowRange="SearchTerm",ROW(RowRange),"N/A"))
    
    'MINIMUM NON-BLANK ROW: array formula (ctrl-shift-enter)
        '=MIN(IF(NOT(ISBLANK(RowRange)),ROW(RowRange),"N/A"))
        '=MIN(IF(RowRange="SearchTerm",ROW(RowRange),"N/A"))
    
    'VLOOKUP USING MULTIPLE SEARCH CRITERIA
        '=INDEX($B$1:$F$20,MATCH(A1&"-"&A2,$B$1:$B$20&"-"&$C$1:$C$20,2))
        'http://support.microsoft.com/kb/214142
        
    'FOR EACH CELL IN RANGE
    Dim Rng As Range, RngCell As Range
    Set Rng = Sheet1.Range("A1:A50")
    For Each RngCell In Rng
        RngCell.Value
    Next
    Set RngCell = Nothing
    
    'FOR EACH WORKSHEET
    Dim WS As Worksheet
    Application.DisplayAlerts = False
    For Each WS In ThisWorkbook.Worksheets
        If WS.Name = "Delete Me" Then
            WS.Delete
        End If
    Next
    Application.DisplayAlerts = True
    Set wsSheet = Nothing
    
    'FOR EACH NAMED RANGE
    Dim RangeName
    Dim WS As Worksheet
    'workbook named ranges
    For Each RangeName In ThisWorkbook.Names
        If InStr(1, RangeName.RefersTo, "#REF!", vbTextCompare) > 0 Then
            RangeName.Delete
        End If
    Next
    'worksheet named ranges
    For Each WS In ThisWorkbook.Worksheets
        For Each RangeName In WS.Names
            If InStr(1, RangeName.RefersTo, "#REF!", vbTextCompare) > 0 Then
                RangeName.Delete
            End If
        Next
    Next
End Sub

Private Sub SearchWithinDirectory()
'DOESN'T WORK IN EXCEL 2007, ONLY EXCEL 2003
    'search with a folder
    With Application.FileSearch
        .NewSearch
        .LookIn = PDF_dir
        .SearchSubFolders = False
        .FileName = "*.pdf" 'wildcard
        .Execute
        'now do something for each file found
        For i = 1 To .FoundFiles.Count
        Next i
    End With
End Sub

Sub OpenWorkbook()
    Dim OtherWB As Workbook
    
    'opens current WB
    Workbooks.Open FileName:=WBDir & WBName
    Set OtherWB = Workbooks(WBName)
        
        'run a macro
        OtherWB.Activate
        Application.Run ("'Book123.xlsm'!MacroName")
        Application.Run ("'" & OtherWB.Name & "'!MacroName")
        
        'copy from one workbook to this workbook
        Set CopyRange = OtherWB.Sheets("Results").Range("B1:C5")
        Set PasteRange = ThisWorkbook.Sheets("Results").Range("B1:C5")
        PasteRange.Value = CopyRange.Value 'values only
        CopyRange.Copy PasteRange   'paste everything
        
    'close workbook and save changes
    OtherWB.Close SaveChanges:=True
End Sub

'DICTIONARY BASICS
Sub DictBasics()
    Dim Dict As Variant, eachkey As Variant

    'Creation
    Set Dict = CreateObject("Scripting.Dictionary")

    'Addition
    Dict.Add "Key", "Value"
    Dict.Add "Key2", "Value2"
    
    'Update
    Dict("Key2") = "Value3"

    'Retrieval
    For Each eachkey In Dict
        Debug.Print "Key: " & eachkey
        Debug.Print "Value: " & Dict(eachkey)
    Next
    
    'Reset
    Dict.RemoveAll
End Sub

