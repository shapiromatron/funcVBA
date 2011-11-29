Attribute VB_Name = "z_Numeric"
Option Explicit

Enum NumberPrintReturnType
    '---------------------------
    ' Used with NumberToPrint
    '---------------------------
    ReturnNumber = 1
    ReturnFormat = 2
End Enum

Function NumberToPrint(Number As Variant, ReturnType As NumberPrintReturnType, ShowCommas As Boolean) As Variant
    '---------------------------------------------------------------------------------------------------------
    ' NumberToPrint - Returns the number formatted or the format type
    '               - In :  Number as Variant
    '                       ReturnType As NumberPrintReturnType
    '               - Out: Number or number format if numeric, text or text format if not numeric, error if other
    '               - Last Updated: 7/4/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim NumAfterDecimal As Double
    Dim ConvertedText As String
    Dim ShowDec As String
    Dim TextFormat As String
    
    If IsNumeric(Number) = False Then GoTo NonNumeric
        
    On Error GoTo IsError
    
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
            TextFormat = "0"
        Case Is > 100000
            TextFormat = "0.0E+00"
        Case Is > 10000
            If ShowCommas = True Then
               TextFormat = "0,000"
            Else
                TextFormat = "0000"
            End If
        Case Is > 1000  'and < 10,000
            If ShowCommas = True Then
               TextFormat = "0,000"
            Else
                TextFormat = "0000"
            End If
        Case Is > 100   'and < 1,000
            TextFormat = "0"
        Case Is > 10    'and < 100
            TextFormat = "0" & ShowDec & WorksheetFunction.Rept("0", WorksheetFunction.Min(1, NumAfterDecimal))
        Case Is > 1     'and < 10
            TextFormat = "0" & ShowDec & WorksheetFunction.Rept("0", WorksheetFunction.Min(2, NumAfterDecimal))
        Case Is > 0.1   'and < 1
            TextFormat = "0." & WorksheetFunction.Rept("0", WorksheetFunction.Min(3, NumAfterDecimal))
        Case Is > 0.01  'and < 0.1
            TextFormat = "0." & WorksheetFunction.Rept("0", WorksheetFunction.Min(4, NumAfterDecimal))
        Case Is > 0.001 'and < 0.01
            TextFormat = "0." & WorksheetFunction.Rept("0", WorksheetFunction.Min(5, NumAfterDecimal))
        Case Else   ' <= 0.001
            TextFormat = "0.00E-00"
    End Select
    Select Case ReturnType
        Case ReturnFormat
            NumberToPrint = TextFormat
        Case ReturnNumber
            NumberToPrint = Format(Number, TextFormat)
    End Select
    Exit Function
NonNumeric:
    Select Case ReturnType
        Case ReturnFormat
            NumberToPrint = "@"
        Case ReturnNumber
            NumberToPrint = Format(Number, "@")
    End Select
    Exit Function
IsError:
    NumberToPrint = CVErr(xlErrNA)
    Debug.Print "Error in NumberToPrint: " & Err.Number & ": " & Err.Description
End Function

Function LogX(Number As Double, Optional Base As Double = 10) As Variant
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
    Debug.Print "Error in LogX: " & Err.Number & ": " & Err.Description
    LogX = CVErr(xlErrNA)
End Function

Function Log10(Number As Double) As Variant
    '---------------------------------------------------------------------------------------------------------
    '   Log10              - Converts a number to Log10
    '                      - In : Number as Double
    '                      - Out: Log10 as Double
    '                      - Last Updated: 5/31/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    On Error GoTo IsError
    Log10 = Log(Number) / Log(10)
    Exit Function
IsError:
    Debug.Print "Error in Log10: " & Err.Number & ": " & Err.Description
    Log10 = CVErr(xlErrNA)
End Function

Function StudentTText_UnequalVar(ByVal ControlMean As Double, _
                                    ByVal ControlSD As Double, _
                                    ByVal ControlN As Integer, _
                                    ByVal DoseMean As Double, _
                                    ByVal DoseSD As Double, _
                                    ByVal DoseN As Integer) As Double
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
    Debug.Print "Error in StudentTText_UnequalVar: " & Err.Number & ": " & Err.Description
    StudentTText_UnequalVar = -999
End Function

Function StudentTText_EqualVar(ByVal ControlMean As Double, _
                                ByVal ControlSD As Double, _
                                ByVal ControlN As Integer, _
                                ByVal DoseMean As Double, _
                                ByVal DoseSD As Double, _
                                ByVal DoseN As Integer) As Double
    '----------------------------------------------------------------
    ' StudentTText_EqualVar     - Calculates a student t-test with equal sample size and
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
    Debug.Print "Error in StudentTText_EqualVar: " & Err.Number & ": " & Err.Description
    StudentTText_EqualVar = -999
End Function

Function FishersExactText(ByVal A1 As Long, ByVal B1 As Long, _
                            ByVal A2 As Long, ByVal B2 As Long) As Variant
    '---------------------------------------------------------------------------------------------------------
    '   FishersExactTest   - Calculates a pair-wise significance test using fisher's exact test method
    '                        Modified to calculate in log-space to allow for much larger matrices (tested w/ integer value of 12,000+)
    '                        Adapted from http://mathworld.wolfram.com/FishersExactTest.html
    '                      - In : A1 As Long, B1 As Long, A2 As Long, B2 As Long
    '                      - Out: FishersExactTest as Double
    '                      - Last Updated: 5/31/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim LogMatrix(1 To 3, 1 To 3) As Double
    On Error GoTo IsError
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
    Exit Function
IsError:
    Debug.Print "Error in FishersExactText: " & Err.Number & ": " & Err.Description
    FishersExactText = CVErr(xlErrNA)
End Function

Function LogXFactorial(ByVal Value As Long, Optional Base As Integer = 10) As Variant
    '---------------------------------------------------------------------------------------------------------
    '   LogXFactorial     - Calculates the factorial of any number, using log10 space by default
    '                      - Returns the result in the specified log
    '                      - In : Value as Long, Optional Base as Integer = 10
    '                      - Out: LogXFactorial as Double
    '                      - Last Updated: 6/15/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo IsError
    LogXFactorial = 0
    For i = 1 To Value
        LogXFactorial = LogXFactorial + LogX(CDbl(i), CDbl(Base))
    Next i
    Exit Function
IsError:
    Debug.Print "Error in LogXFactorial: " & Err.Number & ": " & Err.Description
    LogXFactorial = CVErr(xlErrNA)
End Function

Public Function LinInterpolate(XValue As Range, XRange As Range, YRange As Range) As Variant
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
    Debug.Print "Error in LinInterpolate: " & Err.Number & ": " & Err.Description
    LinInterpolate = CVErr(xlErrNA)
End Function


Function SigFig(Value As Double, SigFigs As Integer) As String
    '----------------------------------------------------------------
    ' SigFig        - Returns a string with the specified number of significant digits
    '                 http://excel.tips.net/T001983_Thoughts_and_Ideas_on_Significant_Digits_in_Excel.html
    '               - In : Value As Double, SigFigs As Integer
    '               - Out: Value as string with specified significant digits
    '               - Last Updated: 11/28/11 by AJS
    '               - Things to add:
    '                   a) Specify general or scientific format
    '                   b) Don't show more sig figs than actually exist
    '----------------------------------------------------------------
    Dim val As String
    val = WorksheetFunction.Fixed(Value, SigFigs - Int(WorksheetFunction.Log10(Value)) - 1)
    SigFig = val
End Function
