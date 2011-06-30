Option Explicit

Public Function ExtTbl(Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0) As Range
    '---------------------------------------------------------------------------------------------------------
    ' ExtTbl             - Entends the table down to the first blank at bottom of top right row/column
    '                           will stop at the first blank row
    '                    - In : Rng As Range, Optional RowOffset As Long = 0, Optional ColOffset As Long = 0
    '                    - Out: ExtTbl as Range
    '                    - Last Updated: 4/7/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Application.Volatile (True)
    On Error GoTo IsErr
    Set ExtTbl = ExtRight(ExtDown(Rng.OffSet(RowOffset, ColOffset), 0, 0), 0, 0)
    Exit Function
IsErr:
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
    Application.Volatile (True)
    On Error GoTo IsErr
    Set Rng = Rng.OffSet(RowOffset, ColOffset)
    If IsEmpty(Rng.OffSet(1, 0)) Then
        Set ExtDown = Rng
    Else
        Set ExtDown = Range(Rng, Rng.End(xlDown))
    End If
    Exit Function
IsErr:
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
    Application.Volatile (True)
    On Error GoTo IsErr
    Set Rng = Rng.OffSet(RowOffset, ColOffset)
    If IsEmpty(Rng.OffSet(0, 1)) Then
        Set ExtRight = Rng
    Else
        Set ExtRight = Range(Rng, Rng.End(xlToRight))
    End If
    Exit Function
IsErr:
    ExtRight = Rng
End Function

