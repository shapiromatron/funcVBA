' ARCHIVE OF OLD FUNCTIONS

Public Function ColumnLetter(ColumnNumber As Variant) As String
	 '---------------------------------------------------------------------------------------------------------
	 ' ColumnLetter - Returns column letter of input column number, for up to 16348 columns
	 ' - Tested 3/25/11 - significantly quicker than function ColumnLetter2; validated same results either way
	 ' - In : ColumnNumber As Integer
	 ' - Out: ColumnLetter as String
	 ' - Last Updated: 3/25/11 by AJS
	 '---------------------------------------------------------------------------------------------------------
	 On Error GoTo isErr
	 If ColumnNumber > 1378 Then 'special case, the first 26 column set should be subtracted , 26*26 = 676
		ColumnLetter = Chr(Int((ColumnNumber - 26 - 1) / 676) + 64) & Chr(Int(((ColumnNumber - 1 - 26) Mod 676) / 26) + 65) & Chr(((ColumnNumber - 1) Mod 26) + 65)
	 ElseIf ColumnNumber > 702 Then 'includes first column, 26*26 + 26=702
		ColumnLetter = Chr(Int(ColumnNumber / 702) + 64) & Chr(Int(((ColumnNumber - 1) Mod 702) / 26) + 65) & Chr(((ColumnNumber - 1) Mod 26) + 65)
	 ElseIf ColumnNumber > 26 Then
		ColumnLetter = Chr(Int((ColumnNumber - 1) / 26) + 64) & Chr(((ColumnNumber - 1) Mod 26) + 65)
	 Else
		ColumnLetter = Chr(ColumnNumber + 64)
	 End If
    Exit Function
isErr:
    ColumnLetter = -999
End Function

Function ColumnLetter2(ColumnNumber As Variant) As String
    '---------------------------------------------------------------------------------------------------------
    ' ColumnLetter2      - Returns column letter of input column number
    '                    -    Tested 3/25/11 - significantly slower than function ColumnLetter; validated same results either way
    '                    - In : ColumnNumber As Integer
    '                    - Out: ColumnLetter as String
    '                    - Last Updated: 3/25/11 by AJS
    '---------------------------------------------------------------------------------------------------------
On Error GoTo IsError:
    ColumnLetter2 = Application.ConvertFormula("R1C" & ColumnNumber, xlR1C1, xlA1)
    ColumnLetter2 = Mid(ColumnLetter, 2, Len(ColumnLetter) - 3)
    Exit Function
IsError:
    ColumnLetter2 = -999
End Function
