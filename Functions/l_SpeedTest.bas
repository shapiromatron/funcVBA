'|--------------------------------------------------------|
'| VERY PRECISE TIMER TO USE FOR DEBUGGING RUNS FOR SPEED |
'|--------------------------------------------------------|
'   TAKEN FROM: http://stackoverflow.com/questions/198409/how-do-you-test-running-time-of-vba-code
'   USE WITH: CLASS MODULE QueryPerformanceCounter

Sub TimeDebbugger()
    Dim Timer As QueryPerformanceCounter
    Dim i As Long, j As Long
    Dim Output As Variant
    Dim TimeValue As Double
    Dim TimeCounter As Long
    Dim NumAveragingPeriods As Integer
    Dim NumberFunctionCalls As Integer

    NumAveragingPeriods = 2500
    NumberFunctionCalls = 2500

    'set timer options
    TimeValue = 0
    Set Timer = New QueryPerformanceCounter
    Timer.StartCounter
    For i = 1 To NumAveragingPeriods
        For j = 1 To NumberFunctionCalls
            '|----------------------|
            '| CALL A FUNCTION HERE |
            '|----------------------|
            Output = ColumnLetter(j)
			'|----------------------|
			'|----------------------|
			'|----------------------|			
        Next j
        TimeValue = Timer.TimeElapsed
    Next i
    Debug.Print "Average Time = " & TimeValue / NumAveragingPeriods
End Sub