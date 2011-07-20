'|--------------------------------------------------------|
'| VERY PRECISE TIMER TO USE FOR DEBUGGING RUNS FOR SPEED |
'|--------------------------------------------------------|
'   TAKEN FROM: http://stackoverflow.com/questions/198409/how-do-you-test-running-time-of-vba-code
'   USE WITH: Sub TimeDebbugger (see below)

'Sub TimeDebbugger()
'    Dim Timer As QueryPerformanceCounter
'    Dim i As Long, j As Long
'    Dim Output As Variant
'    Dim TimeValue As Double
'    Dim TimeCounter As Long
'    Dim NumAveragingPeriods As Integer
'    Dim NumberFunctionCalls As Integer
'
'    NumAveragingPeriods = 2500
'    NumberFunctionCalls = 2500
'
'    'set timer options
'    TimeValue = 0
'    Set Timer = New QueryPerformanceCounter
'    Timer.StartCounter
'    For i = 1 To NumAveragingPeriods
'        For j = 1 To NumberFunctionCalls
'            '|----------------------|
'            '| CALL A FUNCTION HERE |
'            '|----------------------|
'            Output = ColumnLetter(j)
'        Next j
'        TimeValue = Timer.TimeElapsed
'    Next i
'    Debug.Print "Average Time = " & TimeValue / NumAveragingPeriods
'End Sub

Option Explicit

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long

Private m_CounterStart As LARGE_INTEGER
Private m_CounterEnd As LARGE_INTEGER
Private m_crFrequency As Double

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#

Private Function LI2Double(LI As LARGE_INTEGER) As Double
    Dim Low As Double
    Low = LI.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LI2Double = LI.highpart * TWO_32 + Low
End Function

Private Sub Class_Initialize()
    Dim PerfFrequency As LARGE_INTEGER
    QueryPerformanceFrequency PerfFrequency
    m_crFrequency = LI2Double(PerfFrequency)
End Sub

Public Sub StartCounter()
    QueryPerformanceCounter m_CounterStart
End Sub

Property Get TimeElapsed() As Double
    Dim crStart As Double
    Dim crStop As Double
    QueryPerformanceCounter m_CounterEnd
    crStart = LI2Double(m_CounterStart)
    crStop = LI2Double(m_CounterEnd)
    TimeElapsed = 1000# * (crStop - crStart) / m_crFrequency
End Property
