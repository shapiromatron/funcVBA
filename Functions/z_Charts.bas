Attribute VB_Name = "z_Charts"
Option Explicit

'/-------------------------------\
'| FOR EACH chart in a workBOOK  |
'\-------------------------------/
Private Sub eachChart()
    Dim ChartID As Chart
    Application.DisplayAlerts = False
    For Each ChartID In Charts
        If ChartID.Name <> "ENV Template" And Chart.Name <> "Region Template" Then
            ChartID.Delete
        End If
    Next
    Application.DisplayAlerts = True
    Set ChartID = Nothing
End Sub

'/-------------------------------\
'| FOR EACH chart in a workSHEET |
'\-------------------------------/
Private Sub EachChartObject()
    Dim ChartID As ChartObject
    Application.DisplayAlerts = False
    For Each ChartID In ActiveSheet.ChartObject
        If ChartID.Name <> "ENV Template" And Chart.Name <> "Region Template" Then
            ChartID.Delete
        End If
    Next
    Application.DisplayAlerts = True
    Set ChartID = Nothing
End Sub

Function Series_Update(ChartName As String, SeriesName As String, XRange As Range, YRange As Range) As Variant
    '---------------------------------------------------------------------------------------------------------
    ' Series_Update      - Updates a series for a chart in a workbook
    '                    - In : ChartName As String, SeriesName As String, XRange As Range, YRange As Range
    '                    - Out: TRUE if succesful, error if false
    '                    - Last Updated: 8/22/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    Dim eachChart As Chart
    Dim eachSeries As Series
    Dim FullSeries As String
    
    On Error GoTo IsError:
    For Each eachChart In ActiveWorkbook.Charts
        If eachChart.Name = ChartName Then
            eachChart.Activate
            For Each eachSeries In eachChart.SeriesCollection
                If eachSeries.Name = SeriesName Then
                    eachSeries.Formula = "=SERIES(" & _
                        Chr(34) & SeriesName & Chr(34) & "," & _
                        z_Charts.PrintChartAddress(XRange) & ", " & _
                        z_Charts.PrintChartAddress(YRange) & "," & _
                        eachSeries.PlotOrder & ")"
                End If
            Next
        End If
    Next
    Series_Update = True
    Exit Function
IsError:
    Series_Update = CVErr(xlErrNA)
    Debug.Print "Error in Series_Update: " & Err.Number & ": " & Err.Description
End Function

Function PrintChartAddress(thisRange As Range) As Variant
    '---------------------------------------------------------------------------------------------------------
    ' PrintChartAddress  - Prints the address of a range as written in chart form (ex: 'Sheet1'!$B4$B26)
    '                    - In : thisRange As Range
    '                    - Out: formatted string if succesful, error if false
    '                    - Last Updated: 8/22/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    On Error GoTo IsError:
    PrintChartAddress = "'" & thisRange.Worksheet.Name & "'!" & thisRange.Address
    Exit Function
IsError:
    PrintChartAddress = CVErr(xlErrNA)
    Debug.Print "Error in PrintChartAddress: " & Err.Number & ": " & Err.Description
End Function
