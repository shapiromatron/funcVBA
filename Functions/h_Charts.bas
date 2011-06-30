Option Explicit

'/-------------------------------\
'| FOR EACH chart in a workBOOK  |
'\-------------------------------/
Private Sub EachChart()
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
