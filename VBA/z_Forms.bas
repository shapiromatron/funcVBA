Attribute VB_Name = "z_Forms"
Option Explicit

Sub PreviousClick(formname As UserForm, ControlName As String)
    'Goes to previous value in control if possible; otherwise no change
    Dim j As Integer
    If formname.Controls(ControlName).Locked = False Then
        For j = 0 To formname.Controls(ControlName).ListCount - 1
            If CStr(formname.Controls(ControlName).Value) = CStr(formname.Controls(ControlName).List(j, 0)) Then
                If j <> 0 Then
                    formname.Controls(ControlName).Value = formname.Controls(ControlName).List(j - 1, 0)
                End If
                Exit For
            End If
        Next j
    End If
End Sub
Sub NextClick(formname As UserForm, ControlName As String)
    'Goes to next value in control if possible; otherwise no change
    Dim j As Integer
    If formname.Controls(ControlName).Locked = False Then
        For j = 0 To formname.Controls(ControlName).ListCount - 1
            If CStr(formname.Controls(ControlName).Value) = CStr(formname.Controls(ControlName).List(j, 0)) Then
                If j <> formname.Controls(ControlName).ListCount - 1 Then
                    formname.Controls(ControlName).Value = formname.Controls(ControlName).List(j + 1, 0)
                End If
                Exit For
            End If
        Next j
    End If
End Sub

Sub UpdateFormList(formname As UserForm, FieldName As String, NamedRange As Range)
    'Updates the list of values in a drop-down combo box in a userform
    Dim CellRng As Range
    With formname.Controls(FieldName)
        .Clear
        For Each CellRng In NamedRange
            .AddItem CellRng.Value
        Next
    End With
End Sub

Public Sub OpenAndCenterForm(formname As Variant)
    With formname
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub
