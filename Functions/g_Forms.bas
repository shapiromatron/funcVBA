Option Explicit
Private Sub ShowStudyForm()
    With Frm_AddExcerpt
        'center window on Excel before showing
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub
