Option Explicit

'|----------------------|
'| View and Center Form |
'|----------------------|
Private Sub ShowStudyForm()
    With Frm_AddExcerpt
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub


'|-----------------------------|
'| Common Initialization Tasks |
'|-----------------------------|
Private Sub UserForm_Initialize()


	'LOAD SHAPE IMAGE INTO A IMAGE OBJECT ON FORM (requires z_PastePicture)
	Do
		Sheets(SN_Validation).Shapes(EMFName).Copy
		Me.Image1.Picture = z_PastePicture.PastePicture(xlPicture)
	Loop Until IsPictureLoaded(Me.Image1) = True
	
	
End Sub

Private Function IsPictureLoaded(ImageObject As Image) as Boolean
    '---------------------------------------------------------------------------------------------------------
    ' IsPictureLoaded - Returns TRUE if a picture is loaded into an image object in a form, FALSE if otherwise
    '                 - In : ImageObject as Image
    '                 - Out: Boolean TRUE or FALSE
    '                 - Last Updated: 7/29/11 by AJS
    '---------------------------------------------------------------------------------------------------------
    On Error Resume Next
        IsPictureLoaded = True
        If IsError(ImageObject.Picture.handle) <> False Then
            IsPictureLoaded = False
        End If
    On Error GoTo 0
End Function