option explicit
'Tools -> References -> Adobe Acrobat 7.0 Type Library
'JSO = java script object
Private Sub SavePDFasTXT(FileName As String)
    Dim PDFname As String, TXTname As String
    Dim AcroXApp As Object, AcroXAVDoc As Object, AcroXPDDoc As Object, jsObj As Object
    
    PDFname = FileName
    TXTname = Left(FileName, Len(FileName) - 4) & ".txt"
    
    Set AcroXApp = CreateObject("AcroExch.App")
    AcroXApp.Hide
    Set AcroXAVDoc = CreateObject("AcroExch.AVDoc")
    AcroXAVDoc.Open PDFname, "Acrobat"
    AcroXAVDoc.BringToFront
    Set AcroXPDDoc = AcroXAVDoc.GetPDDoc
    Set jsObj = AcroXPDDoc.GetJSObject
    jsObj.SaveAs TXTname, "com.adobe.acrobat.plain-text" 'save as a text file
    AcroXAVDoc.Close False
    AcroXApp.Hide
    AcroXApp.Exit
End Sub
