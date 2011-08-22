Attribute VB_Name = "z_Binary"
Option Explicit

'Demonstration routine
Sub Test()
    Dim avValues() As Long, lThisRow As Long
    Dim avFileData As Variant, vThisBlock As Variant, vThisSubBlock As Variant
    Dim lThisBlock As Long
    
'    'Create an array of numbers
'    ReDim avValues(1 To 10)
'    For lThisRow = 1 To 10
'        avValues(lThisRow) = lThisRow
'    Next
'    'Write the array to a file
'    If FileWriteBinary(avValues, "C:\Test.dat") Then
        'Read the data back from the file
        avFileData = FileReadBinary("C:\Documents and Settings\16955\My Documents\SCAQMD\Binary to ASCII\grd1.wnd.990801.gcos")
        If IsArray(avFileData) Then
            'Print data
            Debug.Print "Values returned:"
            For Each vThisBlock In avFileData
                lThisBlock = lThisBlock + 1
                Debug.Print "Data Set:" & CStr(lThisBlock)
                For Each vThisSubBlock In vThisBlock
                    Debug.Print vThisSubBlock
                Next
            Next
            'Completed
            MsgBox "The array has been successfully retrieved!", vbInformation
        End If
'    End If
End Sub


Function FileWriteBinary(vData As Variant, sFileName As String, Optional bAppendToFile As Boolean = True) As Boolean
'http://www.visualbasic.happycodings.com/Files_Directories_Drives/code52.html
'Purpose     :  Saves/writes a block of data to a file
'Inputs      :  vData                   The data to store in the file. Can be an
'                                       array or any simple data type.
'               sFileName               The path and file name where the data is to be stored
'               [bAppendToFile]         If True will append the data to the existing file
'Outputs     :  Returns True if succeeded in saving data
'Notes       :  Saves data type (text and binary).
    Dim iFileNum As Integer, lWritePos As Long
    
    On Error GoTo ErrFailed
    If bAppendToFile = False Then
        If Len(Dir$(sFileName)) > 0 And Len(sFileName) > 0 Then
            'Delete the existing file
            VBA.Kill sFileName
        End If
    End If
    
    iFileNum = FreeFile
    Open sFileName For Binary Access Write As #iFileNum
    
    If bAppendToFile = False Then
        'Write to first byte
        lWritePos = 1
    Else
        'Write to last byte + 1
        lWritePos = LOF(iFileNum) + 1
    End If
    
    Put #iFileNum, lWritePos, vData
    Close iFileNum
    
    FileWriteBinary = True
    Exit Function

ErrFailed:
    FileWriteBinary = False
    Close iFileNum
    Debug.Print Err.Description
End Function


Function FileReadBinary(sFileName As String) As Variant
'http://www.visualbasic.happycodings.com/Files_Directories_Drives/code52.html
'Purpose     :  Reads the contents of a binary file
'Inputs      :  sFileName               The path and file name where the data is stored
'Outputs     :  Returns an array containing all the data stored in the file.
'               e.g. ArrayResults(1 to lNumDataBlocks)
'               Where lNumDataBlocks is the number of data blocks stored in file.
'               If the file was created using FileWriteBinary, this will be the number
'               of times data was appended to the file.
    Dim iFileNum As Integer, lFileLen As Long
    Dim vThisBlock As Variant, lThisBlock As Long, vFileData As Variant
    
    On Error GoTo ErrFailed
    
    If Len(Dir$(sFileName)) > 0 And Len(sFileName) > 0 Then
        iFileNum = FreeFile
        Open sFileName For Binary Access Read As #iFileNum
        
        lFileLen = LOF(iFileNum)
        
        Do
            lThisBlock = lThisBlock + 1
            Get #iFileNum, , vThisBlock
            If IsEmpty(vThisBlock) = False Then
                If lThisBlock = 1 Then
                    ReDim vFileData(1 To 1)
                Else
                    ReDim Preserve vFileData(1 To lThisBlock)
                End If
                vFileData(lThisBlock) = vThisBlock
            End If
        Loop While EOF(iFileNum) = False
        Close iFileNum
        
        FileReadBinary = vFileData
    End If

    Exit Function
    
ErrFailed:
    Close iFileNum
    Debug.Print Err.Description
End Function

