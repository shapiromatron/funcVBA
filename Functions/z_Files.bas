Attribute VB_Name = "z_Files"
Option Explicit

Enum FileTypes
        '----------------------------------------------------------------
        ' FileTypes    - Used with SelectNewOrExistingFile, SelectExistingFile, GetCustomFilterList
        '----------------------------------------------------------------
        AnyExtension = 0
        ExcelFiles = 1
        ExcelFileOrTemplate = 2
        WordFiles = 3
        WordFileOrTemplate = 4
        TextFiles = 5
        CSVFiles = 6
        Custom = 99
End Enum

Enum GetFileInfo
        '----------------------------------------------------------------
        ' GetFileInfo    - Used with GetFileInfo
        '----------------------------------------------------------------
        PathOnly = 1
        NameAndExtension = 2
        NameOnly = 3
        ExtensionOnly = 4
        ParentFolder = 5
        FileExists = 6
        FolderExists = 7
        DateLastMod = 8
        FileSizeKB = 9
End Enum

Public Function SelectNewOrExistingFile(Optional FileType As FileTypes = 0, Optional MenuTitleName = "Select File", Optional StartingPath = "WBPath", Optional CustomFilter As String = "Any File (*.*), *.*") As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' SelectNewOrExistingFile   - Select a new or an existing file, using custom filters for specific file types if needed
        '                               New Function in Excel 2007; will not work with previous versions of Excel (http://msdn.microsoft.com/en-us/library/bb209903(v=office.12).aspx)
        '                           - In : Optional FileType as FileTypes (defined above, specify file filters, by default any file)
        '                               Optional MenuTitleName = "Select File" (Default)
        '                               Optional Strpath = Workbook Path (Default)
        '                               Optional CustomFilter As String = "Any File (*.*), *.*" (Custom Filter if User-defined)
        '                           - Out: Full Path to selected file, or FALSE if user cancelled
        '                           - Requires: Function ReturnCustomFilterList
        '                           - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim OutputFile As Variant
        On Error GoTo IsError
        CustomFilter = GetCustomFilterList(FileType, CustomFilter)
        If StartingPath = "WBPath" Then StartingPath = CStr(ThisWorkbook.Path)
        Do
                OutputFile = Application.GetSaveAsFilename(StartingPath, CustomFilter, 1, MenuTitleName)
                If GetFileInfo(CStr(OutputFile), FileExists) = False Then
                        Exit Do
                Else
                        If vbYes = MsgBox("File already exists, replace existing file?" & vbNewLine & vbNewLine & OutputFile, vbYesNo, "Replace existing file?") Then Exit Do
                End If
        Loop
        SelectNewOrExistingFile = OutputFile
        Exit Function
IsError:
        SelectNewOrExistingFile = CVErr(xlErrNA)
        Debug.Print "Error in SelectNewOrExistingFile: " & Err.Number & ": " & Err.Description
End Function

Public Function SelectExistingFolder(Optional MenuTitleName As String = "Select Folder", Optional ByVal StartingPath As String = "WBPath") As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' SelectExistingFolder  - Selecting an existing folder
        '                       - In :  Optional MenuTitleName = "Select Folder" (Default)
        '                               Optional Strpath = Workbook Path (Default)
        '                       - Out: Folder Path including final backslash "\"
        '                       - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim fldr As FileDialog
        On Error GoTo IsError
        If StartingPath = "WBPath" Then StartingPath = CStr(ThisWorkbook.Path)
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
        With fldr
                .InitialView = msoFileDialogViewDetails
                .Title = MenuTitleName
                .AllowMultiSelect = False
                .InitialFileName = StartingPath
                If .Show <> -1 Then GoTo UserCancelled
                SelectExistingFolder = .SelectedItems(1) & "\"
        End With
        Exit Function
UserCancelled:
        SelectExistingFolder = False
IsError:
        SelectExistingFolder = CVErr(xlErrNA)
        Debug.Print "Error in SelectExistingFolder: " & Err.Number & ": " & Err.Description
End Function

Public Function SelectExistingFile(Optional FileType As FileTypes = 0, Optional MenuTitleName = "Select File", Optional StartingPath = "WBPath", Optional CustomFilter As String = "Any File (*.*), *.*") As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' SelectExistingFile    - Selecting an exisiting file, using custom filters for pre-defined file types or create a new custom file type
        '                       - In : Optional FileType as FileTypes (defined above, specify file filters, by default any file)
        '                               Optional MenuTitleName = "Select File" (Default)
        '                               Optional Strpath = Workbook Path (Default)
        '                               Optional CustomFilter As String = "Any File (*.*), *.*" (Custom Filter if User-defined)
        '                       - Out: Full Path to selected file, or FALSE if user cancelled
        '                       - Requires: Function ReturnCustomFilterList
        '                       - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        
        CustomFilter = GetCustomFilterList(FileType, CustomFilter)
        If StartingPath = "WBPath" Then StartingPath = CStr(ThisWorkbook.Path)
        ChDir StartingPath
        SelectExistingFile = Application.GetOpenFilename(FileFilter:=CustomFilter, Title:=MenuTitleName, MultiSelect:=False)
        Exit Function
IsError:
        SelectExistingFile = CVErr(xlErrNA)
        Debug.Print "Error in SelectExistingFile: " & Err.Number & ": " & Err.Description
End Function

Private Function GetCustomFilterList(FileTypeNumber As FileTypes, CustomFilter As String) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' ReturnCustomFilterList    - Returns custom filter lists for each specified type of file
        '                           - In : FileTypeNumber FileType as FileTypes (defined above, specify file filters, by default any file)
        '                                    CustomFilter As String = "Any File (*.*), *.*" (only used if custom filetypes is selected)
        '                           - Out: FilterList as string
        '                           - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim ReturnValue As String
        On Error GoTo IsError
        Select Case FileTypeNumber
                Case 0
                        ReturnValue = "Any File (*.*),*.*"
                Case 1
                        ReturnValue = "Excel File (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls"
                Case 2
                        ReturnValue = "Excel File or Excel Template (*.xlsx; *.xlsm; *.xls; *.xlt; *.xltx; *.xltm), *.xlsx; *.xlsm; *.xls; *.xlt; *.xltx; *.xltm"
                Case 3
                        ReturnValue = "Word File (*.docx; *.docm; *.doc), *.docx; *.docm; *.doc"
                Case 4
                        ReturnValue = "Word File or Word Template (*.docx; *.docm; *.doc; *.dotx; *.dotm; *.dot), *.docx; *.docm; *.doc; *.dotx; *.dotm; *.dot"
                Case 5
                        ReturnValue = "Text File (*.txt; *.dat), *.txt; *.dat"
                Case 6
                        ReturnValue = "CSV File (*.csv), *.csv"
                Case 99
                        ReturnValue = CustomFilter
        End Select
        GetCustomFilterList = ReturnValue
        Exit Function
IsError:
        GetCustomFilterList = CVErr(xlErrNA)
        Debug.Print "Error in GetCustomFilterList: " & Err.Number & ": " & Err.Description
End Function

Public Function MakeDirString(PathString As String) As Variant
        '---------------------------------------------------------------------------------------------------------
        ' MakeDirString      - Adds a parenthesis to the end of a path if it doesn't already exist
        '                    - In : PathString As String
        '                    - Out: MakeDirString as string if valid, error if not valid
        '                    - Last Updated: 7/3/11 by AJS
        '---------------------------------------------------------------------------------------------------------
        On Error GoTo IsError
        If Right(PathString, 1) <> "\" Then
                MakeDirString = PathString & "\"
        Else
                MakeDirString = PathString
        End If
        Exit Function
IsError:
        MakeDirString = CVErr(xlErrNA)
        Debug.Print "Error in MakeDirString: " & Err.Number & ": " & Err.Description & vbNewLine & PathString
End Function

Public Function MakeDirFullPath(Path As String) As Boolean
        '-----------------------------------------------------------------------------------------------------------
        ' MakeDirFullPath   - Creates the full path directory if it doesn't already exist, can for example
        '                       create C:\Temp\Temp\Temp if it doesn't alreay dexist
        '                   - In : Path as String
        '                   - Out: TRUE if path exists, FALSE if path doesn't exist
        '                   - Last Updated: 7/2/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim UncreatedPaths As Collection, EachPath As Variant
        Set UncreatedPaths = New Collection
        Dim NewPath As String
        
        On Error GoTo IsError
        NewPath = Path
        Do While GetFileInfo(NewPath, FolderExists) = False
                UncreatedPaths.Add NewPath
                NewPath = GetFileInfo(NewPath, ParentFolder)
        Loop
        Do While UncreatedPaths.Count > 0
                MkDir UncreatedPaths(UncreatedPaths.Count)
                UncreatedPaths.Remove UncreatedPaths.Count
        Loop
        MakeDirFullPath = GetFileInfo(Path, FolderExists)
        Exit Function
IsError:
        MakeDirFullPath = GetFileInfo(Path, FolderExists)
        Debug.Print "Error in MakeDirFullPath: " & Err.Number & ": " & Err.Description & vbNewLine & Path
End Function

Public Function FileListInFolder(ByVal PathName As String, Optional ByVal FileFilter As String = "*.*") As Collection
        '-----------------------------------------------------------------------------------------------------------
        ' FileList           - Returns a collection of files in a given folder with the specified filter
        '                       Can filter by a certain type of filename, if file filter is set to equal a certain extension
        '                       Replacement for Application.FileSearch, removed from Excel 2007
        '                       Uses MSDOS Dir function: http://www.computerhope.com/dirhlp.htm
        '                    - In : PathName As String, Optional FileFilter As String
        '                    - Out: A string collection of file names in the specified folder
        '                    - Created: Greg Haskins
        '                    - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim sTemp As String, sHldr As String
        Dim RetVal As New Collection
        
        On Error GoTo IsError
        If Right$(PathName, 1) <> "\" Then PathName = PathName & "\"
        sTemp = Dir(PathName & FileFilter)
        If sTemp = "" Then
                Set FileListInFolder = RetVal
                Exit Function
        Else
                RetVal.Add sTemp
        End If
        Do
                sHldr = Dir
                If sHldr = "" Then Exit Do
                'sTemp = sTemp & "|" & sHldr
                RetVal.Add sHldr
        Loop
        'FileList = Split(sTemp, "|")
        Set FileListInFolder = RetVal
        Exit Function
IsError:
        FileListInFolder.Add CVErr(xlErrNA)
        Debug.Print "Error in FileListInFolder: " & Err.Number & ": " & Err.Description & vbNewLine & PathName & FileFilter
End Function

Public Function GetFileInfo(FN As String, FileInfo As GetFileInfo, Optional ShowErrorPopup As Boolean = False) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' GetFileInfo        - Returns key file information for a file or folder passed to the function, uses the enumeration GetFileInfo
        '                            1: PathOnly            (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "C:\USEPA\BMDS212\00Hill.exe")
        '                            2: NameAndExtension    (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "00Hill.exe")
        '                            3: NameOnly            (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "00Hill")
        '                            4: ExtensionOnly       (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "exe")
        '                            5: ParentFolder        (FN = "C:\USEPA\BMDS212\",           Return = "C:\USEPA\")
        '                            6: FileExists          (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = TRUE)
        '                            7: FolderExists        (FN = "C:\USEPA\BMDS212\",           Return = TRUE)
        '                            8: DateLastMod         (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = "5/20/2010 1:23:56 AM")
        '                            9: FileSizeKB          (FN = "C:\USEPA\BMDS212\00Hill.exe", Return = 12.8)
        '                       (May also display a popup message if file or folder doesn't exist)
        '                    - In : FN As String, FileInfo As GetFileInfo
        '                    - Out: Depends on the file info type selected, error if error
        '                    - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim fso As Object
        On Error GoTo IsError
        Set fso = CreateObject("Scripting.FileSystemObject")
        Select Case FileInfo
                Case 1
                        GetFileInfo = fso.GetParentFolderName(FN) & "\"
                Case 2
                        GetFileInfo = fso.GetFileName(FN)
                Case 3
                        GetFileInfo = fso.GetBaseName(FN)
                Case 4
                        GetFileInfo = fso.GetExtensionName(FN)
                Case 5
                        GetFileInfo = fso.GetParentFolderName(FN) & "\"
                Case 6
                        GetFileInfo = fso.FileExists(FN)
                        If ShowErrorPopup = True And GetFileInfo = False Then MsgBox "Error- file doesn't exist!" & vbNewLine & vbNewLine & FN, vbCritical, "File does not exist!"
                Case 7
                        GetFileInfo = fso.FolderExists(FN)
                        If ShowErrorPopup = True And GetFileInfo = False Then MsgBox "Error- folder doesn't exist!" & vbNewLine & vbNewLine & FN, vbCritical, "Folder does not exist!"
                Case 8
                        GetFileInfo = CStr(fso.GetFile(FN).DateLastModified)
                Case 9
                        GetFileInfo = FileLen(FN) / 1000
                Case Else
                        GoTo IsError
        End Select
        Exit Function
IsError:
        GetFileInfo = CVErr(xlErrNA)
        Debug.Print "Error in GetFileInfo: " & Err.Number & ": " & Err.Description & vbNewLine & FN
End Function

Private Function DoesFileExist(FN As String) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' DoesFileExist      - Alternate way to test to see if file exists (instead of GetFileInfo)
        '                    - In : FN as String
        '                    - Out: TRUE/FALSE if filename is valid
        '                    - Last Updated: 7/20/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
    Dim fso As Object
    On Error GoTo IsError
    Set fso = CreateObject("Scripting.FileSystemObject")
    DoesFileExist = fso.FileExists(FN)
    Exit Function
IsError:
    DoesFileExist = CVErr(xlErrNA)
    Debug.Print "Error in Private Function DoesFileExist: " & Err.Number & ": " & Err.Description
End Function

Public Function IsValidFileName(FN As String) As Boolean
        '-----------------------------------------------------------------------------------------------------------
        ' IsValidFileName    - Returns true if filename is valid using the Win32 naming scheme
        '                    - Adapted from: http://www.bytemycode.com/snippets/snippet/334/
        '                    - In : FN as String
        '                    - Out: TRUE/FALSE if filename is valid
        '                    - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim RE As Object, REMatches As Object
        On Error GoTo IsError
        Set RE = CreateObject("vbscript.regexp")
        
        With RE
                .MultiLine = False
                .Global = False
                .IgnoreCase = True
                .Pattern = "[\\\/\:\*\?\" & Chr(34) & "\<\>\|]" 'If any of the following characters are found: \ / : * ? " < > |
        End With
        Set REMatches = RE.Execute(FN)
        If REMatches.Count > 0 Or FN = "" Then
                MsgBox "Filename not valid: " & vbNewLine & FN, vbCritical, "Filename not valid"
                IsValidFileName = False
        Else
                IsValidFileName = True
        End If
        Exit Function
IsError:
        IsValidFileName = False
        Debug.Print "Error in IsValidFileName: " & Err.Number & ": " & Err.Description & vbNewLine & FN
End Function

Public Function IsFileOpen(FN As String) As Variant
        '-----------------------------------------------------------------------------------------------------------
        ' IsFileOpen    - Returns TRUE if file is currently open, FALSE if it's not open, or error if other error occurs
        '               - Adapted from: http://www.vbaexpress.com/kb/getarticle.php?kb_id=468
        '               - In : FN as String
        '               - Out: TRUE if file is currently open, FALSE if it's not open, or error if other error occurs
        '               - Last Updated: 7/3/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
        Dim iErr As Long, iFilenum As Long
        On Error Resume Next
                Err.Clear
                iFilenum = FreeFile()
                Open FN For Input Lock Read As #iFilenum
                Close iFilenum
                iErr = Err
        On Error GoTo 0
        Select Case iErr
                Case 0:    IsFileOpen = False
                Case 70:   IsFileOpen = True
                Case Else: Error iErr
        End Select
End Function

Public Function Kill2(ByVal PathName As String) As Boolean
        '----------------------------------------------------------------
        ' Kill2             - Deletes file; continues until succesfullly deleted
        '                   - In : ByVal PathName As String
        '                   - Out: Boolean true if file is succesfully removed
        '                   - Last Updated: 7/3/11 by AJS
        '----------------------------------------------------------------
        Dim TimeOut As String
        TimeOut = Now + TimeValue("00:00:10")
        On Error Resume Next
                Do While GetFileInfo(PathName, FileExists) = True
                        Kill PathName
                        If Now > TimeOut Then
                                MsgBox "Error- File deletion has time out, file cannot be deleted:" & vbNewLine & vbNewLine & PathName, vbCritical, "Error in deleting file"
                                GoTo IsError
                        End If
                Loop
        On Error GoTo 0
        Kill2 = True
        Exit Function
IsError:
        Kill2 = False
        Debug.Print "Error in Kill2: " & Err.Number & ": " & Err.Description & vbNewLine & PathName
End Function

Public Function FileCopy2(ByVal SourceFile As String, ByVal DestinationFile As String) As Boolean
        '----------------------------------------------------------------
        ' FileCopy2             - Revised version of FileCopy that will return TRUE when file is actually copied
        '                       - In : SourceFile As String, DestinationFile As String
        '                       - Out: Boolean true if file is succesfully copied; false otherwise
        '                       - Last Updated: 7/3/11 by AJS
        '----------------------------------------------------------------
        Dim TimeOut As String
        TimeOut = Now + TimeValue("00:00:10")
        If GetFileInfo(SourceFile, FileExists) = False Then
                MsgBox "Error- file does not exist and cannot be copied:" & vbNewLine & vbNewLine & SourceFile, vbCritical, "File cannot be copied"
                GoTo IsError
        End If
        If GetFileInfo(DestinationFile, FileExists) = True Then Kill2 DestinationFile
        On Error Resume Next
        Do While GetFileInfo(DestinationFile, FileExists) = False
                FileCopy SourceFile, DestinationFile
                If Now > TimeOut Then
                        MsgBox "Error- File copy has timed out, file was probably not succesfully copied (may be open?):" & vbNewLine & vbNewLine & _
                                        "Source: " & SourceFile & vbNewLine & _
                                        "Destination: " & DestinationFile, vbCritical, "Error in copying file"
                        GoTo IsError
                End If
        Loop
        FileCopy2 = True
        Exit Function
IsError:
        FileCopy2 = False
        Debug.Print "Error in FileCopy2: " & Err.Number & ": " & Err.Description & vbNewLine & SourceFile & vbNewLine & DestinationFile
End Function

