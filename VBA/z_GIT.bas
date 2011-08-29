Attribute VB_Name = "z_GIT"
Option Explicit

'|----------------------------------------------------------|
'| SUPPPORTING INCLUDES FOR SHELL SCRIPTS FOR MAIN FUNCTION |
'|----------------------------------------------------------|
Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
                ByVal lpFile As String, ByVal lpParameters As String, _
                ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
                                                            ByVal bInheritHandle As Long, _
                                                            ByVal dwProcId As Long) As Long

'|---------------|
'| MAIN FUNCTION |
'|---------------|
Function CommitToGIT(OutputDir As String) As Boolean
        '-----------------------------------------------------------------------------------------------------------
        ' CommitToGIT   - Commits changes in all files to GIT repository
                '                               - Directory must already contain a GIT repository
        '               - In : OutputDir as string
        '               - Out: TRUE if succesful, FALSE if otherwise
        '               - Last Updated: 7/20/11 by AJS
        '-----------------------------------------------------------------------------------------------------------
    Dim CommitMessage As String
    Dim ProcessID As Long
    
    On Error GoTo IsError
        Do
                CommitMessage = InputBox("Enter GIT commit input message: ", "GIT Revisions Message")
        Loop Until CommitMessage <> ""
    Open OutputDir & "GITbat.bat" For Output As #1
        Print #1, "cd " & OutputDir     'change directory to tracking folder
        Print #1, "git add -f . && git commit -a -m " & Chr(34) & CommitMessage & Chr(34)
        'add files to staging area, wait until completes succesfully, commit changes into tool
    Close #1
    ChDir OutputDir
    
    ProcessID = Shell("GITbat.bat > GITout.txt", vbNormalFocus)
    Do While IsProcessOpen(ProcessID) = True
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    OpenAnyFileUsingDefaultProgram (OutputDir & "GITout.txt")
    CommitToGIT = True
    Exit Function
IsError:
    CommitToGIT = False
End Function

'|------------------------------------|
'| SUPPPORTING SUBS FOR MAIN FUNCTION |
'|------------------------------------|
Private Sub OpenAnyFileUsingDefaultProgram(FullFileName As String)
    '-----------------------------------------------------------------------------------------------------------
    ' OpenAnyFileUsingDefaultProgram    - Tests to see if ProcessID is currently open
    '                                   - Last Updated: 7/14/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    ' Requires the following system functions (declare at top of module):
    '   Declare Function ShellExecute Lib "shell32.dll" Alias _
    '   "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    '                   ByVal lpFile As String, ByVal lpParameters As String, _
    '                   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    '-----------------------------------------------------------------------------------------------------------
    ShellExecute 0, "open", FullFileName, 0, 0, 1
End Sub

Private Function IsProcessOpen(PID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    ' IsProcessOpen    - Tests to see if ProcessID is currently open
    '                  - Last Updated: 7/14/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    ' Requires the following system functions (declare at top of module):
    '   Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
    '   Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
    '                                                               ByVal bInheritHandle As Long, _
    '                                                               ByVal dwProcId As Long) As Long
    '-----------------------------------------------------------------------------------------------------------
    Dim h As Long
    h = OpenProcess(&H1, True, PID)
    If h <> 0 Then
        CloseHandle h
      IsProcessOpen = True
    Else
      IsProcessOpen = False
    End If
End Function
