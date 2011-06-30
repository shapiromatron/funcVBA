Option Explicit

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long

Sub RunAndKill()
'runs a program, and then kills it, also kills if timeout
    Dim ProcessID As Integer
    Dim TimeOut As Double
    
    'must change directory using ChDir before running Shell
    ChDir "C:\FOLDER\"
    ProcessID = Shell("PROGRAM.EXE" & " " & "C:\FOLDER\INPUTFILE.INP", vbHide)
    Do While IsProcessOpen(ProcessID) = True
        Application.StatusBar = StatusBarScenario & StatusBarDistance & " ||| Waiting for indoor dust model to complete... " & Now
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    
    
'    AppWait , , 1
'    TimeOut = Now + TimeValue("00:00:10")
'    'run BMDS for PLT file after output exists
'    Do While DoesFileExist("C:\FOLDER\OUTPUTFILE.OUT") = False
'        AppWait , , 1
'        If Now > TimeOut Then
'            MsgBox "Output file not created: C:\FOLDER\OUTPUTFILE.OUT"
'            Exit Sub
'        End If
'    Loop

    'force kill program
    Shell "TASKKILL /F /PID " & ProcessID, vbHide
    'Task Kill: http://commandwindows.com/taskkill.htm
End Sub

Function IsProcessOpen(PID As Long) As Boolean
'test to see if a process is currently running
    Dim h As Long
    h = OpenProcess(&H1, True, PID)
    If h <> 0 Then
        CloseHandle h
      IsProcessOpen = True
    Else
      IsProcessOpen = False
    End If
End Function

'Private Sub DummyTest()
''create file, then check if file changed
'    'create dummy output file
'    Open "C:\FOLDER\OUTPUT.OUT" For Output As #3
'        Print #3, "-999"
'    Close #3
'    LastMod = FileLastModified("C:\FOLDER\OUTPUT.OUT") 'gets the first time the file was modified
'    AppWait , , 2
'    'now call model
'    ProcessID = Shell("C:\FOLDER\PROGRAM.EXE" & " " & "C:\FOLDER\INPUTFILE.INP", vbHide)
'    'now wait until file last modified changes
'    Do While ThisMod = LastMod
'        AppWait , , 2
'        ThisMod = FileLastModified("C:\FOLDER\OUTPUT.OUT")
'    Loop
'End Sub
'Function FileLastModified(FullFileName As String) As String
''http://www.ozgrid.com/forum/showthread.php?t=27740
'    Dim fs As Object, f As Object
'    On Error GoTo IsError
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set f = fs.GetFile(FullFileName)
'    FileLastModified = f.DateLastModified
'    Exit Function
'IsError:
'    FileLastModified = "-999"
'End Function
