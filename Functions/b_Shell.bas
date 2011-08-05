Option Explicit

Sub OpenAnyFileUsingDefaultProgram(FullFileName as String)
    '-----------------------------------------------------------------------------------------------------------
    ' OpenAnyFileUsingDefaultProgram	- Opens any file on your computer using the default program
    '                  					- Last Updated: 7/14/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
	' Requires the following system functions (declare at top of module): 	
	'	Declare Function ShellExecute Lib "shell32.dll" Alias _
    '	"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
	'					ByVal lpFile As String, ByVal lpParameters As String, _
	'					ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
	'-----------------------------------------------------------------------------------------------------------	
	ShellExecute 0, "open", FullFileName, 0, 0, 1
End Sub


Sub RunAndKill()
    '-----------------------------------------------------------------------------------------------------------
    ' RunAndKill	   	- Runs a program, then kills it, if time is greater than timeout
	'					- Requires: Function IsProcessOpen
    '                  	- Last Updated: 7/14/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    Dim ProcessID As Long	'Must be set to Long, not Variant
    Dim TimeOut As Double
	Dim EXE_Dir as String
	DIM EXE_Name as String
	Dim TimeoutSeconds as Integer
	Dim RunSuccessful as Boolean
	
	EXE_Dir = "C:\USEPA\BMDS212\"
	EXE_Name = "BMDS2.exe"
	TimeoutSeconds = 30
	
	TimeOut = Now + TimeSerial(0,0,TimeoutSeconds)        
    ChDir EXE_Dir
    ProcessID = Shell(EXE_Name & " " & EXE_Dir & "\INPUTFILE.INP", vbHide)
	RunSuccessful=True
    Do While IsProcessOpen(ProcessID) = True
        Application.StatusBar = " Waiting for " & EXE_Name & " to complete... " & Now
        Application.Wait Now + TimeSerial(0, 0, 1)
		
		' See if input file is deleted (indication that run is complete)
		If z_Files.GetFileInfo(Range("IEUBK_AutoItFN"), FileExists) = False Then
            RunSuccessful = True
            Exit Do
        End If
		
		' See if time is exceeded (force kill program)
		if Now > TimeOut then 			
			RunSuccessful= False
			Shell "TASKKILL /F /PID " & ProcessID, vbHide
			Exit Do
		end if
    Loop
	
	Application.ScreenUpdating = False
	
	If RunSuccessful = TRUE Then
		Debug.Print EXE_Name & " completed succesfully"
	Else
		Debug.Print EXE_Name & " timed out"
	end If
End Sub

Private Function IsProcessOpen(PID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    ' IsProcessOpen	   - Tests to see if ProcessID is currently open
    '                  - Last Updated: 7/14/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
	' Requires the following system functions (declare at top of module): 
	'	Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
	'	Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
	'																ByVal bInheritHandle As Long, _
	'																ByVal dwProcId As Long) As Long	
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