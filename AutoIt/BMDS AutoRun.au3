#cs

 AutoIt Version: 3.0
 Language:       English
 Platform:       Win XP
 Author:         Andy Shapiro
 Last Updated:	7/8/11

	This script opens a text file to get the filename of the BMDS execuatable and the name and 
	location of the session file to be run.  Then, it opens BMDS, runs the session file, and
	exits BMDS.  
	
	Last update (7/8/11) includes major revisions to improve stability, the major shortcoming of 
	previous versions of the Autorun tool. It also appears to be much quicker than previous versions.

#ce

;==========================================================
;=============== VARIABLE DECLARATION =====================
;==========================================================
Dim $BMDS_Dir 
Dim $File
Dim $OutCount
Dim $ReadLine
Dim $i, $j
Dim $BMDS_Executable
Dim $WindowTitleBar

;for $j = 1 to 1		;DEBUG, MAKE >1
Opt("WinWaitDelay", 500)        ; in milliseconds
;AutoItSetOption("SendKeyDelay", 500)
$SessionLinkFileName = "BMDS_Session.txt"
$OutCount=0

;==========================================================
;================ START MAIN SCRIPT =======================
;==========================================================

;|----------------------------------------------------------|
;| OPEN TEXT INPUT FILE TO GET SESSION AND EXECUTABLE NAMES |
;|----------------------------------------------------------|
$File = FileOpen(@ScriptDir & "\" & $SessionLinkFileName,0)
	; Check if file opened for reading OK		
	If $File = -1 Then
		MsgBox(0, "Error", @ScriptDir & "\" & $SessionLinkFileName & " not found.")
		Exit
	EndIf
	; Read in lines of text until the EOF is reached
	While 1
		$OutCount = $OutCount + 1
		$ReadLine = FileReadLine($File)			
		If @error = -1 Then ExitLoop
		if $OutCount = 1 Then $BMDS_Dir = $ReadLine
		if $OutCount = 2 Then $BMDS_Executable = $ReadLine						
		if $OutCount = 3 Then $WindowTitleBar = $ReadLine			
	Wend
FileClose($File)	

;|------------------------------------|
;| OPEN BMDS AND WAIT UNTIL ACTIVATED |
;|------------------------------------|
Run(@ComSpec & " /c " & $BMDS_Executable, "", @SW_HIDE) 

;|-----------------------------------|
;| WAIT UNTIL BMDS COMPLETES LOADING |
;|-----------------------------------|
do
	WinWait($WindowTitleBar)
	WinWaitActive($WindowTitleBar)
	WinActivate($WindowTitleBar)
	WinSetOnTop($WindowTitleBar,"",1)
Until StringInStr ( WinGetText($WindowTitleBar, ""), "toolStrip1" ) <> 0

;|-------------|
;| OPEN WINDOW |
;|-------------|
Do
	for $i = 1 to 10
		send("{ESC}")
	Next
	;AutoItSetOption("SendKeyDelay", 200)
	Send("!fo{ENTER}")
	WinWait("Open","",10)
Until WinActive("Open","") = TRUE

;|-------------------|
;| LOAD SESSION FILE |
;|-------------------|
ControlSetText("Open","","Edit1",$BMDS_Dir)
ControlClick ( "Open", "", "Button2")

;|--------------------------|
;| WAIT UNTIL SESSION LOADS |
;|--------------------------|
do 
	sleep(250)
Until StringInStr ( WinGetText($WindowTitleBar, ""), "Run" ) <> 0

;|-------------|
;| RUN SESSION |
;|-------------|
ControlClick ($WindowTitleBar, "Run", "Run" , "left")

;|----------------------------------------------|
;| WAIT UNTIL "->" APPEARS, INDICATING COMPLETE |
;|----------------------------------------------|
do 
	sleep(250)
	Send("{ENTER}")
Until StringInStr ( WinGetText($WindowTitleBar, ""), "->" ) <> 0

;|-----------|
;| EXIT BMDS |
;|-----------|
Do
;	Send("{ENTER}")
;	sleep(250)
	ControlFocus($WindowTitleBar,"","menuStrip1")
;	send("!FX")
;	sleep(250)
	WinClose($WindowTitleBar)
until WinExists($WindowTitleBar,"") = False

;|-----------------------------------------------------------|
;| DELETE INPUT FILE; TRIGGER FOR EXCEL THAT RUN IS COMPLETE |
;|-----------------------------------------------------------|
FileDelete(@ScriptDir & "\" & $SessionLinkFileName)	; comment in debug mode

;Next