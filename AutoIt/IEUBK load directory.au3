#cs
	Name: Batch IEUBK autorun tool
	Created by: Andy Shapiro
	Last updated: 07/15/11
	
   Loops through the number of input files specified below and copies each input file in the IEUBK 
   directory, run each batch file, and copies each output file to the specified directy and file notation
   below.
   
   Note that the selected parameter file is loaded into IEUBK initally, and therefore all batch runs
   use the same set of input parameters.
   
#ce

;|-----------------------|
;| IEUBK STANDARD INPUTS |
;|-----------------------|
const $AUTOIT_InputFN = "C:\Program Files\IEUBKwin1_1 Build11\Input\IEUBK_AutoRun.inp"
Const $IEUBK_DIR = "C:\Program Files\IEUBKwin1_1 Build11"
Const $IEUBK_EXE = "IEUBKwin32.exe"
Const $BATCH_DIR = "C:\Program Files\IEUBKwin1_1 Build11\Input\"
Const $BATCH_FN = "AutoItInput.dat"
Const $IEUBK_DEFAULT_OUTFN = "C:\Program Files\IEUBKwin1_1 Build11\Output\BatchRun.txt"
Dim $PID, $i, $j
Dim  $File, $OutCount, $ReadLine
Dim $BATCH_INPUT_FN, $BATCH_PARAM_FN, $BATCH_OUTPUT_FN

;|-------------------------------|
;| TEST TO SEE INPUT FILES EXIST |
;|-------------------------------|
if FileExists($AUTOIT_InputFN) = false Then
	MsgBox(0, "Error", $AUTOIT_InputFN & " not found.")
	Exit
EndIf
	
;|-----------------------|
;| LOAD INPUTS FORM FILE |
;|-----------------------|
$File = FileOpen($AUTOIT_InputFN,0)
	; Read in lines of text until the EOF is reached
	$OutCount=1
	While 1
		$ReadLine = FileReadLine($File)
		If @error = -1 Then ExitLoop
		Select 
		   case $OutCount=1
				$BATCH_INPUT_FN = $ReadLine
		   case $OutCount=2
				$BATCH_PARAM_FN = $ReadLine
		   case $OutCount=3	   
				$BATCH_OUTPUT_FN = $ReadLine
		EndSelect
		$OutCount = $OutCount + 1
	Wend
FileClose($File)    

;|------------|
;| OPEN IEUBK |
;|------------|
FileChangeDir($IEUBK_DIR)
$PID = Run($IEUBK_EXE)
WinWaitActive("IEUBKwin32 Lead Model Version 1.1 Build11")

;|-------------------|
;| CLOSE INTRO STUFF |
;|-------------------|
splashoff()
for $i = 1 to 10
	send("{ESCAPE}")
next

;|---------------------|
;| LOAD PARAMETER MENU |
;|---------------------|
DO 
	for $i = 1 to 10
		send("{ESCAPE}")
	next
	send("!p{UP}{UP}{ENTER}")
	sleep(300)
Until WinActive("Open SVD File") = TRUE

;|-------------------------------|
;| LOAD PARAMETER FILE AND CLOSE |
;|-------------------------------|
do 
	ControlSetText("Open SVD File","","Edit1",$BATCH_PARAM_FN)
	ControlClick ( "Open SVD File", "", "Button2")
until WinActive("IEUBKwin32") = True
Do
	send("{ENTER}")
until winactive("IEUBKwin32 Lead Model Version 1.1 Build11") = True

;|----------------------------------------------|
;| COPY INPUT FILE AND PUT INTO INPUT DIRECTORY |
;|----------------------------------------------|
FileCopy($BATCH_INPUT_FN,$BATCH_DIR & $BATCH_FN,1)

;|-------------------|
;| LOAD BATCH FILE   |
;|-------------------|
DO 
	for $i = 1 to 10
		send("{ESCAPE}")
	next
	send("!CBB{ENTER}{ENTER}")
	Sleep(200)
Until WinActive("Select Batch File(s)") = TRUE

;|--------------------------------|
;| LOAD BATCH FILE AND SELECT RUN |
;|--------------------------------|
do 
	ControlCommand ( "Select Batch File(s)", "", "ListBox1", "SelectString", $BATCH_FN)
	ControlClick ( "Select Batch File(s)", "", "Button4")	
	ControlClick ( "Select Batch File(s)", "", "Button1")
until WinActive("Batch Mode Run") = True

;|--------------------------------|
;| RUN BATCH, WAIT UNTIL COMPLETE |
;|--------------------------------|
ControlClick ( "Batch Mode Run", "", "Button2")	
do 
	SLEEP (100)
Until WinActive("Batch Mode Run") = FALSE
	
;|-------------------|
;| CLOSE OUTPUT FILE |
;|-------------------|
for $i = 1 to 10
	send("{ESCAPE}")
next
send("!FCN")

;|------------------------------------------------|
;| COPY OUTPUT FILE AND PUT INTO OUTPUT DIRECTORY |
;|------------------------------------------------|
FileCopy($IEUBK_DEFAULT_OUTFN, $BATCH_OUTPUT_FN,1)
	
;|-------------|
;| CLOSE IEUBK |
;|-------------|
ProcessClose($PID)

;|-------------------|
;| DELETE INPUT FILE |
;|-------------------|
FileDelete($AUTOIT_InputFN)