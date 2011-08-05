
::|--------------------------------------------------|
::| BATCH COMMANDS FILE FOR FREQUENTLY USED COMMANDS |
::|--------------------------------------------------|

:: GOOD WEBSITE:
:: -----------------------------------------
:: http://www.computerhope.com/msdos.htm


:: SYNTAX TIPS:
:: -----------------------------------------
:: ">" pushes all batch output from command to text file
:: "&" wait for commmand to complete before continuing
:: "&&" wait for command to complete succesfully before continuing

:: GRAB ALL FILES IN FOLDER, MOVE TO ROOT:
:: -----------------------------------------
@echo off
FOR /D /r %%G IN ("*") DO xcopy "%%G\*.*" "%~dp0"


:: GRAB ALL PDF FILES IN FOLDER, MOVE TO ROOT:
:: -----------------------------------------
@echo off
FOR /D /r %%G IN ("*") DO xcopy "%%G\*.pdf" "%~dp0"

:: SEND OUTPUT OF BATCH FILE TO A TEXT FILE:
:: -----------------------------------------
:: BATCHFN.bat > OutputText.txt
GITbat.bat > GITout.txt

:: DELETE ALL FILES IN A DIRECTORY:
:: -----------------------------------------
:: @echo off
:: FOR /D %%G IN ("*") DO rmdir /S /Q "%%G"


:: DELETE ONE FILE
::------------------------------------------
DEL outputfile.txt

:: CHANGE DIRECTORY AND DRIVE AT ONCE:
cd /D M:\16955-Andy Shapiro

:: Set a variable
set OutputFile=GITStatus.txt

:: Set a user input variable
Set Input= 
Set /P Input=Type Add Description of Changes to GIT:

:: Run command, send output to file
git add -f . && git commit -a -m "%Input%" >> %OutputFile%

:: cmd && cmd (do second comand if only true)

:: Open File using default program
START %OutputFile%

:: BEEP
::------------------------------------------


::#==================#
::#   ZIP COMMANDS   #
::#==================#

:: ADD FILE TO ZIP FILE, THEN DELETE FILE
::------------------------------------------
:: http://www.codejacked.com/zip-up-files-from-the-command-line/
"C:\Program Files\7-zip\7z.exe" u -tzip text.zip outfile.txt && del outfile.txt


::#==================#
::#   GIT COMMANDS   #
::#==================#

:: GIT: ADD AND COMMIT ALL FILES
::-------------------------------------
"cd " & OutputDir
git add -f . && git commit -a -m "Commit Message"

:: add files to staging area: 
git add -f .

:: commit and add commit note:
commit -a -m "Commit Note"


::#=====================#
::#   PYTHON COMMANDS   #
::#=====================#

:: PYTHON: PROFILE A SCRIPT
:: -----------------------------------------
cd C:\ScriptLocation
python -m cProfile C:/ScriptLocation/ScriptName.py >> profile_summary.txt