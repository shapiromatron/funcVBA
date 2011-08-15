
::|--------------------------------------------------|
::| BATCH COMMANDS FILE FOR FREQUENTLY USED COMMANDS |
::|--------------------------------------------------|

:: GOOD WEBSITE:
:: -----------------------------------------
:: http://www.computerhope.com/msdos.htm


:: SYNTAX TIPS:
:: -----------------------------------------
:: ">" create new file; print all output to file
:: ">>" append output to existing file or create new
:: "&" wait for commmand to complete before continuing
:: "&&" wait for command to complete succesfully before continuing

:: GRAB ALL FILES IN FOLDER, MOVE TO ROOT:
@echo off
FOR /D /r %%G IN ("*") DO xcopy "%%G\*.*" "%~dp0"


:: GRAB ALL PDF FILES IN FOLDER, MOVE TO ROOT:
@echo off
FOR /D /r %%G IN ("*") DO xcopy "%%G\*.pdf" "%~dp0"

:: SEND OUTPUT OF BATCH FILE TO A TEXT FILE:
:: BATCHFN.bat > OutputText.txt
GITbat.bat > GITout.txt

:: DELETE ALL FILES IN A DIRECTORY:
:: @echo off
:: FOR /D %%G IN ("*") DO rmdir /S /Q "%%G"

:: DELETE ONE FILE
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


:: WAIT FOR A NUMBER OF MILLISECONDS
PING 1.1.1.1 -n 1 -w 5000 >NUL

:: Append a blank line to text file
echo. >> output.txt