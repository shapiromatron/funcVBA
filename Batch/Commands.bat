
::|--------------------------------------------------|
::| BATCH COMMANDS FILE FOR FREQUENTLY USED COMMANDS |
::|--------------------------------------------------|

:: GOOD WEBSITE:
:: -----------------------------------------
:: http://www.computerhope.com/msdos.htm


:: SYNTAX TIPS:
:: -----------------------------------------
:: ">" pushes all batch output from command to text file
::  example dir > dir.txt
:: "&" wait for commmand to complete before continuing
:: "&&" wait for command to complete succesfully before continuing

:: GRAB ALL FILES IN FOLDER, MOVE TO ROOT:
:: -----------------------------------------
@echo off
FOR /D /r %%G IN ("*") DO xcopy "%%G\*.*" "%~dp0"


:: GRAB ALL PDF FILES IN FOLDER, MOVE TO ROOT:
:: -----------------------------------------
FOR /D /r %%G IN ("*") DO xcopy "%%G\*.pdf" "%~dp0"

:: SEND OUTPUT OF BATCH FILE TO A TEXT FILE:
:: -----------------------------------------
:: BATCHFN.bat > OutputText.txt
GITbat.bat > GITout.txt

:: ADD AND COMMIT ALL FILES TO GIT 
::-------------------------------------
"cd " & OutputDir
git add -f . && git commit -a -m "Commit Message"

:: add files to staging area: 
git add -f .

:: commit and add commit note:
commit -a -m "Commit Note"

:: DELETE ALL FILES IN A DIRECTORY:
:: -----------------------------------------
:: @echo off
:: FOR /D %%G IN ("*") DO rmdir /S /Q "%%G"

