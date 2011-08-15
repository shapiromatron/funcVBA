:: Zip Commands
:: http://www.codejacked.com/zip-up-files-from-the-command-line/

:: ADD FILE TO ZIP FILE
::------------------------------------------
"C:\Program Files\7-zip\7z.exe" u -tzip text.zip outfile.txt

:: ADD FILE TO ZIP FILE, THEN DELETE FILE
::------------------------------------------
"C:\Program Files\7-zip\7z.exe" u -tzip text.zip outfile.txt && del outfile.txt

