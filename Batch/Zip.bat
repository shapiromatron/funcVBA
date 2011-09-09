:: Zip Commands
:: http://www.codejacked.com/zip-up-files-from-the-command-line/
:: http://www.dotnetperls.com/7-zip-examples

:: ADD FILE TO ZIP FILE
::------------------------------------------
"C:\Program Files\7-zip\7z.exe" u -tzip text.zip outfile.txt

:: ADD FILE TO ZIP FILE, THEN DELETE FILE
::------------------------------------------
"C:\Program Files\7-zip\7z.exe" u -tzip text.zip outfile.txt && del outfile.txt

:: Extract a specific file from a zip file
::------------------------------------------
:: If root directory:
"C:\Program Files\7-zip\7z.exe" e TestZipFile.zip TestZipFile\outfile.txt
:: If deeper directory (in this case, an Excel file):
"C:\Program Files\7-zip\7z.exe" e Test.xlsx xl\media\image1.gif