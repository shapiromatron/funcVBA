:: BACKUP USING XCOPY
:: http://commandwindows.com/xcopy.htm
:: ----------------------------
XCOPY C:\Tools U:\Backup\Tools /M /S /D /Y >> BackupStatus.txt
:: XCOPY Source File, Destination File
:: /M = Copies only files with the archive attribute set, turns off the archive attribute. 
::      Useful in backup.
:: /S = Copies directories and subdirectories except empty ones.
:: /D = Copies files changed on or after the specified date. If no date is given, copies 
::      only those files whose source time is newer than the destination time. Useful in backup.
:: /Y = Suppress overwrite warning; always overwrite if date is newer
echo %date% - %time% - Backup Complete >> BackupStatus.txt
echo. >> BackupStatus.txt
