@echo off
if not exist %WinDir%\dquota.exe copy dquota.exe %WinDir%
%WinDir%\dquota.exe /c:m > c:\dquotaresult.txt
exit
