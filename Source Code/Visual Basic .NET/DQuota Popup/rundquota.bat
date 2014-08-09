@echo off
copy dquota.exe %WinDir%
cls
%WinDir%\dquota.exe /c:m > C:\dquotaresult.txt
copy "Dquota Popup.exe" %WinDir%
cls
%WinDir%\"Dquota Popup.exe"
