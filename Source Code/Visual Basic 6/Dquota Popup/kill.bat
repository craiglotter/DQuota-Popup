@echo off
setlocal
set ProgramName=dquota_popup.exe
set PID=
for /f "tokens=1" %%a in ('tasklist.exe ^| find /i "%ProgramName%"') do set PID=%%a
if "%PID%"=="" (
  echo The program "%ProgramName%" is not running.
  goto leave
)
echo The program "%ProgramName%" is running, the PID is %PID%.
taskkill.exe /F /IM %PID%
