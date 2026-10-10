@echo off
rem ------------------------------------------------------------------
rem  Runs Get-DeviceInventory.ps1 placed in the SAME folder as this bat
rem  and writes COMPUTERNAME.csv to the shared folder.
rem
rem  Usage:  run-inventory.bat                 (uses SHARE below)
rem          run-inventory.bat \\server\share   (overrides SHARE)
rem
rem  Works from the server copy tool, GPO startup script or Task Scheduler.
rem  The result of the last run is kept in last-run.log next to this bat.
rem ------------------------------------------------------------------
setlocal
set SHARE=\\fs01\inventory$\raw
if not "%~1"=="" set SHARE=%~1
set LOG=%~dp0last-run.log

echo [%date% %time%] start computer=%COMPUTERNAME% user=%USERNAME% output=%SHARE%> "%LOG%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Get-DeviceInventory.ps1" -OutputDir "%SHARE%" >> "%LOG%" 2>&1
set RC=%ERRORLEVEL%
echo [%date% %time%] end exit=%RC%>> "%LOG%"
exit /b %RC%
