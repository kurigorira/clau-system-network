@echo off
rem ------------------------------------------------------------------
rem  Terminal inventory collector
rem
rem  Copied with Get-DeviceInventory.ps1 to C:\inventory by the startup
rem  batch, then started from it (see startup-snippet.bat).
rem  Writes COMPUTERNAME.csv to SHARE. Log: last-run.log next to this bat.
rem
rem  Exit code: 0 = OK, 1 = could not write the CSV,
rem             2 = SHARE not configured or not reachable
rem
rem  Usage:  run-inventory.bat                 (uses SHARE below)
rem          run-inventory.bat \\server\share   (overrides SHARE)
rem ------------------------------------------------------------------
setlocal

rem ===== Result folder (CHANGE THIS to the shared folder for results) =====
set SHARE=\\nagasakinet.local\dfsroot\CHANGE_ME
rem =======================================================================

if not "%~1"=="" set SHARE=%~1
set LOG=%~dp0last-run.log
echo [%date% %time%] start computer=%COMPUTERNAME% user=%USERNAME% output=%SHARE%> "%LOG%"

if not "%SHARE:CHANGE_ME=%"=="%SHARE%" (
    echo ERROR: SHARE is not configured. Edit run-inventory.bat on the server.>> "%LOG%"
    set RC=2
    goto :end
)

rem Network may not be ready right after boot: retry 3 times, 10 s apart.
set TRY=0
:waitshare
if exist "%SHARE%\" goto :run
set /a TRY+=1
if %TRY% geq 3 (
    echo ERROR: cannot reach %SHARE%>> "%LOG%"
    set RC=2
    goto :end
)
ping -n 11 127.0.0.1 >nul
goto :waitshare

:run
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Get-DeviceInventory.ps1" -OutputDir "%SHARE%" >> "%LOG%" 2>&1
set RC=%ERRORLEVEL%

:end
echo [%date% %time%] end exit=%RC%>> "%LOG%"
exit /b %RC%
