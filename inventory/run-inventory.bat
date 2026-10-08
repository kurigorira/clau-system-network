@echo off
rem GPO startup/logon script or Task Scheduler entry.
rem Change the share path below to your file server.
set SHARE=\\fs01\inventory$
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SHARE%\Get-DeviceInventory.ps1" -OutputDir "%SHARE%\raw"
