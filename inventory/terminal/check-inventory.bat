@echo off
rem ------------------------------------------------------------------
rem  Diagnostic tool: double-click on a terminal to check every step.
rem  Shows [OK] / [NG] for each step and stops at the first problem.
rem  Edit SRC below if the distribution folder path is different.
rem  Run it as an ordinary user (the logon batch also runs as the
rem  logged-on user) to test the real permissions.
rem ------------------------------------------------------------------
setlocal
set SRC=\\nagasakinet.local\dfsroot\newton\startup\inventory
set DST=C:\inventory

echo ==================================================================
echo  Inventory check   computer=%COMPUTERNAME%  user=%USERNAME%
echo ==================================================================

echo.
echo [1] Distribution folder: %SRC%
if exist "%SRC%\" goto :step2
echo   [NG] Cannot reach this folder.
echo        Check the path in Explorer - newton or newtons? - and edit SRC
echo        in this file and in the startup batch.
goto :done

:step2
echo   [OK] reachable
echo.
echo [2] Files directly in the distribution folder
if exist "%SRC%\run-inventory.bat" if exist "%SRC%\Get-DeviceInventory.ps1" goto :step2ok
echo   [NG] run-inventory.bat / Get-DeviceInventory.ps1 not found directly in
echo        %SRC%
if exist "%SRC%\terminal\run-inventory.bat" echo        They are inside the "terminal" subfolder. Move the 2 files UP into
if exist "%SRC%\terminal\run-inventory.bat" echo        %SRC% and remove admin, terminal and .md files there.
goto :done
:step2ok
echo   [OK] both files found
if exist "%SRC%\admin\" echo   [!!] "admin" folder is also there - not needed on terminals, please move it.
if exist "%SRC%\terminal\" echo   [!!] "terminal" folder is also there - not needed, please move it.

echo.
echo [3] Copy to %DST%
robocopy "%SRC%" "%DST%" /E /R:1 /W:1 /XF last-run.log /NP /NFL /NDL /NJH
if errorlevel 8 goto :copyng
if not exist "%DST%\run-inventory.bat" goto :copyng
echo   [OK] copied
goto :step4
:copyng
echo   [NG] Copy failed. See the robocopy messages above.
echo        "Access is denied" on C:\inventory: ordinary users cannot create
echo        or write C:\inventory on this PC - typical on Windows 7.
echo        Run prepare-inventory-folder.bat once as administrator on this PC,
echo        or admin\Prepare-InventoryFolder.ps1 from the admin PC.
goto :done

:step4
echo.
echo [4] Result folder - SHARE in run-inventory.bat
set CSHARE=
for /f "tokens=1,* delims==" %%a in ('findstr /b /i /c:"set SHARE=" "%DST%\run-inventory.bat"') do set CSHARE=%%b
echo   SHARE=%CSHARE%
if "%CSHARE%"=="" goto :shareunset
if not "%CSHARE:CHANGE_ME=%"=="%CSHARE%" goto :shareunset
if exist "%CSHARE%\" goto :step5
echo   [NG] Cannot reach the result folder. Check the path and that it exists.
goto :done
:shareunset
echo   [NG] SHARE is not set. Edit "set SHARE=" in run-inventory.bat on the SERVER
echo        - %SRC%\run-inventory.bat - then run this check again.
goto :done

:step5
echo   [OK] reachable
echo.
echo [5] Run the collector - takes 10 to 60 seconds, please wait ...
call "%DST%\run-inventory.bat" "%CSHARE%" 0
set RC=%ERRORLEVEL%
echo   exit code = %RC%   - 0 = OK, 1 = could not write CSV, 2 = SHARE problem
echo   ---- %DST%\last-run.log ----
type "%DST%\last-run.log"
echo   ----------------------------
if not "%RC%"=="0" goto :runng

echo.
echo [6] Result file
if exist "%CSHARE%\%COMPUTERNAME%_%USERNAME%.csv" goto :allok
:runng
echo   [NG] The CSV was not written. See last-run.log above.
echo        "Access is denied" means no write permission on the result folder.
goto :done
:allok
echo   [OK] %CSHARE%\%COMPUTERNAME%_%USERNAME%.csv
echo.
echo   All steps OK.

:done
echo.
pause
