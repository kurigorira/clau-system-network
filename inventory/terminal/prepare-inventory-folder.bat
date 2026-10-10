@echo off
rem ------------------------------------------------------------------
rem  One-time preparation for PCs where ordinary users cannot create
rem  C:\inventory (typical on Windows 7).
rem  Right-click - "Run as administrator". Creates C:\inventory and gives
rem  the local Users group Modify permission on it, so the logon batch
rem  (running as an ordinary user) can robocopy into it.
rem  Safe to run again.
rem  For many PCs at once use admin\Prepare-InventoryFolder.ps1 instead.
rem ------------------------------------------------------------------
setlocal
set DIR=C:\inventory

net session >nul 2>&1
if not errorlevel 1 goto :isadmin
echo [NG] Not running as administrator.
echo      Right-click this file and choose "Run as administrator".
goto :done

:isadmin
if not exist "%DIR%\" mkdir "%DIR%"
if exist "%DIR%\" goto :grant
echo [NG] Could not create %DIR%
goto :done

:grant
rem *S-1-5-32-545 = local "Users" group, works on any Windows language
icacls "%DIR%" /grant *S-1-5-32-545:(OI)(CI)M
if errorlevel 1 goto :grantng
echo.
icacls "%DIR%"
echo [OK] %DIR% is ready. Ordinary users can now copy into it.
goto :done

:grantng
echo [NG] icacls failed. See the message above.

:done
echo.
pause
