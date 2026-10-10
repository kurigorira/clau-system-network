rem ==================================================================
rem  Add these lines to the EXISTING startup copy batch (runs at boot).
rem  1) copy \\...\startup\inventory to C:\inventory
rem  2) start the collector in the background (does not delay boot)
rem  robocopy exit codes 0-7 mean success, 8 or more mean failure.
rem ==================================================================
robocopy "\\nagasakinet.local\dfsroot\newtons\startup\inventory" "C:\inventory" /E /R:1 /W:1 /XF last-run.log /NP /NFL /NDL /NJH /NJS >nul
if %ERRORLEVEL% lss 8 start "" /b cmd /c "C:\inventory\run-inventory.bat"
