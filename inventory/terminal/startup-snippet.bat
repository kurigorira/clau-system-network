rem ==================================================================
rem  Add these 2 lines to the END of the EXISTING startup copy batch.
rem  Do NOT put this file in the distribution folder itself.
rem
rem  CHECK THE PATH: Explorer shows the share name, e.g.
rem    "newton (\\NAGASAKINET.local\dfsroot)" = \\nagasakinet.local\dfsroot\newton
rem  Test it first:  dir \\nagasakinet.local\dfsroot\newton\startup\inventory
rem
rem  1) copy the distribution folder to C:\inventory
rem     (copy log: C:\inventory-copy.log)
rem  2) start the collector in the background (does not delay boot)
rem     robocopy exit codes 0-7 mean success, 8 or more mean failure.
rem ==================================================================
robocopy "\\nagasakinet.local\dfsroot\newton\startup\inventory" "C:\inventory" /E /R:1 /W:1 /XF last-run.log /NP /NFL /NDL /LOG:C:\inventory-copy.log
if %ERRORLEVEL% lss 8 start "" /b cmd /c "C:\inventory\run-inventory.bat"
