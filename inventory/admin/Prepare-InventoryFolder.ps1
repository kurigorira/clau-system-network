<#
.SYNOPSIS
  One-time preparation of C:\inventory on PCs where ordinary users cannot
  create it (typical on Windows 7). Run from the admin PC as a domain admin.

.DESCRIPTION
  For each target PC, through the administrative share \\PC\C$:
    1. ping (skips PCs that are switched off)
    2. create \\PC\C$\inventory
    3. icacls: give the local Users group (SID S-1-5-32-545) Modify on it
  Then the logon batch, which runs as an ordinary user, can robocopy into
  C:\inventory. Safe to run again: existing folders only get the permission
  re-applied. Results go to -OutFile (CSV): OK / Offline / CreateFailed /
  IcaclsFailed. Re-run later for the PCs that were Offline.

.EXAMPLE
  # All Windows 7 computers in Active Directory
  .\Prepare-InventoryFolder.ps1 -FromActiveDirectory
.EXAMPLE
  # PCs listed in a text file (one name per line)
  .\Prepare-InventoryFolder.ps1 -ComputerListFile .\win7.txt
#>
param(
    [string[]]$ComputerName,
    [string]$ComputerListFile,
    [switch]$FromActiveDirectory,
    [string]$OSFilter = '*Windows 7*',
    [string]$OutFile = '.\prepare-result.csv',
    # For testing only
    [string]$PathTemplate = '\\{0}\C$\inventory',
    [switch]$SkipPing
)

$targets = @()
if ($ComputerName)     { $targets += $ComputerName }
if ($ComputerListFile) { $targets += Get-Content $ComputerListFile | Where-Object { $_.Trim() -and $_ -notmatch '^\s*#' } | ForEach-Object { $_.Trim() } }
if ($FromActiveDirectory) {
    Import-Module ActiveDirectory -ErrorAction Stop
    $targets += Get-ADComputer -Filter "Enabled -eq 'True' -and OperatingSystem -like '$OSFilter'" | Select-Object -ExpandProperty Name
}
$targets = @($targets | Select-Object -Unique)
if ($targets.Count -eq 0) { throw 'No target computers. Use -ComputerName, -ComputerListFile or -FromActiveDirectory.' }

$results = foreach ($pc in $targets) {
    $path = $PathTemplate -f $pc
    $status = 'OK'; $detail = ''
    if (-not $SkipPing -and -not (Test-Connection -ComputerName $pc -Count 1 -Quiet)) {
        $status = 'Offline'
    } else {
        try {
            New-Item -ItemType Directory -Path $path -Force -ErrorAction Stop | Out-Null
            $out = & icacls.exe $path /grant '*S-1-5-32-545:(OI)(CI)M' 2>&1
            if ($LASTEXITCODE -ne 0) { $status = 'IcaclsFailed'; $detail = ($out | Out-String).Trim() }
        } catch [System.Management.Automation.CommandNotFoundException] {
            $status = 'IcaclsFailed'; $detail = 'icacls.exe not found'
        } catch {
            $status = 'CreateFailed'; $detail = $_.Exception.Message
        }
    }
    Write-Host ("{0,-20} {1}" -f $pc, $status)
    New-Object PSObject -Property @{ ComputerName = $pc; Status = $status; Detail = $detail; Path = $path }
}

$results | Select-Object ComputerName, Status, Detail, Path | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
$summary = $results | Group-Object Status | ForEach-Object { "$($_.Name)=$($_.Count)" }
Write-Host ("Done: {0}  -> {1}" -f ($summary -join ', '), $OutFile)
