<#
.SYNOPSIS
  Runs Get-DeviceInventory.ps1 on many terminals from an admin PC (WinRM).

.DESCRIPTION
  Requires WinRM (PowerShell Remoting) enabled on the terminals and an account
  with local administrator rights. Each reachable host is saved as
  <OutputDir>\<COMPUTERNAME>.csv; unreachable hosts are listed in
  <OutputDir>\_failed.txt.

.PARAMETER ComputerName
  Host names or IP addresses. Use -ComputerListFile for a text file
  (one name per line) or -FromActiveDirectory to take all enabled computers.

.EXAMPLE
  .\Invoke-RemoteInventory.ps1 -FromActiveDirectory -OutputDir C:\inventory\raw
.EXAMPLE
  .\Invoke-RemoteInventory.ps1 -ComputerListFile .\hosts.txt -OutputDir C:\inventory\raw -Credential (Get-Credential)
#>
param(
    [string[]]$ComputerName,
    [string]$ComputerListFile,
    [switch]$FromActiveDirectory,
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [System.Management.Automation.PSCredential]$Credential,
    [int]$ThrottleLimit = 32
)

$targets = @()
if ($ComputerName)     { $targets += $ComputerName }
if ($ComputerListFile) { $targets += Get-Content $ComputerListFile | Where-Object { $_.Trim() -and $_ -notmatch '^\s*#' } | ForEach-Object { $_.Trim() } }
if ($FromActiveDirectory) {
    Import-Module ActiveDirectory -ErrorAction Stop
    $targets += Get-ADComputer -Filter 'Enabled -eq $true' | Select-Object -ExpandProperty Name
}
$targets = $targets | Select-Object -Unique
if (-not $targets) { throw 'No target computers. Use -ComputerName, -ComputerListFile or -FromActiveDirectory.' }

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$collector = Join-Path $PSScriptRoot 'Get-DeviceInventory.ps1'

$invokeArgs = @{
    ComputerName  = $targets
    FilePath      = $collector
    ThrottleLimit = $ThrottleLimit
    ErrorAction   = 'SilentlyContinue'
    ErrorVariable = 'remoteErrors'
}
if ($Credential) { $invokeArgs.Credential = $Credential }

$results = Invoke-Command @invokeArgs
foreach ($r in $results) {
    $r | Select-Object -Property * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName |
        Export-Csv -Path (Join-Path $OutputDir "$($r.ComputerName).csv") -NoTypeInformation -Encoding UTF8
}

$failed = $remoteErrors | ForEach-Object { "$($_.TargetObject)`t$($_.Exception.Message)" }
$failed | Set-Content -Path (Join-Path $OutputDir '_failed.txt') -Encoding UTF8

Write-Host ("Collected: {0} / {1}  (failed: {2}, see _failed.txt)" -f @($results).Count, @($targets).Count, @($failed).Count)
