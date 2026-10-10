<#
.SYNOPSIS
  Ping-sweeps IPv4 ranges and lists live hosts (IP, MAC, host name).

.DESCRIPTION
  Used to find terminals that did NOT report through Get-DeviceInventory.ps1
  (script not deployed, non-Windows devices, printers, medical equipment...).
  MAC addresses come from the ARP table, so they are only available for hosts
  on the same subnet (L2 segment) as the PC running this script. Run it on a
  PC in each segment, or use the switch / DHCP server tables for MACs.

.PARAMETER Subnet
  One or more /24 prefixes, e.g. 192.168.10 (scans .1 - .254).

.EXAMPLE
  .\Find-NetworkHosts.ps1 -Subnet 192.168.10,192.168.11 -OutFile C:\inventory\discovery.csv
#>
param(
    [Parameter(Mandatory = $true)][string[]]$Subnet,
    [Parameter(Mandatory = $true)][string]$OutFile,
    [int]$TimeoutMs = 500
)

$live = @()
foreach ($s in $Subnet) {
    $s = $s.TrimEnd('.')
    Write-Host "Scanning $s.1-254 ..."
    $jobs = foreach ($i in 1..254) {
        $ip = "$s.$i"
        $p = New-Object System.Net.NetworkInformation.Ping
        New-Object PSObject -Property @{ IP = $ip; Task = $p.SendPingAsync($ip, $TimeoutMs) }
    }
    try { [System.Threading.Tasks.Task]::WaitAll([System.Threading.Tasks.Task[]]($jobs | ForEach-Object { $_.Task })) } catch {}
    $live += $jobs | Where-Object { $_.Task.Status -eq 'RanToCompletion' -and $_.Task.Result.Status -eq 'Success' } | ForEach-Object { $_.IP }
}

# ARP / neighbor table (filled by the pings above)
$macByIp = @{}
foreach ($line in (arp -a)) {
    if ($line -match '^\s*(\d+\.\d+\.\d+\.\d+)\s+([0-9a-fA-F]{2}(-[0-9a-fA-F]{2}){5})') {
        $macByIp[$Matches[1]] = $Matches[2].ToUpper()
    }
}

$rows = foreach ($ip in $live) {
    $name = ''
    try { $name = [System.Net.Dns]::GetHostEntry($ip).HostName } catch {}
    New-Object PSObject -Property @{ IPAddress = $ip; MACAddress = $macByIp[$ip]; HostName = $name }
}

$rows | Select-Object IPAddress, MACAddress, HostName |
    Sort-Object { [Version]$_.IPAddress } |
    Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
Write-Host "Live hosts: $(@($rows).Count) -> $OutFile"
