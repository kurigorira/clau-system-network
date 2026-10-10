<#
.SYNOPSIS
  Collects inventory information from the local Windows terminal.

.DESCRIPTION
  Runs on each terminal (GPO startup/logon script, scheduled task, USB, or
  remotely through Invoke-Command) and gathers:
    host name, IPv4 / MAC addresses, Windows version, Office version,
    Office license key (full key only when recoverable, otherwise last 5 chars),
    last boot time, OS install date, asset tag, serial number, model.

  No Internet access is required. Compatible with Windows PowerShell 2.0+.
  This file is intentionally ASCII only so that it runs regardless of the
  code page of the terminal.

.PARAMETER OutputDir
  Folder (usually a UNC share such as \\fileserver\inventory$\raw) where one
  CSV per host (<COMPUTERNAME>.csv) is written. When omitted, the result
  object is written to the pipeline (used by Invoke-RemoteInventory.ps1).

.PARAMETER FileName
  Output file name without ".csv". Default: COMPUTERNAME. run-inventory.bat
  passes COMPUTERNAME_USERNAME so that every user who logs on writes (and can
  overwrite) only their own file.

.PARAMETER SkipIfNewerThanHours
  When the output file already exists and was written within this many hours,
  exit immediately with code 0 (avoids collecting at every logon). 0 = always.

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File Get-DeviceInventory.ps1 -OutputDir \\fs01\inventory$\raw
#>
param(
    [string]$OutputDir,
    [string]$FileName = $env:COMPUTERNAME,
    [int]$SkipIfNewerThanHours = 0
)

$ErrorActionPreference = 'Continue'

if ($OutputDir -and $SkipIfNewerThanHours -gt 0) {
    $existing = Join-Path $OutputDir "$FileName.csv"
    try {
        if (Test-Path -LiteralPath $existing) {
            $age = (Get-Date) - (Get-Item -LiteralPath $existing).LastWriteTime
            if ($age.TotalHours -lt $SkipIfNewerThanHours) {
                Write-Host ("SKIP: {0} was written {1:N1} hours ago (limit {2} h)" -f $existing, $age.TotalHours, $SkipIfNewerThanHours)
                exit 0
            }
        }
    } catch {}
}

function Get-RegValue {
    param([string]$Path, [string]$Name)
    try {
        $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return $item.$Name
    } catch {
        return $null
    }
}

function ConvertFrom-WmiDate {
    param([string]$Value)
    if (-not $Value) { return $null }
    try {
        return [System.Management.ManagementDateTimeConverter]::ToDateTime($Value)
    } catch {
        return $null
    }
}

# Decodes a product key stored in a DigitalProductId blob (Office 2010 and
# earlier). Office 2013 and later do not store the full key on the PC.
function ConvertFrom-DigitalProductId {
    param([byte[]]$Bytes)
    if (-not $Bytes -or $Bytes.Length -lt 67) { return $null }
    $chars = 'BCDFGHJKMPQRTVWXY2346789'
    $key = New-Object byte[] 15
    [Array]::Copy($Bytes, 52, $key, 0, 15)
    $result = ''
    for ($i = 24; $i -ge 0; $i--) {
        $cur = 0
        for ($j = 14; $j -ge 0; $j--) {
            $cur = $cur * 256 + $key[$j]
            $key[$j] = [byte][Math]::Floor($cur / 24)
            $cur = $cur % 24
        }
        $result = [string]$chars[$cur] + $result
        if (($i % 5) -eq 0 -and $i -ne 0) { $result = '-' + $result }
    }
    if ($result -eq 'BBBBB-BBBBB-BBBBB-BBBBB-BBBBB') { return $null }
    return $result
}

# ---------------------------------------------------------------- network
$ipList  = @()
$macList = @()
try {
    $adapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' -ErrorAction Stop
    foreach ($a in $adapters) {
        $v4 = @($a.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' -and $_ -notmatch '^(0\.|127\.|169\.254\.)' })
        if ($v4.Count -eq 0) { continue }
        $ipList  += $v4
        if ($a.MACAddress) { $macList += ($a.MACAddress -replace ':', '-').ToUpper() }
    }
} catch {}

# ---------------------------------------------------------------- OS
$os = $null
try { $os = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop } catch {}

$cvKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$displayVersion = Get-RegValue $cvKey 'DisplayVersion'
if (-not $displayVersion) { $displayVersion = Get-RegValue $cvKey 'ReleaseId' }
$build = Get-RegValue $cvKey 'CurrentBuild'
$ubr   = Get-RegValue $cvKey 'UBR'
if ($build -and $ubr -ne $null) { $build = "$build.$ubr" }

$windowsVersion = ''
if ($os) {
    $windowsVersion = $os.Caption.Trim()
    if ($displayVersion) { $windowsVersion += " $displayVersion" }
    if ($os.OSArchitecture) { $windowsVersion += " ($($os.OSArchitecture))" }
}

$lastBoot    = $null
$installDate = $null
if ($os) {
    $lastBoot    = ConvertFrom-WmiDate $os.LastBootUpTime
    $installDate = ConvertFrom-WmiDate $os.InstallDate
}

# ---------------------------------------------------------------- estimated in-service date
# The OS install date is reset by Windows 10/11 feature updates, so take the
# oldest of: current install date, pre-upgrade install dates kept under
# HKLM\SYSTEM\Setup\Source OS (...), and the oldest user profile folder.
$dateCandidates = @()
if ($installDate) { $dateCandidates += $installDate }
$epoch = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
foreach ($k in (Get-ChildItem 'HKLM:\SYSTEM\Setup' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like 'Source OS*' })) {
    $sec = Get-RegValue $k.PSPath 'InstallDate'
    if ($sec) { $dateCandidates += $epoch.AddSeconds([double]$sec).ToLocalTime() }
}
$usersDir = Join-Path $env:SystemDrive 'Users'
foreach ($d in (Get-ChildItem $usersDir -Force -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer })) {
    if (@('Default', 'Default User', 'Public', 'All Users') -contains $d.Name) { continue }
    $dateCandidates += $d.CreationTime
}
$estimatedStart = $null
$valid = @($dateCandidates | Where-Object { $_ -and $_.Year -ge 2000 -and $_ -le (Get-Date) } | Sort-Object)
if ($valid.Count -gt 0) { $estimatedStart = $valid[0] }

# ---------------------------------------------------------------- hardware
$serial = ''
$maker  = ''
$model  = ''
$assetTag = ''
$biosDate = $null
try {
    $bios = Get-WmiObject -Class Win32_BIOS -ErrorAction Stop
    $serial = "$($bios.SerialNumber)".Trim()
    $biosDate = ConvertFrom-WmiDate $bios.ReleaseDate
} catch {}
try {
    $cs = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
    $maker = "$($cs.Manufacturer)".Trim()
    $model = "$($cs.Model)".Trim()
} catch {}
# Asset tag: 1) registry value set by the hospital  2) BIOS (SMBIOS) asset tag
$assetTag = Get-RegValue 'HKLM:\SOFTWARE\HospitalInventory' 'AssetTag'
if (-not $assetTag) {
    try {
        $enc = Get-WmiObject -Class Win32_SystemEnclosure -ErrorAction Stop | Select-Object -First 1
        $t = "$($enc.SMBIOSAssetTag)".Trim()
        if ($t -and $t -notmatch '^(No Asset (Tag|Information)|Default string|To Be Filled By O\.E\.M\.|None|0+|\s*)$') {
            $assetTag = $t
        }
    } catch {}
}

# ---------------------------------------------------------------- Office version
$officeProducts = @()
$officeVersions = @()

# Click-to-Run (Office 2016/2019/2021/2024/Microsoft 365)
$c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$c2rIds = Get-RegValue $c2r 'ProductReleaseIds'
if ($c2rIds) {
    $officeProducts += ($c2rIds -split ',' | ForEach-Object { $_.Trim() })
    $v = Get-RegValue $c2r 'VersionToReport'
    $p = Get-RegValue $c2r 'Platform'
    if ($v) { $officeVersions += "$v $p (C2R)".Trim() }
}

# MSI installations (Office 2007/2010/2013/2016 volume etc.)
$uninstallRoots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)
foreach ($root in $uninstallRoots) {
    if (-not (Test-Path $root)) { continue }
    foreach ($k in (Get-ChildItem $root -ErrorAction SilentlyContinue)) {
        $p = $null
        try { $p = Get-ItemProperty $k.PSPath -ErrorAction Stop } catch { continue }
        $name = "$($p.DisplayName)"
        if ($name -notmatch '^Microsoft (Office|365)') { continue }
        if ($p.SystemComponent -eq 1) { continue }
        if ($name -match 'Language|Proofing|MUI|Viewer|Interop|Runtime|Components|Click-to-Run|Licensing|Shared|Web Components|Validation|Update|Add-in') { continue }
        if ($officeProducts -notcontains $name) {
            $officeProducts += $name
            if ($p.DisplayVersion) { $officeVersions += "$($p.DisplayVersion)" }
        }
    }
}

# ---------------------------------------------------------------- Office license
$officeAppId = '0ff1ce15-a989-479d-af46-f275c6370663'
$licenses = @()
$wql = "SELECT Name, PartialProductKey, LicenseStatus FROM {0} WHERE ApplicationId = '$officeAppId' AND PartialProductKey IS NOT NULL"
foreach ($cls in @('SoftwareLicensingProduct', 'OfficeSoftwareProtectionProduct')) {
    try {
        foreach ($l in (Get-WmiObject -Query ($wql -f $cls) -ErrorAction Stop)) {
            $status = switch ($l.LicenseStatus) {
                0 { 'Unlicensed' } 1 { 'Licensed' } 2 { 'OOBGrace' } 3 { 'OOTGrace' }
                4 { 'NonGenuineGrace' } 5 { 'Notification' } 6 { 'ExtendedGrace' } default { "$($l.LicenseStatus)" }
            }
            $licenses += New-Object PSObject -Property @{
                Name   = "$($l.Name)"
                Key    = "XXXXX-XXXXX-XXXXX-XXXXX-$($l.PartialProductKey)"
                Status = $status
            }
        }
    } catch {}
}

# Full key from DigitalProductId (Office 2010 / 2007 only)
$fullKeys = @()
foreach ($base in @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office')) {
    foreach ($ver in @('14.0', '12.0')) {
        $reg = "$base\$ver\Registration"
        if (-not (Test-Path $reg)) { continue }
        foreach ($k in (Get-ChildItem $reg -ErrorAction SilentlyContinue)) {
            $dpid = Get-RegValue $k.PSPath 'DigitalProductID'
            $key = ConvertFrom-DigitalProductId $dpid
            if ($key -and $fullKeys -notcontains $key) { $fullKeys += $key }
        }
    }
}

$licenseKeyText = ''
if ($fullKeys.Count -gt 0) {
    $licenseKeyText = ($fullKeys -join '; ')
} elseif ($licenses.Count -gt 0) {
    $licenseKeyText = (($licenses | ForEach-Object { $_.Key }) | Select-Object -Unique) -join '; '
}
$licenseDetail = (($licenses | ForEach-Object { "$($_.Name) [$($_.Status)]" }) | Select-Object -Unique) -join '; '

# ---------------------------------------------------------------- result
function Format-Date([object]$d) {
    if ($d) { return $d.ToString('yyyy/MM/dd HH:mm:ss') } else { return '' }
}

$result = New-Object PSObject
$result | Add-Member NoteProperty ComputerName   $env:COMPUTERNAME
$result | Add-Member NoteProperty IPAddress      (($ipList  | Select-Object -Unique) -join '; ')
$result | Add-Member NoteProperty MACAddress     (($macList | Select-Object -Unique) -join '; ')
$result | Add-Member NoteProperty WindowsVersion $windowsVersion
$result | Add-Member NoteProperty WindowsBuild   "$build"
$result | Add-Member NoteProperty OfficeProduct  (($officeProducts | Select-Object -Unique) -join '; ')
$result | Add-Member NoteProperty OfficeVersion  (($officeVersions | Select-Object -Unique) -join '; ')
$result | Add-Member NoteProperty OfficeLicenseKey    $licenseKeyText
$result | Add-Member NoteProperty OfficeLicenseDetail $licenseDetail
$result | Add-Member NoteProperty LastBootTime   (Format-Date $lastBoot)
$result | Add-Member NoteProperty OSInstallDate  (Format-Date $installDate)
$result | Add-Member NoteProperty EstimatedStartDate (Format-Date $estimatedStart)
$result | Add-Member NoteProperty BiosDate       (Format-Date $biosDate)
$result | Add-Member NoteProperty AssetTag       "$assetTag"
$result | Add-Member NoteProperty SerialNumber   $serial
$result | Add-Member NoteProperty Manufacturer   $maker
$result | Add-Member NoteProperty Model          $model
$result | Add-Member NoteProperty CollectedAt    (Format-Date (Get-Date))
$result | Add-Member NoteProperty CollectedBy    "$env:USERDOMAIN\$env:USERNAME"

if ($OutputDir) {
    # Write to a temp file first, then rename, so the merge script never reads a half-written file.
    # Delete + rename instead of "Move-Item -Force" so overwriting also works on PowerShell 2.0.
    # Exit code: 0 = written, 1 = could not write (see the log written by run-inventory.bat).
    $final = Join-Path $OutputDir "$FileName.csv"
    $tmp   = "$final.tmp"
    try {
        $result | Export-Csv -Path $tmp -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        if (Test-Path -LiteralPath $final) { Remove-Item -LiteralPath $final -Force -ErrorAction Stop }
        Rename-Item -LiteralPath $tmp -NewName (Split-Path $final -Leaf) -ErrorAction Stop
        Write-Host "OK: $final"
        exit 0
    } catch {
        Write-Host "ERROR: could not write $final : $($_.Exception.Message)"
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        exit 1
    }
} else {
    $result
}
