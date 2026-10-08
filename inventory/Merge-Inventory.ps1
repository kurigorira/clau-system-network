<#
.SYNOPSIS
  収集した端末ごとの CSV を 1 つの一覧 (Excel で開ける CSV) にまとめる。

.DESCRIPTION
  - RawDir 内の <端末名>.csv (Get-DeviceInventory.ps1 の出力) をすべて結合
  - AssetMaster (資産台帳 CSV) があれば シリアル番号 → MAC → 端末名 の順で突合し、
    タグNo・導入日・設置場所を補完
  - Discovery (Find-NetworkHosts.ps1 の出力) があれば、ネットワーク上に居るのに
    未収集の機器を「未収集」行として追加
  - 稼働開始日 = 台帳の導入日 (無ければ OS インストール日)、稼働月数は本日時点で計算

  ※ このファイルは日本語を含むため UTF-8 (BOM 付き) で保存しておくこと。

.EXAMPLE
  .\Merge-Inventory.ps1 -RawDir \\fs01\inventory$\raw -AssetMaster .\asset-master.csv -Discovery .\discovery.csv -OutFile .\端末一覧.csv
#>
param(
    [Parameter(Mandatory = $true)][string]$RawDir,
    [string]$AssetMaster,
    [string]$Discovery,
    [Parameter(Mandatory = $true)][string]$OutFile,
    [int]$StaleDays = 30
)

$today = Get-Date

# BOM 付きなら UTF-8、無ければ Shift-JIS (Excel で保存した CSV) として読む
function Import-CsvAuto([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return Import-Csv -Path $Path -Encoding UTF8
    }
    return Import-Csv -Path $Path -Encoding Default
}

function Get-Norm([string]$s) { return ("$s".Trim().ToUpper() -replace ':', '-') }

function Split-Multi([string]$s) {
    return @("$s" -split ';' | ForEach-Object { Get-Norm $_ } | Where-Object { $_ })
}

function ConvertTo-DateOrNull([string]$s) {
    if (-not "$s".Trim()) { return $null }
    $d = [datetime]::MinValue
    if ([datetime]::TryParse($s, [ref]$d)) { return $d }
    return $null
}

function Get-MonthsBetween([datetime]$from, [datetime]$to) {
    $m = ($to.Year - $from.Year) * 12 + ($to.Month - $from.Month)
    if ($to.Day -lt $from.Day) { $m-- }
    return [Math]::Max($m, 0)
}

# ------------------------------------------------------------ 資産台帳
$bySerial = @{}; $byMac = @{}; $byName = @{}
$masterRows = @()
if ($AssetMaster) {
    $masterRows = @(Import-CsvAuto $AssetMaster)
    $i = 0
    foreach ($m in $masterRows) {
        $m | Add-Member -NotePropertyName '__idx' -NotePropertyValue ($i++)
        if ($m.'シリアル番号') { $bySerial[(Get-Norm $m.'シリアル番号')] = $m }
        foreach ($mac in (Split-Multi $m.'MACアドレス')) { $byMac[$mac] = $m }
        if ($m.'端末名') { $byName[(Get-Norm $m.'端末名')] = $m }
    }
}
$usedMaster = @{}

# ------------------------------------------------------------ 収集データ
$rows = @()
$seenIp = @{}; $seenMac = @{}
foreach ($f in (Get-ChildItem -Path $RawDir -Filter '*.csv')) {
    foreach ($r in (Import-Csv -Path $f.FullName -Encoding UTF8)) {
        $master = $null
        $serialKey = Get-Norm $r.SerialNumber
        if ($serialKey -and $bySerial.ContainsKey($serialKey)) { $master = $bySerial[$serialKey] }
        if (-not $master) {
            foreach ($mac in (Split-Multi $r.MACAddress)) { if ($byMac.ContainsKey($mac)) { $master = $byMac[$mac]; break } }
        }
        if (-not $master -and $byName.ContainsKey((Get-Norm $r.ComputerName))) { $master = $byName[(Get-Norm $r.ComputerName)] }
        if ($master) { $usedMaster[$master.'__idx'] = $true }

        foreach ($ip in (Split-Multi $r.IPAddress))   { $seenIp[$ip] = $true }
        foreach ($mac in (Split-Multi $r.MACAddress)) { $seenMac[$mac] = $true }

        # タグNo: 台帳 > 端末側 (レジストリ / BIOS)
        $tag = ''
        if ($master -and $master.'タグNo') { $tag = $master.'タグNo' } elseif ($r.AssetTag) { $tag = $r.AssetTag }

        # 稼働開始日: 台帳の導入日 > OS インストール日
        $start = $null; $startSrc = ''
        if ($master) { $start = ConvertTo-DateOrNull $master.'導入日'; if ($start) { $startSrc = '資産台帳' } }
        if (-not $start) { $start = ConvertTo-DateOrNull $r.OSInstallDate; if ($start) { $startSrc = 'OSインストール日' } }

        $boot = ConvertTo-DateOrNull $r.LastBootTime
        $collected = ConvertTo-DateOrNull $r.CollectedAt
        $status = '収集済'
        if ($collected -and ($today - $collected).TotalDays -gt $StaleDays) { $status = "収集済(${StaleDays}日以上前)" }

        $rows += [pscustomobject][ordered]@{
            'タグNo'               = $tag
            '端末名'               = $r.ComputerName
            'IPアドレス'           = $r.IPAddress
            'MACアドレス'          = $r.MACAddress
            'Windowsバージョン'    = $r.WindowsVersion
            'Officeバージョン'     = (@($r.OfficeProduct, $r.OfficeVersion) | Where-Object { $_ }) -join ' / '
            'Officeライセンスキー' = $r.OfficeLicenseKey
            '起動日'               = $(if ($boot) { $boot.ToString('yyyy/MM/dd HH:mm') } else { '' })
            '連続稼働日数'         = $(if ($boot) { [Math]::Floor(($today - $boot).TotalDays) } else { '' })
            '端末稼働日'           = $(if ($start) { $start.ToString('yyyy/MM/dd') } else { '' })
            '端末稼働月数'         = $(if ($start) { Get-MonthsBetween $start $today } else { '' })
            '稼働日の根拠'         = $startSrc
            'ライセンス状態'       = $r.OfficeLicenseDetail
            'Windowsビルド'        = $r.WindowsBuild
            'メーカー'             = $r.Manufacturer
            '機種'                 = $r.Model
            'シリアル番号'         = $r.SerialNumber
            '設置場所'             = $(if ($master) { $master.'設置場所' } else { '' })
            '最終収集日時'         = $r.CollectedAt
            '状態'                 = $status
        }
    }
}

function New-EmptyRow([string]$status) {
    return [pscustomobject][ordered]@{
        'タグNo' = ''; '端末名' = ''; 'IPアドレス' = ''; 'MACアドレス' = ''; 'Windowsバージョン' = ''
        'Officeバージョン' = ''; 'Officeライセンスキー' = ''; '起動日' = ''; '連続稼働日数' = ''
        '端末稼働日' = ''; '端末稼働月数' = ''; '稼働日の根拠' = ''; 'ライセンス状態' = ''; 'Windowsビルド' = ''
        'メーカー' = ''; '機種' = ''; 'シリアル番号' = ''; '設置場所' = ''; '最終収集日時' = ''; '状態' = $status
    }
}

# ------------------------------------------------------------ 台帳にあるが未収集
foreach ($m in $masterRows) {
    if ($usedMaster.ContainsKey($m.'__idx')) { continue }
    $x = New-EmptyRow '未収集(台帳のみ)'
    $x.'タグNo' = $m.'タグNo'; $x.'端末名' = $m.'端末名'; $x.'MACアドレス' = $m.'MACアドレス'
    $x.'シリアル番号' = $m.'シリアル番号'; $x.'設置場所' = $m.'設置場所'
    $start = ConvertTo-DateOrNull $m.'導入日'
    if ($start) {
        $x.'端末稼働日' = $start.ToString('yyyy/MM/dd'); $x.'端末稼働月数' = Get-MonthsBetween $start $today; $x.'稼働日の根拠' = '資産台帳'
    }
    $rows += $x
}

# ------------------------------------------------------------ ネットワーク上にあるが未収集
if ($Discovery) {
    foreach ($d in (Import-CsvAuto $Discovery)) {
        $ip = Get-Norm $d.IPAddress; $mac = Get-Norm $d.MACAddress
        if ($seenIp.ContainsKey($ip) -or ($mac -and $seenMac.ContainsKey($mac))) { continue }
        $x = New-EmptyRow '未収集(ネットワーク検出のみ)'
        $x.'IPアドレス' = $d.IPAddress; $x.'MACアドレス' = $d.MACAddress; $x.'端末名' = $d.HostName
        if ($mac -and $byMac.ContainsKey($mac)) {
            $m = $byMac[$mac]; $x.'タグNo' = $m.'タグNo'; $x.'設置場所' = $m.'設置場所'
        }
        $rows += $x
    }
}

# Excel で文字化けしないよう UTF-8 (BOM 付き) で出力
$rows | Sort-Object '状態', '端末名' | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
Write-Host ("出力: {0}  (収集済 {1} 台 / 未収集 {2} 件)" -f $OutFile,
    @($rows | Where-Object { $_.'状態' -like '収集済*' }).Count,
    @($rows | Where-Object { $_.'状態' -like '未収集*' }).Count)
