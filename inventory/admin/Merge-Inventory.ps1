<#
.SYNOPSIS
  収集した端末ごとの CSV を 1 つの Excel 一覧 (.xlsx) にまとめ、台帳.csv を更新する。

.DESCRIPTION
  - RawDir 内の <端末名>.csv (Get-DeviceInventory.ps1 の出力) をすべて読み込む
  - 台帳.csv (LedgerFile) を シリアル番号 → 端末名 の順で突合する
      * 台帳に無い端末は自動で行を追加 (タグNo は BIOS/レジストリにあれば自動入力)
      * 台帳の「タグNo」「導入日」「設置場所」「備考」は担当者が手入力する欄で、上書きしない
  - 端末稼働日 = 台帳の導入日 (未入力なら端末から推定した導入日)
  - Discovery (Find-NetworkHosts.ps1 の出力) があれば、ネットワーク上に居るのに
    未収集の機器も一覧に載せる
  - 出力シート: 端末一覧 / 未収集 / 集計
  Excel も外部モジュールも不要 (Windows PowerShell 5.1 以降)。

  ※ このファイルは日本語を含むため UTF-8 (BOM 付き) のまま保存すること。

.EXAMPLE
  .\Merge-Inventory.ps1 -RawDir \\fs01\inventory$\raw -LedgerFile \\fs01\inventory$\台帳.csv -OutFile .\端末一覧.xlsx
#>
param(
    [Parameter(Mandatory = $true)][string]$RawDir,
    [Parameter(Mandatory = $true)][string]$LedgerFile,
    [Parameter(Mandatory = $true)][string]$OutFile,
    [string]$Discovery,
    [int]$StaleDays = 30,
    [datetime]$Today = (Get-Date)
)

$ErrorActionPreference = 'Stop'
. (Join-Path (Join-Path $PSScriptRoot 'lib') 'Write-Xlsx.ps1')

$LedgerColumns = @('タグNo', '端末名', 'シリアル番号', 'MACアドレス', '導入日', '設置場所', '備考', '初回検出日', '最終検出日')

# ------------------------------------------------------------ helpers
# UTF-8 (BOM 有無どちらも) と Shift-JIS (Excel で「CSV」保存した場合) の両方を読む
function Import-CsvAuto([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    try {
        # 厳密な UTF-8 として読めなければ Shift-JIS とみなす
        $text = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes)
    } catch {
        $text = [System.Text.Encoding]::GetEncoding(932).GetString($bytes)
    }
    return @($text.TrimStart([char]0xFEFF) | ConvertFrom-Csv)
}

function Export-CsvUtf8Bom($Rows, [string[]]$Columns, [string]$Path) {
    $lines = @($Rows | Select-Object -Property $Columns | ConvertTo-Csv -NoTypeInformation)
    if ($lines.Count -eq 0) { $lines = @(($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') }
    [System.IO.File]::WriteAllLines($Path, [string[]]$lines, (New-Object System.Text.UTF8Encoding($true)))
}

function Get-Norm([string]$s) { return ("$s".Trim().ToUpper() -replace ':', '-') }

function Split-Multi([string]$s) {
    return @("$s" -split ';' | ForEach-Object { Get-Norm $_ } | Where-Object { $_ })
}

# メーカー未設定などで全機種共通になっているシリアルは照合に使わない
function Test-UsableSerial([string]$s) {
    $n = Get-Norm $s
    if ($n.Length -lt 4) { return $false }
    return ($n -notmatch '^(TO BE FILLED BY O\.E\.M\.|SYSTEM SERIAL NUMBER|DEFAULT STRING|NONE|N/A|NOT SPECIFIED|INVALID|0+|1234567890|CHASSIS SERIAL NUMBER)$')
}

$DateFormats = [string[]]@('yyyy/M/d H:mm:ss', 'yyyy/M/d H:mm', 'yyyy/M/d', 'yyyy-M-d H:mm:ss', 'yyyy-M-d', 'yyyy.M.d', 'yyyyMMdd')
function ConvertTo-DateOrNull([string]$s) {
    $s = "$s".Trim()
    if (-not $s) { return $null }
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, $DateFormats, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    if ([datetime]::TryParse($s, [ref]$d)) { return $d }
    return $null
}

function Format-Day($d) { if ($d) { return $d.ToString('yyyy/MM/dd') } else { return '' } }

function Get-MonthsBetween([datetime]$from, [datetime]$to) {
    $m = ($to.Year - $from.Year) * 12 + ($to.Month - $from.Month)
    if ($to.Day -lt $from.Day) { $m-- }
    return [int][Math]::Max($m, 0)
}

# ------------------------------------------------------------ 台帳の読み込み
$ledger = New-Object System.Collections.ArrayList
if (Test-Path $LedgerFile) {
    foreach ($row in (Import-CsvAuto $LedgerFile)) {
        $o = New-Object PSObject
        foreach ($c in $LedgerColumns) { $o | Add-Member NoteProperty $c "$($row.$c)".Trim() }
        $o | Add-Member NoteProperty '__id' $ledger.Count
        [void]$ledger.Add($o)
    }
}
$bySerial = @{}; $byName = @{}
foreach ($l in $ledger) {
    if (Test-UsableSerial $l.'シリアル番号') { $bySerial[(Get-Norm $l.'シリアル番号')] = $l }
    if ($l.'端末名') { $byName[(Get-Norm $l.'端末名')] = $l }
}
$matched = @{}   # 今回収集できた台帳行 (__id)

# ------------------------------------------------------------ 収集データ
$OutputColumns = @('タグNo', '端末名', 'IPアドレス', 'MACアドレス', 'Windowsバージョン', 'Officeバージョン',
    'Officeライセンスキー(2013以降は末尾5桁)', '起動日', '端末稼働日(導入日)', '端末稼働月数', '導入日の根拠', '設置場所',
    'ライセンス状態', 'メーカー', '機種', 'シリアル番号', 'BIOS日付', '最終収集日時', '状態')

function New-OutputRow {
    $o = New-Object PSObject
    foreach ($c in $OutputColumns) { $o | Add-Member NoteProperty $c '' }
    return $o
}

$todayText = Format-Day $Today
$rows = New-Object System.Collections.ArrayList
$seenIp = @{}; $seenMac = @{}

# 同じ端末の CSV が複数あれば (改名など) 新しい方を採用
$raw = @{}
foreach ($f in (Get-ChildItem -Path $RawDir -Filter '*.csv' | Where-Object { $_.Name -notlike '_*' })) {
    foreach ($r in (Import-CsvAuto $f.FullName)) {
        if (-not $r.ComputerName) { continue }
        $key = if (Test-UsableSerial $r.SerialNumber) { 'S:' + (Get-Norm $r.SerialNumber) } else { 'N:' + (Get-Norm $r.ComputerName) }
        $prev = $raw[$key]
        if (-not $prev -or (ConvertTo-DateOrNull $r.CollectedAt) -gt (ConvertTo-DateOrNull $prev.CollectedAt)) { $raw[$key] = $r }
    }
}

foreach ($r in $raw.Values) {
    # 台帳との突合: シリアル番号 → 端末名
    $l = $null
    if (Test-UsableSerial $r.SerialNumber) { $l = $bySerial[(Get-Norm $r.SerialNumber)] }
    if (-not $l) { $l = $byName[(Get-Norm $r.ComputerName)] }
    if (-not $l) {
        $l = New-Object PSObject
        foreach ($c in $LedgerColumns) { $l | Add-Member NoteProperty $c '' }
        $l.'タグNo' = "$($r.AssetTag)"
        $l.'初回検出日' = $todayText
        $l | Add-Member NoteProperty '__id' $ledger.Count
        [void]$ledger.Add($l)
        if (Test-UsableSerial $r.SerialNumber) { $bySerial[(Get-Norm $r.SerialNumber)] = $l }
    }
    # 自動で更新する欄 (手入力欄は触らない)
    $l.'端末名' = $r.ComputerName
    $l.'シリアル番号' = $r.SerialNumber
    $l.'MACアドレス' = $r.MACAddress
    $l.'最終検出日' = Format-Day (ConvertTo-DateOrNull $r.CollectedAt)
    if (-not $l.'初回検出日') { $l.'初回検出日' = $todayText }
    $byName[(Get-Norm $r.ComputerName)] = $l
    $matched[$l.'__id'] = $true

    foreach ($ip in (Split-Multi $r.IPAddress))   { $seenIp[$ip] = $true }
    foreach ($mac in (Split-Multi $r.MACAddress)) { $seenMac[$mac] = $true }

    $start = ConvertTo-DateOrNull $l.'導入日'
    $basis = '台帳'
    if (-not $start) {
        $start = ConvertTo-DateOrNull $r.EstimatedStartDate
        if (-not $start) { $start = ConvertTo-DateOrNull $r.OSInstallDate }
        $basis = if ($start) { '推定(端末)' } else { '' }
    }
    $boot = ConvertTo-DateOrNull $r.LastBootTime
    $collected = ConvertTo-DateOrNull $r.CollectedAt

    $o = New-OutputRow
    $o.'タグNo' = $l.'タグNo'
    $o.'端末名' = $r.ComputerName
    $o.'IPアドレス' = $r.IPAddress
    $o.'MACアドレス' = $r.MACAddress
    $o.'Windowsバージョン' = $r.WindowsVersion
    $o.'Officeバージョン' = (@($r.OfficeProduct, $r.OfficeVersion) | Where-Object { $_ }) -join ' / '
    $o.'Officeライセンスキー(2013以降は末尾5桁)' = $r.OfficeLicenseKey
    $o.'起動日' = if ($boot) { $boot.ToString('yyyy/MM/dd HH:mm') } else { '' }
    $o.'端末稼働日(導入日)' = Format-Day $start
    $o.'端末稼働月数' = if ($start) { Get-MonthsBetween $start $Today } else { '' }
    $o.'導入日の根拠' = $basis
    $o.'設置場所' = $l.'設置場所'
    $o.'ライセンス状態' = $r.OfficeLicenseDetail
    $o.'メーカー' = $r.Manufacturer
    $o.'機種' = $r.Model
    $o.'シリアル番号' = $r.SerialNumber
    $o.'BIOS日付' = Format-Day (ConvertTo-DateOrNull $r.BiosDate)
    $o.'最終収集日時' = $r.CollectedAt
    $o.'状態' = if ($collected -and ($Today - $collected).TotalDays -gt $StaleDays) { "収集済(${StaleDays}日以上前)" } else { '収集済' }
    [void]$rows.Add($o)
}

# ------------------------------------------------------------ 未収集
$missing = New-Object System.Collections.ArrayList
foreach ($l in $ledger) {
    if ($matched.ContainsKey($l.'__id')) { continue }
    $o = New-OutputRow
    $o.'タグNo' = $l.'タグNo'; $o.'端末名' = $l.'端末名'; $o.'MACアドレス' = $l.'MACアドレス'
    $o.'シリアル番号' = $l.'シリアル番号'; $o.'設置場所' = $l.'設置場所'; $o.'最終収集日時' = $l.'最終検出日'
    $start = ConvertTo-DateOrNull $l.'導入日'
    if ($start) { $o.'端末稼働日(導入日)' = Format-Day $start; $o.'端末稼働月数' = Get-MonthsBetween $start $Today; $o.'導入日の根拠' = '台帳' }
    $o.'状態' = '未収集(台帳のみ)'
    [void]$missing.Add($o)
}
if ($Discovery) {
    $macToLedger = @{}
    foreach ($l in $ledger) { foreach ($m in (Split-Multi $l.'MACアドレス')) { $macToLedger[$m] = $l } }
    foreach ($d in (Import-CsvAuto $Discovery)) {
        $ip = Get-Norm $d.IPAddress; $mac = Get-Norm $d.MACAddress
        if ($seenIp.ContainsKey($ip) -or ($mac -and $seenMac.ContainsKey($mac))) { continue }
        $o = New-OutputRow
        $o.'IPアドレス' = $d.IPAddress; $o.'MACアドレス' = $d.MACAddress; $o.'端末名' = $d.HostName
        if ($mac -and $macToLedger.ContainsKey($mac)) { $o.'タグNo' = $macToLedger[$mac].'タグNo'; $o.'設置場所' = $macToLedger[$mac].'設置場所' }
        $o.'状態' = '未収集(ネットワーク検出のみ)'
        [void]$missing.Add($o)
    }
}

$allRows = @($rows | Sort-Object '端末名') + @($missing | Sort-Object '状態', '端末名')

# ------------------------------------------------------------ 集計
$summary = New-Object System.Collections.ArrayList
function Add-Summary([string]$cat, [string]$item, [int]$n) {
    $o = New-Object PSObject
    $o | Add-Member NoteProperty '区分' $cat
    $o | Add-Member NoteProperty '項目' $item
    $o | Add-Member NoteProperty '台数' $n
    [void]$summary.Add($o)
}
Add-Summary '全体' '収集済' $rows.Count
Add-Summary '全体' '未収集(台帳のみ)' @($missing | Where-Object { $_.'状態' -eq '未収集(台帳のみ)' }).Count
Add-Summary '全体' '未収集(ネットワーク検出のみ)' @($missing | Where-Object { $_.'状態' -eq '未収集(ネットワーク検出のみ)' }).Count
foreach ($g in ($rows | Group-Object { ($_.'Windowsバージョン' -replace '\s*\([^)]*\)\s*$', '') } | Sort-Object Name)) {
    Add-Summary 'Windows' $(if ($g.Name) { $g.Name } else { '(不明)' }) $g.Count
}
foreach ($g in ($rows | Group-Object { ($_.'Officeバージョン' -split ' / ')[0] } | Sort-Object Name)) {
    Add-Summary 'Office' $(if ($g.Name) { $g.Name } else { '(なし)' }) $g.Count
}
$bands = [ordered]@{ '〜36か月' = 0; '37〜60か月' = 0; '61か月〜' = 0; '不明' = 0 }
foreach ($r in $rows) {
    $m = $r.'端末稼働月数'
    if ($m -is [int]) {
        if ($m -le 36) { $bands['〜36か月']++ } elseif ($m -le 60) { $bands['37〜60か月']++ } else { $bands['61か月〜']++ }
    } else { $bands['不明']++ }
}
foreach ($k in $bands.Keys) { Add-Summary '端末稼働月数' $k $bands[$k] }

# ------------------------------------------------------------ 出力
$rowStyle = {
    param($row)
    if ($row.'状態' -like '未収集*') { 'alert' } elseif ($row.'状態' -ne '収集済') { 'warn' } else { '' }
}
Write-Xlsx -Path $OutFile -Sheets @(
    @{ Name = '端末一覧'; Columns = $OutputColumns; Rows = $allRows; RowStyle = $rowStyle },
    @{ Name = '未収集';   Columns = $OutputColumns; Rows = @($missing | Sort-Object '状態', '端末名'); RowStyle = $rowStyle },
    @{ Name = '集計';     Columns = @('区分', '項目', '台数'); Rows = @($summary) }
)

# 台帳.csv を更新 (前回分は .bak に残す)
try {
    if (Test-Path $LedgerFile) { Copy-Item -Path $LedgerFile -Destination "$LedgerFile.bak" -Force }
    Export-CsvUtf8Bom ($ledger | Sort-Object 'タグNo', '端末名') $LedgerColumns $LedgerFile
} catch {
    Write-Warning "台帳.csv を更新できませんでした (Excel で開いたままになっていませんか？): $($_.Exception.Message)"
}

Write-Host ("出力: {0}  収集済 {1} 台 / 未収集 {2} 件 / 台帳 {3} 行" -f $OutFile, $rows.Count, $missing.Count, $ledger.Count)
