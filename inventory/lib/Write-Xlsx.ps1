<#
  Minimal .xlsx writer with no external dependency (no Excel, no ImportExcel
  module). Requires Windows PowerShell 5.1 or PowerShell 7.

  Usage:
    . .\lib\Write-Xlsx.ps1
    Write-Xlsx -Path out.xlsx -Sheets @(
        @{ Name = 'Sheet1'; Columns = @('A','B'); Rows = $objects;
           RowStyle = { param($row) if ($row.B -eq 'x') { 'warn' } } }  # '', 'warn' or 'alert'
    )

  Every sheet gets: bold header row, frozen header, auto filter, thin borders,
  column widths fitted to the content. Integer values are written as numbers.
#>

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function ConvertTo-XlsxColumnName([int]$Index) {
    # 0 -> A, 25 -> Z, 26 -> AA
    $name = ''
    $n = $Index + 1
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        $name = [string][char](65 + $rem) + $name
        $n = [Math]::Floor(($n - 1) / 26)
    }
    return $name
}

function ConvertTo-XlsxText([object]$Value) {
    $s = "$Value"
    # Drop characters that are not allowed in XML 1.0
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return [System.Security.SecurityElement]::Escape($s)
}

function Get-XlsxDisplayWidth([string]$s) {
    $w = 0
    foreach ($ch in $s.ToCharArray()) { if ([int]$ch -gt 0x2E7F) { $w += 2 } else { $w += 1 } }
    return $w
}

function New-XlsxSheetXml($Sheet) {
    # Style ids (see styles.xml below): 1 header, 2 normal, 3 warn (yellow), 4 alert (pink)
    $styleMap = @{ '' = 2; 'warn' = 3; 'alert' = 4 }
    $cols = @($Sheet.Columns)
    $rows = @($Sheet.Rows)

    $widths = @{}
    for ($c = 0; $c -lt $cols.Count; $c++) { $widths[$c] = Get-XlsxDisplayWidth $cols[$c] }

    $sb = New-Object System.Text.StringBuilder
    $r = 1
    [void]$sb.Append("<row r=`"$r`">")
    for ($c = 0; $c -lt $cols.Count; $c++) {
        $ref = (ConvertTo-XlsxColumnName $c) + $r
        [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`" s=`"1`"><is><t>$(ConvertTo-XlsxText $cols[$c])</t></is></c>")
    }
    [void]$sb.Append('</row>')

    foreach ($row in $rows) {
        $r++
        $style = 2
        if ($Sheet.RowStyle) {
            $key = "$(& $Sheet.RowStyle $row)"
            if ($styleMap.ContainsKey($key)) { $style = $styleMap[$key] }
        }
        [void]$sb.Append("<row r=`"$r`">")
        for ($c = 0; $c -lt $cols.Count; $c++) {
            $ref = (ConvertTo-XlsxColumnName $c) + $r
            $v = $row.($cols[$c])
            $text = "$v"
            $w = Get-XlsxDisplayWidth $text
            if ($w -gt $widths[$c]) { $widths[$c] = $w }
            if ($v -is [int] -or $v -is [long] -or $v -is [double]) {
                [void]$sb.Append("<c r=`"$ref`" s=`"$style`"><v>$v</v></c>")
            } else {
                [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`" s=`"$style`"><is><t xml:space=`"preserve`">$(ConvertTo-XlsxText $text)</t></is></c>")
            }
        }
        [void]$sb.Append('</row>')
    }

    $colXml = ''
    for ($c = 0; $c -lt $cols.Count; $c++) {
        $w = [Math]::Min([Math]::Max($widths[$c] + 2, 8), 60)
        $colXml += "<col min=`"$($c + 1)`" max=`"$($c + 1)`" width=`"$w`" customWidth=`"1`"/>"
    }
    $lastCol = ConvertTo-XlsxColumnName ([Math]::Max($cols.Count - 1, 0))

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        "<cols>$colXml</cols>" +
        "<sheetData>$($sb.ToString())</sheetData>" +
        "<autoFilter ref=`"A1:$lastCol$r`"/>" +
        '</worksheet>'
}

function Write-Xlsx {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object[]]$Sheets
    )

    $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2"><font><sz val="11"/><name val="Meiryo UI"/></font><font><b/><sz val="11"/><name val="Meiryo UI"/></font></fonts>' +
        '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF8CBAD"/></patternFill></fill></fills>' +
        '<borders count="2"><border/><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right>' +
        '<top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="5">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>' +
        '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +
        '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

    $sheetEntries = ''
    $relEntries = ''
    $ctEntries = ''
    $definedNames = ''
    $i = 0
    foreach ($s in $Sheets) {
        $i++
        $sheetEntries += "<sheet name=`"$(ConvertTo-XlsxText $s.Name)`" sheetId=`"$i`" r:id=`"rId$i`"/>"
        $relEntries   += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
        $ctEntries    += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        $lastCol = ConvertTo-XlsxColumnName ([Math]::Max(@($s.Columns).Count - 1, 0))
        $lastRow = @($s.Rows).Count + 1
        $definedNames += "<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$($i - 1)`" hidden=`"1`">'$(ConvertTo-XlsxText $s.Name)'!`$A`$1:`$$lastCol`$$lastRow</definedName>"
    }
    $stylesRelId = $i + 1

    $files = [ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
            "$ctEntries</Types>"
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '</Relationships>'
        'xl/workbook.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            "<sheets>$sheetEntries</sheets><definedNames>$definedNames</definedNames></workbook>"
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            "$relEntries<Relationship Id=`"rId$stylesRelId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>" +
            '</Relationships>'
        'xl/styles.xml' = $styles
    }
    $i = 0
    foreach ($s in $Sheets) {
        $i++
        $files["xl/worksheets/sheet$i.xml"] = New-XlsxSheetXml $s
    }

    $full = [System.IO.Path]::GetFullPath($Path)
    if (Test-Path $full) { Remove-Item $full -Force }
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $zip = [System.IO.Compression.ZipFile]::Open($full, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $files.Keys) {
            $entry = $zip.CreateEntry($name)
            $w = New-Object System.IO.StreamWriter($entry.Open(), $utf8)
            $w.Write($files[$name])
            $w.Close()
        }
    } finally {
        $zip.Dispose()
    }
}
