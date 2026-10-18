<#
.SYNOPSIS
    Convert the CSV files produced by jboss_pool_collector.sh into an Excel workbook (.xlsx).

.DESCRIPTION
    - Font        : Meiryo UI (all cells, charts, print header/footer)
    - Colors      : monotone (black / grays). EXHAUSTED = black, WAIT = dark gray, HIGH = light gray
    - Cell setup  : number / date formats, borders, zebra rows, frozen panes, auto filter,
                    column widths, print setup (landscape, repeated header rows)
    - Sheets      : Summary / Charts / Hourly / Events / Time series / Queued threads /
                    Queued stacks / Thread dump / Pool config / Column descriptions
                    (Japanese sheet names and labels are read from columns.csv)
    - Excel itself is NOT required: the .xlsx (Office Open XML) is generated directly.
    - Runs on Windows PowerShell 5.1 (Windows 11 default) and PowerShell 7.x.
    - This file is intentionally ASCII-only, so it works no matter which encoding it is
      saved with (e.g. pasted into Notepad and saved as UTF-8 without BOM).

.PARAMETER InputDir
    Folder that contains pool_timeseries.csv and the other CSV files.

.PARAMETER OutputPath
    Output .xlsx file. Default: <InputDir>\<folder name>.xlsx

.PARAMETER MaxRows
    Maximum number of rows of the time series sheet (default 1048575 = Excel limit).

.PARAMETER NoCharts
    Do not create the chart sheet.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Convert-PoolCsvToXlsx.ps1 -InputDir C:\work\pool_report_task-1234_20261018_100000
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputDir,
    [string]$OutputPath = '',
    [ValidateRange(1, 1048575)][int]$MaxRows = 1048575,
    [switch]$NoCharts
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName Microsoft.VisualBasic

$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)
$script:FONT = 'Meiryo UI'
$script:NumRe = New-Object System.Text.RegularExpressions.Regex('^-?\d+(\.\d+)?([eE][-+]?\d+)?$')
$script:DtRe = New-Object System.Text.RegularExpressions.Regex('^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$')
$script:CtrlRe = New-Object System.Text.RegularExpressions.Regex('[\x00-\x08\x0B\x0C\x0E-\x1F]')
$script:CellRe = New-Object System.Text.RegularExpressions.Regex('^([A-Z]+)(\d+)$')
$script:Epoch = New-Object DateTime(1899, 12, 30)
$script:SQUARE = [string][char]0x25A0
$script:SEP = [string][char]0x1F
$script:NL = "`n"

if (-not (Test-Path -LiteralPath $InputDir -PathType Container)) { throw "Folder not found: $InputDir" }
$script:IN_DIR = (Resolve-Path -LiteralPath $InputDir).ProviderPath
if (-not (Test-Path -LiteralPath (Join-Path $script:IN_DIR 'pool_timeseries.csv') -PathType Leaf)) {
    throw "pool_timeseries.csv not found in $script:IN_DIR"
}
if ([string]::IsNullOrEmpty($OutputPath)) {
    $OutputPath = Join-Path $script:IN_DIR ((Split-Path -Leaf $script:IN_DIR) + '.xlsx')
}
$script:OUT_PATH = [System.IO.Path]::GetFullPath($OutputPath)
$script:MAX_ROWS = $MaxRows
$script:CHARTS = -not $NoCharts.IsPresent

# ------------------------------------------------------------------------------
# helpers
# ------------------------------------------------------------------------------
function New-Map { return (New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal)) }

function New-List { return , (New-Object System.Collections.ArrayList) }

function Get-CsvPath([string]$name) { return (Join-Path $script:IN_DIR $name) }

function Open-CsvParser([string]$path) {
    $p = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($path, [System.Text.Encoding]::UTF8, $true)
    $p.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $p.SetDelimiters(',')
    $p.HasFieldsEnclosedInQuotes = $true
    $p.TrimWhiteSpace = $false
    return $p
}

function Read-CsvFile([string]$name) {
    $path = Get-CsvPath $name
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    $rows = New-Object System.Collections.ArrayList
    $p = Open-CsvParser $path
    try {
        while (-not $p.EndOfData) {
            $f = $p.ReadFields()
            if ($null -ne $f) { [void]$rows.Add([string[]]$f) }
        }
    } finally { $p.Close() }
    return , $rows
}

function Get-Field([string[]]$row, [int]$i) {
    if ($i -lt $row.Length) { return $row[$i] }
    return ''
}

function Test-Num([string]$v) { return $script:NumRe.IsMatch($v) }
function To-Double([string]$v) { return [double]::Parse($v, [System.Globalization.NumberStyles]::Float, $script:Inv) }
function F8([double]$v) { return $v.ToString('F8', $script:Inv) }

# ------------------------------------------------------------------------------
# input: column definitions / labels (Japanese text comes from columns.csv)
# ------------------------------------------------------------------------------
$script:META = New-Map
$script:STATUS = New-Map
$script:UI = New-Map
$script:COLS_ROWS = Read-CsvFile 'columns.csv'
if ($null -eq $script:COLS_ROWS) { $script:COLS_ROWS = New-Object System.Collections.ArrayList }
for ($i = 1; $i -lt $script:COLS_ROWS.Count; $i++) {
    $r = $script:COLS_ROWS[$i]
    if ($r.Length -lt 6) { continue }
    if ($r[0] -ceq 'ui') { $script:UI[$r[1]] = $r[2] }
    elseif ($r[0] -ceq 'status') { $script:STATUS[$r[1]] = @($r[2], $r[5]) }
    else { $script:META[$r[0] + "`t" + $r[1]] = @($r[2], $r[3], $r[4], $r[5]) }
}

function UI([string]$key) {
    if ($script:UI.ContainsKey($key)) { return $script:UI[$key] }
    return $key
}

function Get-Meta([string]$f, [string]$k) {
    $mk = $f + "`t" + $k
    if ($script:META.ContainsKey($mk)) { return $script:META[$mk] }
    return @($k, '', 'text', '')
}

$script:INFO = New-Map
$script:INFO_ROWS = New-Object System.Collections.ArrayList
$infoCsv = Read-CsvFile 'collection_info.csv'
if ($null -ne $infoCsv) {
    for ($i = 1; $i -lt $infoCsv.Count; $i++) {
        $r = $infoCsv[$i]
        if ($r.Length -ge 3) { $script:INFO[$r[0]] = $r[2]; [void]$script:INFO_ROWS.Add($r) }
    }
}
$script:WARN_PCT = 80.0
if ($script:INFO.ContainsKey('warn_pct')) {
    $tmp = 0.0
    if ([double]::TryParse($script:INFO['warn_pct'], [System.Globalization.NumberStyles]::Float, $script:Inv, [ref]$tmp)) { $script:WARN_PCT = $tmp }
}

# ------------------------------------------------------------------------------
# styles (fixed table, identical to pool_csv_to_xlsx.sh)
# ------------------------------------------------------------------------------
$script:FONTS = @(@(10, 0, 0, '000000'), @(10, 1, 0, '000000'), @(10, 1, 0, 'FFFFFF'), @(16, 1, 0, '000000'),
    @(9, 0, 0, '595959'), @(11, 1, 0, '000000'), @(10, 0, 1, '808080'))
$script:FILLS = @('262626', 'F2F2F2', '000000', '595959', 'BFBFBF', 'D9D9D9')
$script:BORDERS = @($null, @('thin', 'BFBFBF'), @('thin', '808080'), @('bottom', '000000'))
$script:NUMFMTS = @(@(164, '#,##0.0'), @(165, '0.0"%"'), @(166, 'yyyy/mm/dd hh:mm:ss'), @(167, 'yyyy/mm/dd hh:mm'))
# name, font, fill, border, numFmt, horizontal, vertical, wrap
$script:XF_DEFS = @(
    @('default', 0, 0, 0, 0, '', '', 0),
    @('title', 3, 0, 0, 0, '', 'center', 0),
    @('subtitle', 4, 0, 0, 0, '', 'center', 0),
    @('section', 5, 0, 3, 0, '', 'center', 0),
    @('hdr', 2, 2, 2, 0, 'center', 'center', 1),
    @('lbl', 1, 3, 1, 0, 'left', 'center', 1),
    @('text', 0, 0, 1, 0, '', 'center', 0),
    @('text_z', 0, 3, 1, 0, '', 'center', 0),
    @('int', 0, 0, 1, 3, '', 'center', 0),
    @('int_z', 0, 3, 1, 3, '', 'center', 0),
    @('dec1', 0, 0, 1, 164, '', 'center', 0),
    @('dec1_z', 0, 3, 1, 164, '', 'center', 0),
    @('pct', 0, 0, 1, 165, '', 'center', 0),
    @('pct_z', 0, 3, 1, 165, '', 'center', 0),
    @('pctb', 1, 0, 1, 165, '', 'center', 0),
    @('pctb_z', 1, 3, 1, 165, '', 'center', 0),
    @('dt', 0, 0, 1, 166, 'center', 'center', 0),
    @('dt_z', 0, 3, 1, 166, 'center', 'center', 0),
    @('dtm', 0, 0, 1, 167, 'center', 'center', 0),
    @('dtm_z', 0, 3, 1, 167, 'center', 'center', 0),
    @('wrap', 0, 0, 1, 0, 'left', 'top', 1),
    @('wrap_z', 0, 3, 1, 0, 'left', 'top', 1),
    @('st_OK', 0, 0, 1, 0, 'center', 'center', 0),
    @('st_OK_z', 0, 3, 1, 0, 'center', 'center', 0),
    @('st_HIGH', 1, 6, 1, 0, 'center', 'center', 0),
    @('st_WAIT', 2, 5, 1, 0, 'center', 'center', 0),
    @('st_EXHAUSTED', 2, 4, 1, 0, 'center', 'center', 0),
    @('st_GRAY', 6, 0, 1, 0, 'center', 'center', 0),
    @('st_GRAY_z', 6, 3, 1, 0, 'center', 'center', 0),
    @('note', 6, 0, 0, 0, '', 'center', 0),
    @('textb', 1, 0, 1, 0, '', 'center', 0),
    @('textb_z', 1, 3, 1, 0, '', 'center', 0),
    @('textw', 0, 0, 1, 0, 'left', 'center', 1)
)
$script:XF = New-Map
for ($i = 0; $i -lt $script:XF_DEFS.Length; $i++) { $script:XF[$script:XF_DEFS[$i][0]] = $i }

function Get-StylesXml {
    $o = New-Object System.Text.StringBuilder
    [void]$o.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL)
    [void]$o.Append('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$o.Append('<numFmts count="' + $script:NUMFMTS.Length + '">')
    foreach ($nf in $script:NUMFMTS) { [void]$o.Append('<numFmt numFmtId="' + $nf[0] + '" formatCode="' + (EscAttr $nf[1]) + '"/>') }
    [void]$o.Append('</numFmts><fonts count="' + $script:FONTS.Length + '">')
    foreach ($f in $script:FONTS) {
        $b = ''; if ($f[1]) { $b = '<b/>' }
        $it = ''; if ($f[2]) { $it = '<i/>' }
        [void]$o.Append('<font>' + $b + $it + '<sz val="' + $f[0] + '"/><color rgb="FF' + $f[3] + '"/><name val="' + $script:FONT + '"/><family val="3"/><charset val="128"/></font>')
    }
    [void]$o.Append('</fonts><fills count="' + ($script:FILLS.Length + 2) + '"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>')
    foreach ($c in $script:FILLS) { [void]$o.Append('<fill><patternFill patternType="solid"><fgColor rgb="FF' + $c + '"/><bgColor indexed="64"/></patternFill></fill>') }
    [void]$o.Append('</fills><borders count="' + $script:BORDERS.Length + '">')
    foreach ($b in $script:BORDERS) {
        if ($null -eq $b) { [void]$o.Append('<border><left/><right/><top/><bottom/><diagonal/></border>') }
        elseif ($b[0] -eq 'bottom') { [void]$o.Append('<border><left/><right/><top/><bottom style="medium"><color rgb="FF' + $b[1] + '"/></bottom><diagonal/></border>') }
        else {
            $s = '<border>'
            foreach ($e in @('left', 'right', 'top', 'bottom')) { $s += '<' + $e + ' style="thin"><color rgb="FF' + $b[1] + '"/></' + $e + '>' }
            [void]$o.Append($s + '<diagonal/></border>')
        }
    }
    [void]$o.Append('</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
    [void]$o.Append('<cellXfs count="' + $script:XF_DEFS.Length + '">')
    foreach ($d in $script:XF_DEFS) {
        $a = '<xf numFmtId="' + $d[4] + '" fontId="' + $d[1] + '" fillId="' + $d[2] + '" borderId="' + $d[3] + '" xfId="0"'
        if ($d[4]) { $a += ' applyNumberFormat="1"' }
        if ($d[1]) { $a += ' applyFont="1"' }
        if ($d[2]) { $a += ' applyFill="1"' }
        if ($d[3]) { $a += ' applyBorder="1"' }
        if ($d[5] -or $d[6] -or $d[7]) {
            $al = '<alignment'
            if ($d[5]) { $al += ' horizontal="' + $d[5] + '"' }
            if ($d[6]) { $al += ' vertical="' + $d[6] + '"' }
            if ($d[7]) { $al += ' wrapText="1"' }
            [void]$o.Append($a + ' applyAlignment="1">' + $al + '/></xf>')
        } else {
            [void]$o.Append($a + '/>')
        }
    }
    [void]$o.Append('</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
    [void]$o.Append('<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>')
    return $o.ToString()
}

# ------------------------------------------------------------------------------
# cells
# ------------------------------------------------------------------------------
function Esc([string]$s) {
    $s = $script:CtrlRe.Replace($s.Replace("`r", ''), '')
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

function EscAttr([string]$s) { return (Esc $s).Replace('"', '&quot;') }

function Col-Letter([int]$i) {
    $s = ''
    $i += 1
    while ($i -gt 0) {
        $rem = ($i - 1) % 26
        $i = [int][Math]::Floor(($i - 1) / 26)
        $s = [string][char](65 + $rem) + $s
    }
    return $s
}

function DispW([string]$s) {
    $w = 0
    foreach ($ch in $s.ToCharArray()) { if ([int]$ch -gt 0x2E80) { $w += 2 } else { $w += 1 } }
    return $w
}

$script:TextCache = New-Map
$script:DtCache = New-Map
$script:StatusCache = New-Map

function Get-TextBody([string]$text) {
    # <t>...</t> for a cell text (cached: data source names etc. repeat on every row)
    $body = $script:TextCache[$text]
    if ($null -ne $body) { return $body }
    $tx = $text
    if ($tx.Length -gt 32000) { $tx = $tx.Substring(0, 32000) }
    $sp = ''
    if (($tx -cne $tx.Trim()) -or $tx.Contains("`n")) { $sp = ' xml:space="preserve"' }
    $body = '<t' + $sp + '>' + (Esc $tx) + '</t>'
    if ($script:TextCache.Count -lt 50000) { $script:TextCache[$text] = $body }
    return $body
}

function C-Str([string]$ref, [string]$text, [string]$style) {
    if ($text -eq '') { return '<c r="' + $ref + '" s="' + $script:XF[$style] + '"/>' }
    return '<c r="' + $ref + '" s="' + $script:XF[$style] + '" t="inlineStr"><is>' + (Get-TextBody $text) + '</is></c>'
}

function C-Num([string]$ref, [string]$num, [string]$style) {
    return '<c r="' + $ref + '" s="' + $script:XF[$style] + '"><v>' + $num + '</v></c>'
}

function C-Empty([string]$ref, [string]$style) { return '<c r="' + $ref + '" s="' + $script:XF[$style] + '"/>' }

function Get-DtSerial([string]$v) {
    # returns @(serial string, hasSeconds) or $null
    $cached = $script:DtCache[$v]
    if ($null -ne $cached) { return $cached }
    $res = Get-DtSerialUncached $v
    if ($null -ne $res -and $script:DtCache.Count -lt 200000) { $script:DtCache[$v] = $res }
    return $res
}

function Get-DtSerialUncached([string]$v) {
    $m = $script:DtRe.Match($v)
    if (-not $m.Success) { return $null }
    $y = [int]$m.Groups[1].Value; $mo = [int]$m.Groups[2].Value; $d = [int]$m.Groups[3].Value
    $hh = [int]$m.Groups[4].Value; $mi = [int]$m.Groups[5].Value
    $ss = 0
    $hasSec = $m.Groups[6].Success
    if ($hasSec) { $ss = [int]$m.Groups[6].Value }
    try { $date = New-Object DateTime($y, $mo, $d) } catch { return $null }
    $days = ($date - $script:Epoch).Days
    $val = $days + ($hh * 3600 + $mi * 60 + $ss) / 86400.0
    return @((F8 $val), $hasSec)
}

function Get-StatusStyle([string]$code, [bool]$zebra) {
    if ($code -ceq 'HIGH' -or $code -ceq 'WAIT' -or $code -ceq 'EXHAUSTED') { return 'st_' + $code }
    if ($code -ceq 'NOSTATS' -or $code -ceq 'ERROR') { if ($zebra) { return 'st_GRAY_z' } else { return 'st_GRAY' } }
    if ($zebra) { return 'st_OK_z' }
    return 'st_OK'
}

function Get-StatusLabel([string]$code) {
    if ($script:STATUS.ContainsKey($code)) { return $script:STATUS[$code][0] }
    return $code
}

function Get-ValueCell([string]$ref, [string]$v, [string]$kind, [bool]$zebra, [bool]$bold) {
    $z = ''
    if ($zebra) { $z = '_z' }
    if ($v -eq '') {
        switch -CaseSensitive ($kind) {
            'int' { $base = 'int' } 'dec1' { $base = 'dec1' } 'pct' { $base = 'pct' }
            'datetime' { $base = 'dt' } 'wrap' { $base = 'wrap' } default { $base = 'text' }
        }
        return (C-Empty $ref ($base + $z))
    }
    if ($kind -ceq 'int' -or $kind -ceq 'dec1' -or $kind -ceq 'pct') {
        if (Test-Num $v) {
            if ($kind -ceq 'pct' -and $bold) { return (C-Num $ref $v ('pctb' + $z)) }
            return (C-Num $ref $v ($kind + $z))
        }
        return (C-Str $ref $v ('text' + $z))
    }
    if ($kind -ceq 'datetime') {
        $ds = Get-DtSerial $v
        if ($null -eq $ds) { return (C-Str $ref $v ('text' + $z)) }
        if ($ds[1]) { return (C-Num $ref $ds[0] ('dt' + $z)) }
        return (C-Num $ref $ds[0] ('dtm' + $z))
    }
    if ($kind -ceq 'status') { return (C-Str $ref (Get-StatusLabel $v) (Get-StatusStyle $v $zebra)) }
    if ($kind -ceq 'wrap') { return (C-Str $ref $v ('wrap' + $z)) }
    if ($bold) { return (C-Str $ref $v ('textb' + $z)) }
    return (C-Str $ref $v ('text' + $z))
}

function Get-ValueWidth([string]$v, [string]$kind) {
    if ($kind -ceq 'datetime') { if ($v.Length -gt 16) { return 19 } else { return 16 } }
    if ($kind -ceq 'status') { return (DispW (Get-StatusLabel $v)) + 2 }
    if (($kind -ceq 'int' -or $kind -ceq 'dec1' -or $kind -ceq 'pct') -and (Test-Num $v)) {
        $ip = $v.TrimStart('-').Split('.')[0]
        $extra = 0
        if ($kind -ceq 'pct') { $extra = 2 }
        return $v.Length + [int][Math]::Floor(($ip.Length - 1) / 3) + $extra
    }
    return (DispW $v.Split("`n")[0])
}

# ------------------------------------------------------------------------------
# table (1 header row + data rows)
# ------------------------------------------------------------------------------
function New-Table([string]$fname, [string[]]$header, [string[]]$front, [string[]]$back, [string[]]$skip, $fixed) {
    $idx = New-Map
    for ($i = 0; $i -lt $header.Length; $i++) { $idx[$header[$i]] = $i }
    $order = New-Object System.Collections.ArrayList
    foreach ($h in $front) { if ($idx.ContainsKey($h)) { [void]$order.Add($h) } }
    foreach ($h in $header) {
        if (-not ($order -ccontains $h) -and -not ($back -ccontains $h) -and -not ($skip -ccontains $h)) { [void]$order.Add($h) }
    }
    foreach ($h in $back) { if ($idx.ContainsKey($h)) { [void]$order.Add($h) } }
    $cols = New-Object System.Collections.ArrayList
    foreach ($h in $order) {
        $m = Get-Meta $fname $h
        $kind = $m[2]
        if ([string]::IsNullOrEmpty($kind)) { $kind = 'text' }
        $fx = $null
        if ($null -ne $fixed -and $fixed.ContainsKey($h)) { $fx = $fixed[$h] }
        [void]$cols.Add(@{ key = $h; src = $idx[$h]; label = $m[0]; unit = $m[1]; kind = $kind; fixed = $fx })
    }
    return @{ fname = $fname; cols = $cols; ncol = $cols.Count; last_col = (Col-Letter ($cols.Count - 1));
        widths = $null; header_height = 18 }
}

function Get-TableLetter($t, [string]$key) {
    for ($i = 0; $i -lt $t.cols.Count; $i++) { if ($t.cols[$i].key -ceq $key) { return (Col-Letter $i) } }
    return $null
}

function Set-TableWidths($t, $sample) {
    $widths = New-Object System.Collections.ArrayList
    $hdrLines = 1
    foreach ($c in $t.cols) {
        $lt = $c.label
        if ($c.unit) { $lt += ' (' + $c.unit + ')' }
        $lw = DispW $lt
        if ($null -ne $c.fixed) {
            $w = [int]$c.fixed
        } else {
            $vw = 0
            foreach ($r in $sample) {
                if ($c.src -lt $r.Length -and $r[$c.src] -ne '') {
                    $x = Get-ValueWidth $r[$c.src] $c.kind
                    if ($x -gt $vw) { $vw = $x }
                }
            }
            $w = [Math]::Max([Math]::Max($vw + 2, [int][Math]::Floor(($lw + 2) / 3) + 2), 8)
            $w = [Math]::Min($w, 50)
        }
        $den = [double][Math]::Max($w - 1, 1)
        $lines = [int][Math]::Ceiling($lw / $den) + [int][Math]::Ceiling($c.key.Length / $den)
        if ($lines -gt $hdrLines) { $hdrLines = $lines }
        [void]$widths.Add([int]$w)
    }
    $t.widths = $widths
    $t.header_height = [Math]::Min(15 * $hdrLines + 3, 120)
}

function Get-HeaderRowXml($t, [int]$r) {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $t.cols.Count; $i++) {
        $c = $t.cols[$i]
        $txt = $c.label
        if ($c.unit) { $txt += ' (' + $c.unit + ')' }
        $txt += "`n" + $c.key
        [void]$sb.Append((C-Str ((Col-Letter $i) + $r) $txt 'hdr'))
    }
    return '<row r="' + $r + '" ht="' + $t.header_height + '" customHeight="1">' + $sb.ToString() + '</row>'
}

function Get-DataRowXml($t, [int]$r, [string[]]$row, [bool]$zebra) {
    $sb = $t.sb
    [void]$sb.Clear()
    $ht = 0
    $rs = [string]$r
    $n = $t.ncol
    $src = $t.src; $letters = $t.letters; $fast = $t.fast; $numRe = $script:NumRe
    if ($zebra) { $es = $t.es1; $ns = $t.ns1 } else { $es = $t.es0; $ns = $t.ns0 }
    $rowLen = $row.Length
    for ($i = 0; $i -lt $n; $i++) {
        $si = $src[$i]
        if ($si -lt $rowLen) { $v = $row[$si] } else { $v = '' }
        # fast paths: empty cells and plain numbers
        if ($v.Length -eq 0) {
            [void]$sb.Append('<c r="').Append($letters[$i]).Append($rs).Append('" s="').Append($es[$i]).Append('"/>')
            continue
        }
        $f = $fast[$i]
        if ($f -eq 4) {
            $body = $script:TextCache[$v]
            if ($null -eq $body) { $body = Get-TextBody $v }
            [void]$sb.Append('<c r="').Append($letters[$i]).Append($rs).Append('" s="').Append($ns[$i]).Append('" t="inlineStr"><is>').Append($body).Append('</is></c>')
            continue
        }
        if ($f -eq 5) {
            $ds = Get-DtSerial $v
            if ($null -ne $ds) {
                if ($ds[1]) { $sname = 'dt' } else { $sname = 'dtm' }
                if ($zebra) { $sname += '_z' }
                [void]$sb.Append('<c r="').Append($letters[$i]).Append($rs).Append('" s="').Append($script:XF[$sname]).Append('"><v>').Append($ds[0]).Append('</v></c>')
                continue
            }
        }
        if ($f -eq 6) {
            $sk = $v + $script:SEP + $zebra
            $frag = $script:StatusCache[$sk]
            if ($null -eq $frag) {
                $frag = '" s="' + $script:XF[(Get-StatusStyle $v $zebra)] + '" t="inlineStr"><is>' + (Get-TextBody (Get-StatusLabel $v)) + '</is></c>'
                $script:StatusCache[$sk] = $frag
            }
            [void]$sb.Append('<c r="').Append($letters[$i]).Append($rs).Append($frag)
            continue
        }
        if ($f -ge 1 -and $f -le 3 -and $numRe.IsMatch($v)) {
            $sidx = $ns[$i]
            if ($f -eq 3 -and (To-Double $v) -ge $script:WARN_PCT) {
                if ($zebra) { $sidx = $script:XF['pctb_z'] } else { $sidx = $script:XF['pctb'] }
            }
            [void]$sb.Append('<c r="').Append($letters[$i]).Append($rs).Append('" s="').Append($sidx).Append('"><v>').Append($v).Append('</v></c>')
            continue
        }
        $c = $t.cols[$i]
        $kind = $c.kind
        $ref = $letters[$i] + $rs
        $bold = $false
        if ($kind -ceq 'pct' -and $c.key.StartsWith('usage_pct', [System.StringComparison]::Ordinal) -and (Test-Num $v) -and (To-Double $v) -ge $script:WARN_PCT) { $bold = $true }
        if ($c.key -ceq 'leak_suspect' -and $v -ceq 'yes') { $bold = $true }
        if ($kind -ceq 'wrap' -and $v -ne '') {
            $parts = $v.Split("`n")
            $lines = $parts.Length
            foreach ($line in $parts) { $lines += [int][Math]::Floor((DispW $line) / [double][Math]::Max($t.widths[$i] - 1, 1)) }
            $h = [Math]::Min(409, 13 * $lines + 4)
            if ($h -gt $ht) { $ht = $h }
        }
        [void]$sb.Append((Get-ValueCell $ref $v $kind $zebra $bold))
    }
    $hs = ''
    if ($ht -gt 0) { $hs = ' ht="' + $ht + '" customHeight="1"' }
    return '<row r="' + $rs + '"' + $hs + '>' + $sb.ToString() + '</row>'
}

function Complete-Table($t) {
    # per-column arrays used by the row writer (fast path)
    $n = $t.cols.Count
    $t.letters = New-Object 'string[]' $n
    $t.src = New-Object 'int[]' $n
    $t.fast = New-Object 'int[]' $n
    $t.es0 = New-Object 'int[]' $n
    $t.es1 = New-Object 'int[]' $n
    $t.ns0 = New-Object 'int[]' $n
    $t.ns1 = New-Object 'int[]' $n
    $t.sb = New-Object System.Text.StringBuilder 8192
    for ($i = 0; $i -lt $n; $i++) {
        $c = $t.cols[$i]
        $t.letters[$i] = Col-Letter $i
        $t.src[$i] = $c.src
        $base = 'text'
        switch -CaseSensitive ($c.kind) {
            'int' { $base = 'int' } 'dec1' { $base = 'dec1' } 'pct' { $base = 'pct' } 'datetime' { $base = 'dt' } 'wrap' { $base = 'wrap' }
        }
        $t.es0[$i] = $script:XF[$base]
        $t.es1[$i] = $script:XF[$base + '_z']
        # 1 = int/dec1, 2 = pct (not usage_pct), 3 = usage_pct (bold if >= warn),
        # 4 = plain text, 5 = datetime, 6 = status, 0 = other
        $f = 0
        if ($c.kind -ceq 'int' -or $c.kind -ceq 'dec1') { $f = 1 }
        elseif ($c.kind -ceq 'pct') {
            $f = 2
            if ($c.key.StartsWith('usage_pct', [System.StringComparison]::Ordinal)) { $f = 3 }
        }
        elseif ($c.kind -ceq 'text' -and $c.key -cne 'leak_suspect') { $f = 4 }
        elseif ($c.kind -ceq 'datetime') { $f = 5 }
        elseif ($c.kind -ceq 'status') { $f = 6 }
        $t.fast[$i] = $f
        if ($f -ge 1 -and $f -le 3) { $t.ns0[$i] = $script:XF[$c.kind]; $t.ns1[$i] = $script:XF[$c.kind + '_z'] }
        if ($f -eq 4) { $t.ns0[$i] = $script:XF['text']; $t.ns1[$i] = $script:XF['text_z'] }
    }
}

# ------------------------------------------------------------------------------
# sheet XML
# ------------------------------------------------------------------------------
$script:SHEETS = New-Object System.Collections.ArrayList
$script:ZIP = $null

function New-SheetObj([string]$name, [string]$uiKey) {
    return @{ name = $name; ui_key = $uiKey; charts = (New-Object System.Collections.ArrayList); filter_ref = $null;
        print_titles = $false; title_cols = 0; index = 0; drawing_no = 0 }
}

function Get-SheetRef([string]$name) { return "'" + $name.Replace("'", "''") + "'" }

function Write-Entry([string]$name, [string]$text) {
    $e = $script:ZIP.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $w = New-Object System.IO.StreamWriter($e.Open(), $script:Utf8)
    try { $w.Write($text) } finally { $w.Dispose() }
}

function Open-Sheet($sheet, $widths, $freeze, [bool]$grid, [bool]$selected, $defaultWidth) {
    # narrow sheets: fit to 1 page wide / wide sheets: 100% and repeat frozen columns
    $sum = 0
    if ($null -ne $widths) { foreach ($x in $widths) { $sum += $x } }
    $fit = ($null -eq $widths) -or ($sum -le 260)
    if (-not $fit -and $null -ne $freeze) { $sheet.title_cols = $freeze[0] }
    $idx = $script:SHEETS.Count + 1
    $sheet.index = $idx
    [void]$script:SHEETS.Add($sheet)
    $entry = $script:ZIP.CreateEntry('xl/worksheets/sheet' + $idx + '.xml', [System.IO.Compression.CompressionLevel]::Optimal)
    $w = New-Object System.IO.StreamWriter($entry.Open(), $script:Utf8, 65536)
    $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL)
    $w.Write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    if ($fit) { $w.Write('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>') } else { $w.Write('<sheetPr/>') }
    $sv = '<sheetView workbookViewId="0"'
    if (-not $grid) { $sv += ' showGridLines="0"' }
    if ($selected) { $sv += ' tabSelected="1"' }
    $sv += '>'
    if ($null -ne $freeze -and ($freeze[0] -or $freeze[1])) {
        $xs = $freeze[0]; $ys = $freeze[1]
        $tl = (Col-Letter $xs) + ($ys + 1)
        if ($xs -and $ys) {
            $sv += '<pane xSplit="' + $xs + '" ySplit="' + $ys + '" topLeftCell="' + $tl + '" activePane="bottomRight" state="frozen"/>' +
                '<selection pane="topRight"/><selection pane="bottomLeft"/>' +
                '<selection pane="bottomRight" activeCell="' + $tl + '" sqref="' + $tl + '"/>'
        } elseif ($ys) {
            $sv += '<pane ySplit="' + $ys + '" topLeftCell="' + $tl + '" activePane="bottomLeft" state="frozen"/>' +
                '<selection pane="bottomLeft" activeCell="' + $tl + '" sqref="' + $tl + '"/>'
        } else {
            $sv += '<pane xSplit="' + $xs + '" topLeftCell="' + $tl + '" activePane="topRight" state="frozen"/>' +
                '<selection pane="topRight" activeCell="' + $tl + '" sqref="' + $tl + '"/>'
        }
    } else {
        $sv += '<selection activeCell="A1" sqref="A1"/>'
    }
    $w.Write('<sheetViews>' + $sv + '</sheetView></sheetViews>')
    $dw = ''
    if ($null -ne $defaultWidth) { $dw = 'defaultColWidth="' + $defaultWidth + '" ' }
    $w.Write('<sheetFormatPr ' + $dw + 'defaultRowHeight="15"/>')
    if ($null -ne $widths -and $widths.Count -gt 0) {
        $cs = '<cols>'
        for ($i = 0; $i -lt $widths.Count; $i++) {
            $cs += '<col min="' + ($i + 1) + '" max="' + ($i + 1) + '" width="' + $widths[$i] + '" customWidth="1"/>'
        }
        $w.Write($cs + '</cols>')
    }
    $w.Write('<sheetData>')
    return @{ w = $w; sheet = $sheet; fit = $fit }
}

function Close-Sheet($ctx, $autofilter, $merges, [bool]$landscape) {
    $w = $ctx.w
    try {
        $w.Write('</sheetData>')
        if ($autofilter) {
            $w.Write('<autoFilter ref="' + $autofilter + '"/>')
            $ctx.sheet.filter_ref = $autofilter
        }
        if ($null -ne $merges -and $merges.Count -gt 0) {
            $ms = '<mergeCells count="' + $merges.Count + '">'
            foreach ($m in $merges) { $ms += '<mergeCell ref="' + $m + '"/>' }
            $w.Write($ms + '</mergeCells>')
        }
        $w.Write('<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
        $o = 'portrait'
        if ($landscape) { $o = 'landscape' }
        $ft = ''
        if ($ctx.fit) { $ft = ' fitToWidth="1" fitToHeight="0"' }
        $w.Write('<pageSetup paperSize="9" orientation="' + $o + '"' + $ft + '/>')
        $w.Write('<headerFooter><oddHeader>&amp;L&amp;"Meiryo UI,Regular"&amp;9&amp;A</oddHeader>' +
            '<oddFooter>&amp;C&amp;"Meiryo UI,Regular"&amp;9&amp;P / &amp;N</oddFooter></headerFooter>')
        if ($ctx.sheet.charts.Count -gt 0) { $w.Write('<drawing r:id="rId1"/>') }
        $w.Write('</worksheet>')
    } finally { $w.Dispose() }
}

function Write-SimpleSheet($sheet, $rows, $widths, $freeze, $autofilter, $merges, [bool]$grid, [bool]$selected, [bool]$landscape, $defaultWidth) {
    $ctx = Open-Sheet $sheet $widths $freeze $grid $selected $defaultWidth
    foreach ($row in $rows) { $ctx.w.Write($row) }
    Close-Sheet $ctx $autofilter $merges $landscape
}

function Add-DataSheet([string]$nameKey, [string]$fname, $rows, [string[]]$front, [string[]]$back, [string[]]$skip, $fixed, [int]$freezeCols) {
    if ($null -eq $rows -or $rows.Count -eq 0) { return $null }
    $t = New-Table $fname $rows[0] $front $back $skip $fixed
    Complete-Table $t
    $sample = New-Object System.Collections.ArrayList
    for ($i = 1; $i -lt $rows.Count -and $i -le 500; $i++) { [void]$sample.Add($rows[$i]) }
    Set-TableWidths $t $sample
    $sheet = New-SheetObj (UI $nameKey) $nameKey
    $sheet.print_titles = $true
    $last = [Math]::Max($rows.Count, 2)
    $ctx = Open-Sheet $sheet $t.widths @($freezeCols, 1) $true $false $null
    $ctx.w.Write((Get-HeaderRowXml $t 1))
    if ($rows.Count -eq 1) { $ctx.w.Write('<row r="2">' + (C-Str 'A2' (UI 'no_data') 'note') + '</row>') }
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $ctx.w.Write((Get-DataRowXml $t ($i + 1) $rows[$i] (($i - 1) % 2 -eq 1)))
    }
    Close-Sheet $ctx ('A1:' + $t.last_col + $last) $null $true
    return $sheet
}

# ------------------------------------------------------------------------------
# charts (DrawingML)
# ------------------------------------------------------------------------------
function RPr([int]$sz, [int]$bold, [string]$color) {
    return '<a:defRPr sz="' + $sz + '" b="' + $bold + '"><a:solidFill><a:srgbClr val="' + $color + '"/></a:solidFill>' +
        '<a:latin typeface="' + $script:FONT + '"/><a:ea typeface="' + $script:FONT + '"/><a:cs typeface="' + $script:FONT + '"/></a:defRPr>'
}

function TxPr([int]$sz, [int]$bold, [string]$color, $rot) {
    $bp = '<a:bodyPr/>'
    if ($null -ne $rot) { $bp = '<a:bodyPr rot="' + $rot + '" vert="horz"/>' }
    return '<c:txPr>' + $bp + '<a:lstStyle/><a:p><a:pPr>' + (RPr $sz $bold $color) + '</a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>'
}

function Rich-Title([string]$text, [int]$sz, $rot) {
    $bp = '<a:bodyPr/>'
    if ($null -ne $rot) { $bp = '<a:bodyPr rot="' + $rot + '" vert="horz"/>' }
    return '<c:title><c:tx><c:rich>' + $bp + '<a:lstStyle/><a:p><a:pPr>' + (RPr $sz 1 '000000') + '</a:pPr><a:r><a:rPr lang="ja-JP" sz="' + $sz + '" b="1">' +
        '<a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="' + $script:FONT + '"/><a:ea typeface="' + $script:FONT + '"/>' +
        '<a:cs typeface="' + $script:FONT + '"/></a:rPr><a:t>' + (Esc $text) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
}

function Ln([string]$color, [int]$w, [string]$dash) {
    return '<a:ln w="' + $w + '" cap="rnd"><a:solidFill><a:srgbClr val="' + $color + '"/></a:solidFill><a:prstDash val="' + $dash + '"/><a:round/></a:ln>'
}

$script:GRID = '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
$script:AXLINE = '<c:spPr><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr>'

function Val-Axis([int]$axId, [int]$crossId, [string]$pos, $title, [string]$fmt, [bool]$grid, [bool]$min0, [string]$crossBetween, $rot, $vmin, $vmax, $major) {
    $sc = '<c:orientation val="minMax"/>'
    if ($null -ne $vmax) { $sc += '<c:max val="' + $vmax + '"/>' }
    if ($null -ne $vmin) { $sc += '<c:min val="' + $vmin + '"/>' }
    elseif ($min0) { $sc += '<c:min val="0"/>' }
    $o = '<c:valAx><c:axId val="' + $axId + '"/><c:scaling>' + $sc + '</c:scaling><c:delete val="0"/>'
    $o += '<c:axPos val="' + $pos + '"/>'
    if ($grid) { $o += $script:GRID }
    if ($title) {
        $tr = $null
        if ($pos -eq 'l') { $tr = -5400000 }
        $o += Rich-Title $title 900 $tr
    }
    $o += '<c:numFmt formatCode="' + (EscAttr $fmt) + '" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
    $o += '<c:tickLblPos val="low"/>' + $script:AXLINE + (TxPr 800 0 '404040' $rot) + '<c:crossAx val="' + $crossId + '"/><c:crosses val="autoZero"/><c:crossBetween val="' + $crossBetween + '"/>'
    if ($null -ne $major) { $o += '<c:majorUnit val="' + $major + '"/>' }
    return $o + '</c:valAx>'
}

function Chart-Space([string]$title, [string]$plot) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/><c:chart>' + (Rich-Title $title 1100 $null) + '<c:autoTitleDeleted val="0"/>' +
        '<c:plotArea><c:layout/>' + $plot + '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>' +
        '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>' + (TxPr 900 0 '404040' $null) + '</c:legend>' +
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
        '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>' +
        (TxPr 900 0 '404040' $null) + '</c:chartSpace>'
}

function Scatter-Chart([string]$title, [string]$xref, $series, [string]$xfmt, [string]$ytitle, $xmin, $xmax) {
    # series: @(name, yref, color, widthEMU, dash, marker)
    $s = '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>'
    for ($i = 0; $i -lt $series.Count; $i++) {
        $se = $series[$i]
        $s += '<c:ser><c:idx val="' + $i + '"/><c:order val="' + $i + '"/><c:tx><c:v>' + (Esc $se[0]) + '</c:v></c:tx><c:spPr>' + (Ln $se[2] $se[3] $se[4]) + '</c:spPr>'
        if ($se[5] -eq 'none') {
            $s += '<c:marker><c:symbol val="none"/></c:marker>'
        } else {
            $s += '<c:marker><c:symbol val="' + $se[5] + '"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="' + $se[2] + '"/></a:solidFill>' + (Ln $se[2] 9525 'solid') + '</c:spPr></c:marker>'
        }
        $s += '<c:xVal><c:numRef><c:f>' + (Esc $xref) + '</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>' + (Esc $se[1]) + '</c:f></c:numRef></c:yVal><c:smooth val="0"/></c:ser>'
    }
    $s += '<c:axId val="5001"/><c:axId val="5002"/></c:scatterChart>'
    $s += Val-Axis 5001 5002 'b' $null $xfmt $false $false 'midCat' -2700000 $xmin $xmax $null
    $s += Val-Axis 5002 5001 'l' $ytitle '#,##0' $true $true 'midCat' $null $null $null $null
    return (Chart-Space $title $s)
}

function Bar-Chart([string]$title, [string]$catref, $series, [string]$catfmt, [string]$ytitle, [double]$ymax) {
    # series: @(name, valref, color)
    $s = '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
    for ($i = 0; $i -lt $series.Count; $i++) {
        $se = $series[$i]
        $s += '<c:ser><c:idx val="' + $i + '"/><c:order val="' + $i + '"/><c:tx><c:v>' + (Esc $se[0]) + '</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="' + $se[2] + '"/></a:solidFill>' +
            '<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/><c:cat><c:numRef><c:f>' + (Esc $catref) + '</c:f></c:numRef></c:cat>' +
            '<c:val><c:numRef><c:f>' + (Esc $se[1]) + '</c:f></c:numRef></c:val></c:ser>'
    }
    $s += '<c:gapWidth val="60"/><c:axId val="6001"/><c:axId val="6002"/></c:barChart>'
    $s += '<c:catAx><c:axId val="6001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>' +
        '<c:numFmt formatCode="' + (EscAttr $catfmt) + '" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>' +
        $script:AXLINE + (TxPr 800 0 '404040' -2700000) + '<c:crossAx val="6002"/><c:crosses val="autoZero"/><c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>' +
        '<c:noMultiLvlLbl val="0"/></c:catAx>'
    # fixed scale for small counts (avoid 0, 0.2, 0.4 ... ticks)
    $vmax = $null; $major = $null
    if ($ymax -lt 5) { $vmax = '5'; $major = '1' }
    $s += Val-Axis 6002 6001 'l' $ytitle '#,##0' $true $true 'between' $null $null $vmax $major
    return (Chart-Space $title $s)
}

function Get-DrawingXml($charts) {
    $o = New-Object System.Text.StringBuilder
    [void]$o.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">')
    for ($i = 0; $i -lt $charts.Count; $i++) {
        $an = $charts[$i][0]
        [void]$o.Append('<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>' + $an[0] + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + $an[1] + '</xdr:row>' +
            '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>' + $an[2] + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + $an[3] + '</xdr:row>' +
            '<xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + ($i + 2) + '" name="Chart ' + ($i + 1) + '"/>' +
            '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
            '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId' + ($i + 1) + '"/>' +
            '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')
    }
    [void]$o.Append('</xdr:wsDr>')
    return $o.ToString()
}

# ------------------------------------------------------------------------------
# sheets
# ------------------------------------------------------------------------------
function Add-SectionRow($rows, [int]$r, [string]$title, [int]$span) {
    $s = (C-Str ('A' + $r) $title 'section')
    for ($i = 1; $i -le $span; $i++) { $s += (C-Empty ((Col-Letter $i) + $r) 'section') }
    [void]$rows.Add('<row r="' + $r + '" ht="21" customHeight="1">' + $s + '</row>')
}

function Add-SummarySheet($summaryRows, [bool]$truncated) {
    $sheet = New-SheetObj (UI 'sheet_summary') 'sheet_summary'
    $hdr = @()
    $data = New-Object System.Collections.ArrayList
    if ($null -ne $summaryRows -and $summaryRows.Count -gt 0) {
        $hdr = $summaryRows[0]
        for ($i = 1; $i -lt $summaryRows.Count; $i++) { [void]$data.Add($summaryRows[$i]) }
    }
    $idx = New-Map
    for ($i = 0; $i -lt $hdr.Length; $i++) { $idx[$hdr[$i]] = $i }
    $nds = $data.Count
    $span = [Math]::Max($nds, 3)
    $widths = New-Object System.Collections.ArrayList
    [void]$widths.Add(34)
    foreach ($d in $data) {
        $wd = 26
        foreach ($key in @('datasource', 'jndi_name')) {
            if ($idx.ContainsKey($key) -and $idx[$key] -lt $d.Length) { $wd = [Math]::Max($wd, [Math]::Min(44, (DispW $d[$idx[$key]]) + 3)) }
        }
        [void]$widths.Add($wd)
    }
    while ($widths.Count -lt $span + 1) { [void]$widths.Add(26) }
    $last = Col-Letter $span
    $rows = New-Object System.Collections.ArrayList
    $merges = New-Object System.Collections.ArrayList
    [void]$rows.Add('<row r="1" ht="30" customHeight="1">' + (C-Str 'A1' (UI 'report_title') 'title') + '</row>')
    $gen = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss', $script:Inv)
    $sub = (UI 'generated') + ': ' + $gen + '    ' + (UI 'source_dir') + ': ' + $script:IN_DIR
    [void]$rows.Add('<row r="2">' + (C-Str 'A2' $sub 'subtitle') + '</row>')
    if ($truncated) { [void]$rows.Add('<row r="3">' + (C-Str 'A3' (UI 'truncated') 'note') + '</row>') }
    $r = 4
    Add-SectionRow $rows $r (UI 'sec_summary') $span
    $r++
    $cells = (C-Str ('A' + $r) (UI 'col_item') 'hdr')
    for ($j = 0; $j -lt $data.Count; $j++) { $cells += (C-Str ((Col-Letter ($j + 1)) + $r) $data[$j][$idx['datasource']] 'hdr') }
    [void]$rows.Add('<row r="' + $r + '" ht="21" customHeight="1">' + $cells + '</row>')
    $r++
    foreach ($key in $hdr) {
        if ($key -ceq 'datasource' -or $key -ceq 'verdict_code') { continue }
        $m = Get-Meta 'summary' $key
        $label = $m[0]
        if ($m[1]) { $label += ' (' + $m[1] + ')' }
        $fmt = $m[2]
        if ([string]::IsNullOrEmpty($fmt)) { $fmt = 'text' }
        $cells = (C-Str ('A' + $r) $label 'lbl')
        $ht = 0
        for ($j = 0; $j -lt $data.Count; $j++) {
            $d = $data[$j]
            $v = Get-Field $d $idx[$key]
            $ref = (Col-Letter ($j + 1)) + $r
            if ($key -ceq 'verdict') {
                $cells += (C-Str $ref $v (Get-StatusStyle (Get-Field $d $idx['verdict_code']) $false))
            } elseif ($key -ceq 'comment') {
                $cells += (C-Str $ref $v 'wrap')
                $lines = 0
                foreach ($part in $v.Split("`n")) {
                    $n = [int][Math]::Ceiling((DispW $part) / [double]($widths[$j + 1] - 2))
                    if ($n -eq 0) { $n = 1 }
                    $lines += $n
                }
                $h = [Math]::Min(409, 15 * $lines + 6)
                if ($h -gt $ht) { $ht = $h }
            } else {
                $bold = ($key.StartsWith('usage_pct', [System.StringComparison]::Ordinal) -and (Test-Num $v) -and (To-Double $v) -ge $script:WARN_PCT) -or
                    ($key -ceq 'leak_suspect' -and $v -ceq 'yes')
                $cells += (Get-ValueCell $ref $v $fmt $false $bold)
            }
        }
        $hs = ''
        if ($ht -gt 0) { $hs = ' ht="' + $ht + '" customHeight="1"' }
        [void]$rows.Add('<row r="' + $r + '"' + $hs + '>' + $cells + '</row>')
        $r++
    }
    # collection conditions / environment
    $r++
    Add-SectionRow $rows $r (UI 'sec_info') $span
    $r++
    $mergedW = 0
    for ($i = 1; $i -le $span; $i++) { $mergedW += $widths[$i] }
    foreach ($row in $script:INFO_ROWS) {
        $cells = (C-Str ('A' + $r) $row[1] 'lbl') + (C-Str ('B' + $r) $row[2] 'textw')
        for ($i = 2; $i -le $span; $i++) { $cells += (C-Empty ((Col-Letter $i) + $r) 'textw') }
        $lines = 1
        if ($row[2] -ne '') { $lines = [int][Math]::Ceiling((DispW $row[2]) / [double]($mergedW - 2)) }
        $hs = ''
        if ($lines -gt 1) { $hs = ' ht="' + (15 * $lines + 3) + '" customHeight="1"' }
        [void]$rows.Add('<row r="' + $r + '"' + $hs + '>' + $cells + '</row>')
        [void]$merges.Add('B' + $r + ':' + $last + $r)
        $r++
    }
    # status legend
    $r++
    Add-SectionRow $rows $r (UI 'sec_legend') $span
    $r++
    foreach ($code in @('EXHAUSTED', 'WAIT', 'HIGH', 'OK', 'NOSTATS', 'ERROR')) {
        if (-not $script:STATUS.ContainsKey($code)) { continue }
        $st = $script:STATUS[$code]
        $cells = (C-Str ('A' + $r) ($st[0] + ' (' + $code + ')') (Get-StatusStyle $code $false)) + (C-Str ('B' + $r) $st[1] 'textw')
        for ($i = 2; $i -le $span; $i++) { $cells += (C-Empty ((Col-Letter $i) + $r) 'textw') }
        [void]$rows.Add('<row r="' + $r + '">' + $cells + '</row>')
        [void]$merges.Add('B' + $r + ':' + $last + $r)
        $r++
    }
    Write-SimpleSheet $sheet $rows $widths $null $null $merges $false $true $false $null
    return $sheet
}

function Add-ChartSheet([string]$tsSheetName, $ts, [string]$hourlySheetName, $hr) {
    $sheet = New-SheetObj (UI 'sheet_charts') 'sheet_charts'
    $rows = New-Object System.Collections.ArrayList
    $r0 = 0
    $tsn = Get-SheetRef $tsSheetName
    $hn = $null
    if ($hourlySheetName) { $hn = Get-SheetRef $hourlySheetName }
    foreach ($o in $ts.order) {
        $key = $o[0]; $label = $o[1]
        $a = $ts.range[$key][0]; $b = $ts.range[$key][1]
        [void]$rows.Add('<row r="' + ($r0 + 1) + '" ht="21" customHeight="1">' + (C-Str ('A' + ($r0 + 1)) ($script:SQUARE + ' ' + $label) 'section') + '</row>')
        $L = $ts.letters
        $rg = @{}
        foreach ($k in @('timestamp', 'InUseCount', 'ActiveCount', 'IdleCount', 'max_pool_size')) {
            $rg[$k] = $tsn + '!$' + $L[$k] + '$' + $a + ':$' + $L[$k] + '$' + $b
        }
        $series = @(
            @((UI 'ser_inuse'), $rg['InUseCount'], '000000', 28575, 'solid', 'none'),
            @((UI 'ser_active'), $rg['ActiveCount'], '7F7F7F', 19050, 'solid', 'none'),
            @((UI 'ser_idle'), $rg['IdleCount'], 'BFBFBF', 12700, 'solid', 'none'),
            @((UI 'ser_maxpool'), $rg['max_pool_size'], '000000', 15875, 'dash', 'none')
        )
        $xmin = $null; $xmax = $null
        if ($ts.x.ContainsKey($key)) { $xmin = $ts.x[$key][0]; $xmax = $ts.x[$key][1] }
        [void]$sheet.charts.Add(@(@(0, ($r0 + 1), 16, ($r0 + 20)),
                (Scatter-Chart ((UI 'chart_ts') + ' - ' + $label) $rg['timestamp'] $series 'mm/dd hh:mm' (UI 'axis_conn') $xmin $xmax)))
        if ($hn -and $hr.range.ContainsKey($key)) {
            $ha = $hr.range[$key][0]; $hb = $hr.range[$key][1]
            $HL = $hr.letters
            $hg = @{}
            foreach ($k in @('bucket_start', 'inuse_max', 'inuse_avg', 'inuse_min', 'max_pool_size', 'd_WaitCount', 'd_BlockingFailureCount')) {
                $hg[$k] = $hn + '!$' + $HL[$k] + '$' + $ha + ':$' + $HL[$k] + '$' + $hb
            }
            $hs = @(
                @((UI 'ser_inuse_max'), $hg['inuse_max'], '000000', 22225, 'solid', 'circle'),
                @((UI 'ser_inuse_avg'), $hg['inuse_avg'], '7F7F7F', 19050, 'solid', 'square'),
                @((UI 'ser_inuse_min'), $hg['inuse_min'], 'BFBFBF', 12700, 'sysDash', 'none'),
                @((UI 'ser_maxpool'), $hg['max_pool_size'], '000000', 15875, 'dash', 'none')
            )
            $hxmin = $null; $hxmax = $null
            if ($hr.x.ContainsKey($key)) { $hxmin = $hr.x[$key][0]; $hxmax = $hr.x[$key][1] }
            [void]$sheet.charts.Add(@(@(0, ($r0 + 21), 8, ($r0 + 38)),
                    (Scatter-Chart ((UI 'chart_hourly') + ' - ' + $label) $hg['bucket_start'] $hs 'mm/dd hh:mm' (UI 'axis_conn') $hxmin $hxmax)))
            $bs = @(
                @((UI 'ser_wait'), $hg['d_WaitCount'], '808080'),
                @((UI 'ser_bf'), $hg['d_BlockingFailureCount'], '000000')
            )
            $ym = 0.0
            if ($hr.y.ContainsKey($key)) { $ym = $hr.y[$key] }
            [void]$sheet.charts.Add(@(@(8, ($r0 + 21), 16, ($r0 + 38)),
                    (Bar-Chart ((UI 'chart_hourly_err') + ' - ' + $label) $hg['bucket_start'] $bs 'mm/dd hh:mm' (UI 'axis_count') $ym)))
        }
        $r0 += 40
    }
    Write-SimpleSheet $sheet $rows $null $null $null $null $false $false $true 9
    return $sheet
}

function Get-DsKey([string[]]$row, [int]$iDs, [int]$iTy) {
    return (Get-Field $row $iDs) + $script:SEP + (Get-Field $row $iTy)
}

function Prepare-Timeseries {
    # pass 1: data sources, row counts, sample rows (rows are written per data source later)
    $path = Get-CsvPath 'pool_timeseries.csv'
    $p = Open-CsvParser $path
    $order = New-Object System.Collections.ArrayList
    $counts = New-Map
    $stamps = New-Map
    $sample = New-Object System.Collections.ArrayList
    try {
        $header = [string[]]$p.ReadFields()
        $idx = New-Map
        for ($i = 0; $i -lt $header.Length; $i++) { $idx[$header[$i]] = $i }
        $iDs = 0; if ($idx.ContainsKey('datasource')) { $iDs = $idx['datasource'] }
        $iTy = 0; if ($idx.ContainsKey('ds_type')) { $iTy = $idx['ds_type'] }
        $iTs = 0; if ($idx.ContainsKey('timestamp')) { $iTs = $idx['timestamp'] }
        while (-not $p.EndOfData) {
            $row = [string[]]$p.ReadFields()
            if ($null -eq $row -or $row.Length -eq 0) { continue }
            $k = Get-DsKey $row $iDs $iTy
            if (-not $counts.ContainsKey($k)) {
                $counts[$k] = 0
                [void]$order.Add($k)
                $stamps[$k] = New-Object System.Collections.ArrayList
            }
            $counts[$k] = $counts[$k] + 1
            $sv = $null
            if ($iTs -lt $row.Length) { $ds = Get-DtSerial $row[$iTs]; if ($null -ne $ds) { $sv = $ds[0] } }
            [void]$stamps[$k].Add($sv)
            if ($sample.Count -lt 500) { [void]$sample.Add($row) }
        }
    } finally { $p.Close() }
    $t = New-Table 'timeseries' $header @('timestamp', 'datasource') @('epoch', 'elapsed_sec', 'sample_no', 'host') @() $null
    Complete-Table $t
    Set-TableWidths $t $sample
    $nds = [Math]::Max($order.Count, 1)
    $perDs = [int][Math]::Floor($script:MAX_ROWS / $nds)
    $truncated = $false
    foreach ($k in $order) { if ($counts[$k] -gt $perDs) { $truncated = $true } }
    $range = New-Map
    $xr = New-Map
    $ord = New-Object System.Collections.ArrayList
    $r = 2
    foreach ($k in $order) {
        $n = [Math]::Min($counts[$k], $perDs)
        if ($n -gt 0) {
            $range[$k] = @($r, ($r + $n - 1))
            [void]$ord.Add(@($k, ($k.Replace($script:SEP, ' (') + ')')))
            $lo = $null; $hi = $null
            for ($i = 0; $i -lt $n; $i++) {
                $sv = $stamps[$k][$i]
                if ($null -eq $sv) { continue }
                $x = To-Double $sv
                if ($null -eq $lo -or $x -lt $lo) { $lo = $x }
                if ($null -eq $hi -or $x -gt $hi) { $hi = $x }
            }
            if ($null -ne $lo) {
                if ($hi - $lo -lt 1e-9) { $lo = $lo - 5 / 1440.0; $hi = $hi + 5 / 1440.0 }
                $xr[$k] = @((F8 $lo), (F8 $hi))
            }
            $r += $n
        }
    }
    $letters = New-Map
    foreach ($key in @('timestamp', 'InUseCount', 'ActiveCount', 'IdleCount', 'max_pool_size')) { $letters[$key] = Get-TableLetter $t $key }
    return @{ table = $t; order = $ord; range = $range; x = $xr; letters = $letters; truncated = $truncated;
        dsorder = $order; perDs = $perDs; lastRow = [Math]::Max($r - 1, 2); iDs = $iDs; iTy = $iTy }
}

function Add-TimeseriesSheet($ts) {
    $t = $ts.table
    $sheet = New-SheetObj (UI 'sheet_timeseries') 'sheet_timeseries'
    $sheet.print_titles = $true
    $ctx = Open-Sheet $sheet $t.widths @(2, 1) $true $false $null
    $ctx.w.Write((Get-HeaderRowXml $t 1))
    $path = Get-CsvPath 'pool_timeseries.csv'
    $rn = 2
    $iDs = $ts.iDs; $iTy = $ts.iTy
    foreach ($k in $ts.dsorder) {
        $written = 0
        $p = Open-CsvParser $path
        try {
            [void]$p.ReadFields()
            while (-not $p.EndOfData) {
                $row = [string[]]$p.ReadFields()
                if ($null -eq $row -or $row.Length -eq 0) { continue }
                $a = ''; $b = ''
                if ($iDs -lt $row.Length) { $a = $row[$iDs] }
                if ($iTy -lt $row.Length) { $b = $row[$iTy] }
                if (($a + $script:SEP + $b) -cne $k) { continue }
                if ($written -ge $ts.perDs) { break }
                $ctx.w.Write((Get-DataRowXml $t $rn $row ($rn % 2 -eq 1)))
                $rn++
                $written++
            }
        } finally { $p.Close() }
    }
    Close-Sheet $ctx ('A1:' + $t.last_col + $ts.lastRow) $null $true
    return $sheet
}

function Prepare-Hourly($rows) {
    $res = @{ range = (New-Map); x = (New-Map); y = (New-Map); letters = (New-Map) }
    if ($null -eq $rows -or $rows.Count -eq 0) { return $res }
    $hdr = $rows[0]
    $t = New-Table 'hourly' $hdr @('bucket_start', 'bucket_end', 'datasource') @() @() $null
    $idx = New-Map
    for ($i = 0; $i -lt $hdr.Length; $i++) { $idx[$hdr[$i]] = $i }
    $half = 60 / 1440.0 / 2
    if ($script:INFO.ContainsKey('bucket_min')) {
        $bm = 0.0
        if ([double]::TryParse($script:INFO['bucket_min'], [System.Globalization.NumberStyles]::Float, $script:Inv, [ref]$bm)) { $half = $bm / 1440.0 / 2 }
    }
    $iDs = 0; if ($idx.ContainsKey('datasource')) { $iDs = $idx['datasource'] }
    $iTy = 0; if ($idx.ContainsKey('ds_type')) { $iTy = $idx['ds_type'] }
    $xs = New-Map
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $row = $rows[$i]
        $k = Get-DsKey $row $iDs $iTy
        if (-not $res.range.ContainsKey($k)) {
            $res.range[$k] = @(($i + 1), ($i + 1))
            $xs[$k] = New-Object System.Collections.ArrayList
            $res.y[$k] = 0.0
        }
        $res.range[$k][1] = $i + 1
        if ($idx.ContainsKey('bucket_start')) {
            $ds = Get-DtSerial (Get-Field $row $idx['bucket_start'])
            if ($null -ne $ds) { [void]$xs[$k].Add((To-Double $ds[0])) }
        }
        foreach ($key in @('d_WaitCount', 'd_BlockingFailureCount')) {
            if ($idx.ContainsKey($key) -and $idx[$key] -lt $row.Length -and (Test-Num $row[$idx[$key]])) {
                $res.y[$k] = [Math]::Max([double]$res.y[$k], (To-Double $row[$idx[$key]]))
            }
        }
    }
    foreach ($k in @($res.range.Keys)) {
        if ($xs[$k].Count -gt 0) {
            $lo = $xs[$k][0]; $hi = $xs[$k][0]
            foreach ($x in $xs[$k]) { if ($x -lt $lo) { $lo = $x }; if ($x -gt $hi) { $hi = $x } }
            $res.x[$k] = @((F8 ($lo - $half)), (F8 ($hi + $half)))
        }
    }
    foreach ($key in @('bucket_start', 'inuse_max', 'inuse_avg', 'inuse_min', 'max_pool_size', 'd_WaitCount', 'd_BlockingFailureCount')) {
        $res.letters[$key] = Get-TableLetter $t $key
    }
    return $res
}

function Add-ColumnsSheet {
    $sheet = New-SheetObj (UI 'sheet_columns') 'sheet_columns'
    $fmap = New-Map
    $fmap['summary'] = 'sheet_summary'; $fmap['info'] = 'sheet_summary'; $fmap['hourly'] = 'sheet_hourly'
    $fmap['events'] = 'sheet_events'; $fmap['timeseries'] = 'sheet_timeseries'; $fmap['queued'] = 'sheet_queued'
    $fmap['qstack'] = 'sheet_qstack'; $fmap['threaddump'] = 'sheet_threaddump'; $fmap['config'] = 'sheet_config'
    $fmap['status'] = 'sec_legend'
    $widths = @(16, 30, 36, 9, 90)
    $rows = New-Object System.Collections.ArrayList
    $hc = ''
    $hk = @('col_sheet', 'col_item', 'col_key', 'col_unit', 'col_desc')
    for ($i = 0; $i -lt $hk.Length; $i++) { $hc += (C-Str ((Col-Letter $i) + '1') (UI $hk[$i]) 'hdr') }
    [void]$rows.Add('<row r="1" ht="21" customHeight="1">' + $hc + '</row>')
    $r = 2
    for ($i = 1; $i -lt $script:COLS_ROWS.Count; $i++) {
        $row = $script:COLS_ROWS[$i]
        if ($row.Length -lt 6 -or $row[0] -ceq 'ui' -or -not $fmap.ContainsKey($row[0])) { continue }
        $z = ($r % 2 -eq 1)
        $zs = ''
        if ($z) { $zs = '_z' }
        $labStyle = 'text' + $zs
        if ($row[0] -ceq 'status') { $labStyle = Get-StatusStyle $row[1] $z }
        $cells = (C-Str ('A' + $r) (UI $fmap[$row[0]]) ('text' + $zs)) + (C-Str ('B' + $r) $row[2] $labStyle) +
            (C-Str ('C' + $r) $row[1] ('text' + $zs)) + (C-Str ('D' + $r) $row[3] ('text' + $zs)) +
            (C-Str ('E' + $r) $row[5] ('text' + $zs))
        [void]$rows.Add('<row r="' + $r + '">' + $cells + '</row>')
        $r++
    }
    $sheet.print_titles = $true
    Write-SimpleSheet $sheet $rows $widths @(0, 1) ('A1:E' + [Math]::Max($r - 1, 2)) $null $true $false $true $null
    return $sheet
}

# ------------------------------------------------------------------------------
# package (zip)
# ------------------------------------------------------------------------------
function Invoke-Main {
    $tmp = $script:OUT_PATH + '.tmp'
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
    $script:ZIP = [System.IO.Compression.ZipFile]::Open($tmp, [System.IO.Compression.ZipArchiveMode]::Create)
    $ok = $false
    try {
        $summary = Read-CsvFile 'pool_summary.csv'
        $hourly = Read-CsvFile 'pool_hourly.csv'
        $events = Read-CsvFile 'pool_events.csv'
        $queued = Read-CsvFile 'pool_queued_threads.csv'
        $qstack = Read-CsvFile 'pool_queued_stacks.csv'
        $tdump = Read-CsvFile 'pool_threaddump_threads.csv'
        $config = Read-CsvFile 'pool_config.csv'

        $ts = Prepare-Timeseries
        $hr = Prepare-Hourly $hourly

        [void](Add-SummarySheet $summary $ts.truncated)
        if ($script:CHARTS -and $ts.order.Count -gt 0) {
            $hname = $null
            if ($null -ne $hourly -and $hourly.Count -gt 1) { $hname = UI 'sheet_hourly' }
            [void](Add-ChartSheet (UI 'sheet_timeseries') $ts $hname $hr)
        }
        [void](Add-DataSheet 'sheet_hourly' 'hourly' $hourly @('bucket_start', 'bucket_end', 'datasource') @() @() $null 3)
        [void](Add-DataSheet 'sheet_events' 'events' $events @('event_no', 'datasource') @() @('severity_label') $null 2)
        [void](Add-TimeseriesSheet $ts)
        if ($null -ne $queued -and $queued.Count -gt 1) {
            [void](Add-DataSheet 'sheet_queued' 'queued' $queued @('timestamp', 'datasource') @('epoch', 'sample_no', 'host') @() $null 2)
        }
        if ($null -ne $qstack -and $qstack.Count -gt 1) {
            [void](Add-DataSheet 'sheet_qstack' 'qstack' $qstack @() @() @() @{ stack = 110 } 2)
        }
        if ($null -ne $tdump -and $tdump.Count -gt 1) {
            [void](Add-DataSheet 'sheet_threaddump' 'threaddump' $tdump @() @() @() $null 1)
        }
        [void](Add-DataSheet 'sheet_config' 'config' $config @() @() @() @{ value = 60 } 4)
        [void](Add-ColumnsSheet)

        # drawings / charts
        $chartNo = 0
        $d = 0
        foreach ($s in $script:SHEETS) {
            if ($s.charts.Count -eq 0) { continue }
            $d++
            Write-Entry ('xl/drawings/drawing' + $d + '.xml') (Get-DrawingXml $s.charts)
            $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            for ($i = 0; $i -lt $s.charts.Count; $i++) {
                $chartNo++
                Write-Entry ('xl/charts/chart' + $chartNo + '.xml') $s.charts[$i][1]
                $rels += '<Relationship Id="rId' + ($i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" ' +
                    'Target="../charts/chart' + $chartNo + '.xml"/>'
            }
            $rels += '</Relationships>'
            Write-Entry ('xl/drawings/_rels/drawing' + $d + '.xml.rels') $rels
            Write-Entry ('xl/worksheets/_rels/sheet' + $s.index + '.xml.rels') ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ' +
                'Target="../drawings/drawing' + $d + '.xml"/></Relationships>')
            $s.drawing_no = $d
        }

        # workbook
        $wb = New-Object System.Text.StringBuilder
        [void]$wb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15600" activeTab="0"/></bookViews><sheets>')
        foreach ($s in $script:SHEETS) {
            [void]$wb.Append('<sheet name="' + (EscAttr $s.name) + '" sheetId="' + $s.index + '" r:id="rId' + $s.index + '"/>')
        }
        [void]$wb.Append('</sheets>')
        $names = ''
        foreach ($s in $script:SHEETS) {
            if ($s.filter_ref) {
                $parts = $s.filter_ref.Split(':')
                $m1 = $script:CellRe.Match($parts[0]); $m2 = $script:CellRe.Match($parts[1])
                $names += '<definedName name="_xlnm._FilterDatabase" localSheetId="' + ($s.index - 1) + '" hidden="1">' + (Esc (Get-SheetRef $s.name)) +
                    '!$' + $m1.Groups[1].Value + '$' + $m1.Groups[2].Value + ':$' + $m2.Groups[1].Value + '$' + $m2.Groups[2].Value + '</definedName>'
            }
            if ($s.print_titles) {
                $ref = Esc (Get-SheetRef $s.name)
                $cols = ''
                if ($s.title_cols) { $cols = $ref + '!$A:$' + (Col-Letter ($s.title_cols - 1)) + ',' }
                $names += '<definedName name="_xlnm.Print_Titles" localSheetId="' + ($s.index - 1) + '">' + $cols + $ref + '!$1:$1</definedName>'
            }
        }
        if ($names) { [void]$wb.Append('<definedNames>' + $names + '</definedNames>') }
        [void]$wb.Append('<calcPr calcId="191029"/></workbook>')
        Write-Entry 'xl/workbook.xml' $wb.ToString()

        $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        foreach ($s in $script:SHEETS) {
            $rels += '<Relationship Id="rId' + $s.index + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ' +
                'Target="worksheets/sheet' + $s.index + '.xml"/>'
        }
        $rels += '<Relationship Id="rId' + ($script:SHEETS.Count + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" ' +
            'Target="styles.xml"/></Relationships>'
        Write-Entry 'xl/_rels/workbook.xml.rels' $rels
        Write-Entry 'xl/styles.xml' (Get-StylesXml)

        $now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", $script:Inv)
        Write-Entry 'docProps/core.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
            '<dc:title>' + (Esc (UI 'report_title')) + '</dc:title><dc:creator>pool_csv_to_xlsx</dc:creator>' +
            '<dcterms:created xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:created>' +
            '<dcterms:modified xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:modified></cp:coreProperties>')
        Write-Entry 'docProps/app.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
            '<Application>Microsoft Excel</Application></Properties>')
        Write-Entry '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
            '</Relationships>')
        $ct = New-Object System.Text.StringBuilder
        [void]$ct.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + $script:NL +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
        foreach ($s in $script:SHEETS) {
            [void]$ct.Append('<Override PartName="/xl/worksheets/sheet' + $s.index + '.xml" ' +
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            if ($s.charts.Count -gt 0) {
                [void]$ct.Append('<Override PartName="/xl/drawings/drawing' + $s.drawing_no + '.xml" ' +
                    'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
            }
        }
        for ($i = 1; $i -le $chartNo; $i++) {
            [void]$ct.Append('<Override PartName="/xl/charts/chart' + $i + '.xml" ' +
                'ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
        }
        [void]$ct.Append('</Types>')
        Write-Entry '[Content_Types].xml' $ct.ToString()
        $ok = $true
    } finally {
        $script:ZIP.Dispose()
        if (-not $ok) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
    if (Test-Path -LiteralPath $script:OUT_PATH) { Remove-Item -LiteralPath $script:OUT_PATH -Force }
    Move-Item -LiteralPath $tmp -Destination $script:OUT_PATH
    $names = @()
    foreach ($s in $script:SHEETS) { $names += $s.name }
    Write-Host ((UI 'msg_created') + ': ' + $script:OUT_PATH)
    Write-Host ('  ' + (UI 'msg_sheets') + ': ' + ($names -join ' / '))
    if ($ts.truncated) { Write-Host ('  ' + (UI 'truncated') + ' (-MaxRows)') }
}

Invoke-Main
