# Apply quarter-header links through Excel rather than openpyxl.
#
# openpyxl cannot round-trip every workbook: writing BE.xlsx or TSLA.xlsx
# through it produces a file Excel refuses to open, apparently over its handling
# of cell comments. Those models carry real work, so their links are applied
# here instead, by Excel itself.
param(
    [Parameter(Mandatory = $true)][string]$Ticker,
    [Parameter(Mandatory = $true)][string]$JsonPath
)

$ErrorActionPreference = "Stop"
$path = Join-Path (Get-Location) "$Ticker.xlsx"
$links = Get-Content $JsonPath -Raw | ConvertFrom-Json

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($path)
    $ws = $wb.Worksheets.Item("Model")

    $applied = 0; $periodic = 0
    for ($c = 1; $c -le 400; $c++) {
        $v = [string]$ws.Cells.Item(2, $c).Value2
        if (-not $v) { continue }
        $entry = $links.$v
        if (-not $entry) { continue }

        $cell = $ws.Cells.Item(2, $c)
        $ws.Hyperlinks.Add($cell, $entry.url) | Out-Null
        $cell.Value2 = $v                       # Add() can blank the label
        if ($entry.periodic) {
            $cell.Font.Color = 10498160         # purple, BGR for FF7030A0
            $cell.Font.Italic = $true
            $periodic++
        } else {
            $cell.Font.Color = 12419407         # blue, BGR for FF0563C1
            $cell.Font.Italic = $false
        }
        $cell.Font.Underline = 2                # xlUnderlineStyleSingle
        $applied++
    }

    $wb.Save()
    $wb.Close($false)
    Write-Output "$Ticker : applied $applied link(s) ($periodic periodic)"
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}
