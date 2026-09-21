# Prepend historical year columns to a model's annual block, through Excel so
# that formulas, merged ranges and conditional formats are re-pointed. The
# annual block sits to the right of the quarter block, so this must be run
# after extend_quarters.ps1 (or the column it finds will move underneath it).
param(
    [Parameter(Mandatory = $true)][string]$Ticker,
    [Parameter(Mandatory = $true)][int]$FirstYear
)

$ErrorActionPreference = "Stop"
$path = Join-Path (Join-Path (Get-Location) "companies") "$Ticker.xlsx"
if (-not (Test-Path $path)) { throw "no model at $path" }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($path)
    $ws = $wb.Worksheets.Item("Model")

    # First bare 4-digit header on row 2 is the start of the annual block.
    $firstCol = 0; $curY = 0
    for ($c = 1; $c -le 400; $c++) {
        $v = [string]$ws.Cells.Item(2, $c).Value2
        if ($v -match '^\d{4}$') { $firstCol = $c; $curY = [int]$v; break }
    }
    if ($firstCol -eq 0) { Write-Output "$Ticker has no annual block"; $wb.Close($false); return }
    if ($curY -le $FirstYear) { Write-Output "$Ticker annual block already starts $curY"; $wb.Close($false); return }

    $n = $curY - $FirstYear
    $ws.Range($ws.Columns.Item($firstCol), $ws.Columns.Item($firstCol + $n - 1)).Insert() | Out-Null

    $src = $ws.Cells.Item(2, $firstCol + $n)
    for ($i = 0; $i -lt $n; $i++) {
        $cell = $ws.Cells.Item(2, $firstCol + $i)
        $src.Copy() | Out-Null
        $cell.PasteSpecial(-4122) | Out-Null          # xlPasteFormats
        $cell.Value2 = $FirstYear + $i
        $cell.Hyperlinks.Delete()
    }

    $wb.Save()
    $wb.Close($false)
    Write-Output "$Ticker : inserted $n year column(s), annual block now starts $FirstYear"
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}
