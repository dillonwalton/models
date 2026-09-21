# Prepend historical quarter columns to a model via Excel, so that formulas,
# hyperlinks, merged ranges and conditional formats are all re-pointed by Excel
# itself. openpyxl would move the cells but leave every formula reference
# behind, silently corrupting models that carry real data.
param(
    [Parameter(Mandatory = $true)][string]$Ticker,
    [Parameter(Mandatory = $true)][int]$FirstQuarter,   # 1-4
    [Parameter(Mandatory = $true)][int]$FirstYear       # e.g. 2000
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

    # Locate the first quarter header on row 2 and learn its label format.
    $firstCol = 0; $label = $null
    for ($c = 1; $c -le 200; $c++) {
        $v = [string]$ws.Cells.Item(2, $c).Value2
        if ($v -match '^Q[1-4](\d{2}|\d{4})$') { $firstCol = $c; $label = $v; break }
    }
    if ($firstCol -eq 0) { throw "no quarter headers on row 2" }
    $fourDigit = ($label.Length -eq 6)

    # Existing first quarter, so we know how many to prepend.
    $curQ = [int]$label.Substring(1, 1)
    $curY = [int]$label.Substring(2)
    if (-not $fourDigit) { if ($curY -ge 70) { $curY += 1900 } else { $curY += 2000 } }

    $labels = @()
    $y = $FirstYear; $q = $FirstQuarter
    while (($y -lt $curY) -or ($y -eq $curY -and $q -lt $curQ)) {
        $labels += if ($fourDigit) { "Q$q$y" } else { "Q$q" + ("{0:D2}" -f ($y % 100)) }
        $q++; if ($q -gt 4) { $q = 1; $y++ }
    }
    if ($labels.Count -eq 0) { Write-Output "$Ticker already starts at its first public quarter"; $wb.Close($false); return }

    # Insert the block in one operation, then label it. Excel rewrites every
    # reference that pointed at the shifted cells.
    $n = $labels.Count
    $ws.Range($ws.Columns.Item($firstCol), $ws.Columns.Item($firstCol + $n - 1)).Insert() | Out-Null

    # Match the formatting of the (now shifted) original first header.
    $src = $ws.Cells.Item(2, $firstCol + $n)
    for ($i = 0; $i -lt $n; $i++) {
        $cell = $ws.Cells.Item(2, $firstCol + $i)
        $src.Copy() | Out-Null
        $cell.PasteSpecial(-4122) | Out-Null          # xlPasteFormats
        $cell.Value2 = $labels[$i]
        $cell.Font.Underline = -4142                   # xlUnderlineStyleNone
        $cell.Font.ColorIndex = -4105                  # xlColorIndexAutomatic
        $cell.Hyperlinks.Delete()
    }
    $wb.Save()
    $wb.Close($false)
    Write-Output "$Ticker : inserted $n column(s), now starting $($labels[0])"
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}
