$excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($excelPath)
$ws = $wb.Sheets("Facturas")

# Get total rows
$lastRow = $ws.UsedRange.Rows.Count
Write-Host "Total rows: $lastRow"

# Quick scan: count REVISAR by provider
$revisar_providers = @{}
for ($row = 2; $row -le $lastRow; $row++) {
    if ($row % 100 -eq 0) { Write-Host "Scanning row $row..." }
    $cat = $ws.Cells($row, 17).Value
    if ($cat -eq "REVISAR") {
        $prov = $ws.Cells($row, 4).Value -or "UNKNOWN"
        $total = [double]($ws.Cells($row, 16).Value -or 0)
        if (-not $revisar_providers[$prov]) {
            $revisar_providers[$prov] = @{ count = 0; sum = 0 }
        }
        $revisar_providers[$prov].count += 1
        $revisar_providers[$prov].sum += $total
    }
}

$totalRevisar = ($revisar_providers.Values | Measure-Object -Sum count | Select-Object -ExpandProperty Sum)
Write-Host "`nTotal REVISAR: $totalRevisar`n"

# Show top providers
$revisar_providers.GetEnumerator() | Sort-Object { $_.Value.count } -Descending | Select-Object -First 25 | ForEach-Object {
    Write-Host "$($_.Key): $($_.Value.count) items, `$$($_.Value.sum) CLP"
}

$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
