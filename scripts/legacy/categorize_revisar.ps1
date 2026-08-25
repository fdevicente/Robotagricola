# Load Excel file
$excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($excelPath)
$ws = $wb.Sheets("Facturas")

# Extract all REVISAR facturas with their details
$revisar = @()
for ($row = 2; $row -le $ws.UsedRange.Rows.Count; $row++) {
    $cat = $ws.Cells($row, 17).Value
    if ($cat -eq "REVISAR") {
        $revisar += @{
            fila = $row
            fecha = $ws.Cells($row, 1).Value
            proveedor = $ws.Cells($row, 4).Value
            nro = $ws.Cells($row, 7).Value
            detalle = $ws.Cells($row, 8).Value
            total = [double]($ws.Cells($row, 16).Value -or 0)
        }
    }
}

Write-Host "Total REVISAR: $($revisar.Count)"

# Count by provider
$byProvider = @{}
foreach ($f in $revisar) {
    $p = $f.proveedor -or "DESCONOCIDO"
    if (-not $byProvider[$p]) { $byProvider[$p] = @() }
    $byProvider[$p] += $f
}

# Show provider distribution
Write-Host "`nDistribution by Provider:"
$byProvider.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | ForEach-Object {
    $total = ($_.Value | Measure-Object -Sum total | Select-Object -ExpandProperty Sum)
    Write-Host "$($_.Key): $($_.Value.Count) invoices, total `$$total CLP"
} | Select-Object -First 20

# Function to categorize based on provider
function Get-Category {
    param($proveedor, $detalle)

    $p = ($proveedor -or "").ToUpper()
    $d = ($detalle -or "").ToUpper()

    if ($p -match "ECOSMART") { return "SERVICIOS DE EXPORTACION" }
    if ($p -match "S.INVEST|S-INVEST") { return "COSTO ENERGETICO" }
    if ($p -match "SCOTIABANK" -or $d -match "CAMBIO|DIVISA|COMPRA DOLAR") { return "CAMBIO DIVISA" }
    if ($p -match "LUIS QUIROZ|LUIS ALBERTO" -or $d -match "FLETE|TRANSPORTE") { return "TRANSPORTE" }
    if ($p -match "AGUAS|AGUA" -or $d -match "AGUA|RIEGO") { return "AGUA" }
    if ($p -match "LUZ|ENERGETIC|ELECTR|SUMINISTR") { return "ENERGIA" }
    if ($p -match "FERTILI|ABONO|INSECTICIDA|FERTILIZANTE") { return "INSUMOS AGRICOLAS" }
    if ($p -match "MANTENIMIENTO|REPARACION|SERVICIO TECNICO") { return "MANTENIMIENTO MAQUINARIA" }
    if ($p -match "SEGUROS|SEGUROS DE" -or $d -match "SEGURO") { return "SEGUROS" }
    if ($d -match "ARRENDAMIEN|ARRIENDO") { return "ARRENDAMIENTO" }

    return "POR REVISAR"
}

# Show sample categorization for largest unknown providers
Write-Host "`n=== Sample Categorization by Provider ==="
$byProvider.GetEnumerator() | Sort-Object { ($_.Value | Measure-Object -Sum total | Select-Object -ExpandProperty Sum) } -Descending | Select-Object -First 15 | ForEach-Object {
    $prov = $_.Key
    $cat = Get-Category $prov ""
    $total = ($_.Value | Measure-Object -Sum total | Select-Object -ExpandProperty Sum)
    Write-Host "$prov ($($_.Value.Count) items) -> $cat [$$total]"
}

# Clean up
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
