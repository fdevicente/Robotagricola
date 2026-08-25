Function GetCategory(proveedor, detalle)
    Dim p, d
    p = UCase(proveedor)
    d = UCase(detalle)

    ' Known cases from user input
    If InStr(p, "ECOSMART") > 0 Then GetCategory = "SERVICIOS DE EXPORTACION": Exit Function
    If InStr(p, "S-INVEST") > 0 Or InStr(p, "S.INVEST") > 0 Then GetCategory = "COSTO ENERGETICO": Exit Function
    If InStr(p, "SCOTIABANK") > 0 Then GetCategory = "CAMBIO DIVISA": Exit Function
    If InStr(p, "LUIS QUIROZ") > 0 Or InStr(p, "LUIS ALBERTO") > 0 Then GetCategory = "TRANSPORTE": Exit Function

    ' Energy/Utilities
    If InStr(p, "CGE") > 0 Then GetCategory = "ENERGIA": Exit Function
    If InStr(p, "ELECTRICIDAD") > 0 Then GetCategory = "ENERGIA": Exit Function
    If InStr(p, "LUZ") > 0 Or InStr(p, "ENERGETIC") > 0 Or InStr(p, "ELECTR") > 0 Then GetCategory = "ENERGIA": Exit Function
    If InStr(d, "ENERGIA") > 0 Or InStr(d, "KWH") > 0 Then GetCategory = "ENERGIA": Exit Function

    ' Water/Irrigation
    If InStr(p, "AGUA") > 0 Or InStr(p, "RIEGO") > 0 Then GetCategory = "AGUA": Exit Function

    ' Transport/Freight
    If InStr(p, "TRANSPORTE") > 0 Or InStr(p, "FLETE") > 0 Or InStr(p, "DHL") > 0 Then GetCategory = "TRANSPORTE": Exit Function
    If InStr(p, "LATAM") > 0 Or InStr(p, "AIRLINES") > 0 Or InStr(p, "AROMAR") > 0 Then GetCategory = "TRANSPORTE": Exit Function
    If InStr(d, "FLETE") > 0 Or InStr(d, "TRANSPORTE") > 0 Then GetCategory = "TRANSPORTE": Exit Function

    ' Agricultural inputs
    If InStr(p, "AGROKIMUN") > 0 Then GetCategory = "INSUMOS AGRICOLAS": Exit Function
    If InStr(p, "FERTILI") > 0 Or InStr(p, "ABONO") > 0 Or InStr(p, "INSECTICIDA") > 0 Then GetCategory = "INSUMOS AGRICOLAS": Exit Function
    If InStr(d, "FERTILI") > 0 Or InStr(d, "PESTICIDA") > 0 Or InStr(d, "AGROQUIMICA") > 0 Then GetCategory = "INSUMOS AGRICOLAS": Exit Function

    ' Hardware/Tools/Equipment
    If InStr(p, "FERRETERIA") > 0 Then GetCategory = "HERRAMIENTAS Y EQUIPAMIENTO": Exit Function
    If InStr(p, "AGROEQUIPOS") > 0 Then GetCategory = "HERRAMIENTAS Y EQUIPAMIENTO": Exit Function
    If InStr(p, "CONSTRUCCION") > 0 Or InStr(p, "MAQUINARIA") > 0 Then GetCategory = "HERRAMIENTAS Y EQUIPAMIENTO": Exit Function
    If InStr(d, "HERRAMIENTAS") > 0 Or InStr(d, "EQUIPO") > 0 Then GetCategory = "HERRAMIENTAS Y EQUIPAMIENTO": Exit Function

    ' Materials/Aggregates
    If InStr(p, "ARIDOS") > 0 Or InStr(p, "GRAVA") > 0 Or InStr(p, "ARENA") > 0 Then GetCategory = "MATERIALES DE CONSTRUCCION": Exit Function

    ' Office/Admin
    If InStr(p, "LIBRERIA") > 0 Or InStr(p, "TUCAN") > 0 Then GetCategory = "SUMINISTROS OFICINA": Exit Function
    If InStr(p, "MERCADOLIBRE") > 0 Or InStr(p, "OFFICE") > 0 Then GetCategory = "SUMINISTROS OFICINA": Exit Function
    If InStr(p, "NETEXPERTOS") > 0 Or InStr(p, "INFORMATICA") > 0 Then GetCategory = "SUMINISTROS OFICINA": Exit Function

    ' Rentals/Leasing
    If InStr(d, "ARRENDAMIENTO") > 0 Or InStr(d, "ARRIENDO") > 0 Then GetCategory = "ARRENDAMIENTO": Exit Function
    If InStr(d, "HOTEL") > 0 Or InStr(d, "HOSPEDAJE") > 0 Then GetCategory = "HOSPEDAJE": Exit Function

    ' Meals/Catering
    If InStr(p, "PULLMAN") > 0 Or InStr(p, "RESTAURANT") > 0 Or InStr(p, "CAFE") > 0 Then GetCategory = "ALIMENTACION": Exit Function

    ' Supplies/Commerce
    If InStr(p, "COMERCIAL") > 0 Or InStr(p, "COMERCIALIZADORA") > 0 Then GetCategory = "SUMINISTROS": Exit Function

    ' Default to SERVICIOS
    GetCategory = "SERVICIOS"
End Function

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"
Set wb = xl.Workbooks.Open(excelPath)
Set ws = wb.Sheets("Facturas")

lastRow = ws.UsedRange.Rows.Count
WScript.Echo "Processing " & lastRow & " rows..."

updateCount = 0
For row = 2 To lastRow
    if (row Mod 100 = 0) Then
        WScript.Echo "Row " & row & "..."
    End If

    cat = ws.Cells(row, 17).Value
    If cat = "REVISAR" Then
        proveedor = ws.Cells(row, 4).Value
        detalle = ws.Cells(row, 8).Value
        newCat = GetCategory(proveedor, detalle)
        ws.Cells(row, 17).Value = newCat
        updateCount = updateCount + 1
    End If
Next

WScript.Echo vbCrLf & "Updated " & updateCount & " invoices"
wb.Save
wb.Close False
xl.Quit

Set xl = Nothing
WScript.Echo "Done!"
