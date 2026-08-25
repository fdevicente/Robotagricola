Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"
Set wb = xl.Workbooks.Open(excelPath)
Set ws = wb.Sheets("Facturas")

lastRow = ws.UsedRange.Rows.Count
WScript.Echo "Total rows: " & lastRow

' Count REVISAR by provider
Set revisar_providers = CreateObject("Scripting.Dictionary")
totalRevisar = 0

For row = 2 To lastRow
    if (row Mod 100 = 0) Then
        WScript.Echo "Processing row " & row
    End If

    cat = ws.Cells(row, 17).Value
    If cat = "REVISAR" Then
        totalRevisar = totalRevisar + 1
        prov = ws.Cells(row, 4).Value
        If prov = "" Then prov = "UNKNOWN"

        total = ws.Cells(row, 16).Value
        If Not IsNumeric(total) Then total = 0

        If Not revisar_providers.Exists(prov) Then
            revisar_providers.Add prov, Array(0, 0)
        End If

        arr = revisar_providers(prov)
        arr(0) = arr(0) + 1
        arr(1) = arr(1) + total
        revisar_providers(prov) = arr
    End If
Next

WScript.Echo vbCrLf & "Total REVISAR: " & totalRevisar & vbCrLf

' Show top providers
' (VBScript dictionary iteration)
For Each prov In revisar_providers
    arr = revisar_providers(prov)
    WScript.Echo prov & ": " & arr(0) & " items, $" & arr(1) & " CLP"
Next

wb.Close False
xl.Quit

Set xl = Nothing
