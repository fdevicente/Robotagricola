Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"
Set wb = xl.Workbooks.Open(excelPath)
Set ws = wb.Sheets("Facturas")

' Check first 50 REVISAR invoices
WScript.Echo "Sample REVISAR Invoices:"
count = 0
For row = 2 To 100
    cat = ws.Cells(row, 17).Value
    If cat = "REVISAR" Then
        count = count + 1
        prov = ws.Cells(row, 4).Value
        nro = ws.Cells(row, 7).Value
        WScript.Echo count & ". Fila " & row & ": " & prov & " (Nro: " & nro & ") -> " & cat
        If count >= 20 Then Exit For
    End If
Next

wb.Close False
xl.Quit

Set xl = Nothing
