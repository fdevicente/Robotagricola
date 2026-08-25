Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"
Set wb = xl.Workbooks.Open(excelPath)
Set ws = wb.Sheets("Facturas")

' Find first REVISAR row
For row = 2 To 100
    cat = ws.Cells(row, 17).Value
    If cat = "REVISAR" Then
        WScript.Echo "Found REVISAR at row " & row
        WScript.Echo "Current value: " & ws.Cells(row, 17).Value

        ' Update it
        ws.Cells(row, 17).Value = "TEST_UPDATED"
        WScript.Echo "Updated to: " & ws.Cells(row, 17).Value

        ' Save
        wb.Save
        WScript.Echo "File saved"
        Exit For
    End If
Next

wb.Close False
xl.Quit
Set xl = Nothing

' Now reopen and check
WScript.Sleep 1000
Set xl2 = CreateObject("Excel.Application")
xl2.Visible = False
Set wb2 = xl2.Workbooks.Open(excelPath)
Set ws2 = wb2.Sheets("Facturas")
WScript.Echo "Reopened file"
WScript.Echo "Value at that row: " & ws2.Cells(row, 17).Value

wb2.Close False
xl2.Quit
Set xl2 = Nothing
