Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False

excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"
Set wb = xl.Workbooks.Open(excelPath, , False, 1)
Set ws = wb.Sheets("Facturas")

WScript.Echo "Workbook opened in Excel (visible window)"
WScript.Sleep 2000

' Find first REVISAR
For row = 2 To 100
    cat = ws.Cells(row, 17).Value
    If cat = "REVISAR" Then
        WScript.Echo "Found REVISAR at row " & row
        WScript.Echo "Original value: " & ws.Cells(row, 17).Value

        ' Update
        ws.Cells(row, 17).Value = "TEST_VISIBLE"
        WScript.Echo "Updated to: " & ws.Cells(row, 17).Value

        ' Give Excel time to process
        WScript.Sleep 1000

        ' Try Ctrl+S through the application
        xl.Run "StandardModule.Workbook_Save"
        WScript.Sleep 1000

        Exit For
    End If
Next

' Now close Excel and check
WScript.Echo "Closing Excel..."
wb.Close True
xl.Quit
Set xl = Nothing

WScript.Sleep 2000

' Check result
Set xl2 = CreateObject("Excel.Application")
xl2.Visible = False
Set wb2 = xl2.Workbooks.Open(excelPath)
Set ws2 = wb2.Sheets("Facturas")

For row2 = 2 To 100
    cat2 = ws2.Cells(row2, 17).Value
    If ws2.Cells(row2, 4).Value = "FERRETERIA BASCU" & Chr(154) & "AN" Then
        WScript.Echo "After closing and reopening: " & cat2
        Exit For
    End If
Next

wb2.Close False
xl2.Quit
Set xl2 = Nothing
