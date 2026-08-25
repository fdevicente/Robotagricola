Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"
Set wb = xl.Workbooks.Open(excelPath)
Set ws = wb.Sheets("Facturas")

WScript.Echo "Column Headers:"
For col = 1 To 25
    header = ws.Cells(1, col).Value
    WScript.Echo "Col " & col & ": " & header
Next

WScript.Echo vbCrLf & "Sample data from row 6:"
For col = 1 To 10
    value = ws.Cells(6, col).Value
    WScript.Echo "Col " & col & ": " & value
Next

WScript.Echo vbCrLf & "Row 6, Column 17 details:"
Set cell = ws.Cells(6, 17)
WScript.Echo "Value: " & cell.Value
WScript.Echo "Formula: " & cell.Formula
WScript.Echo "Text: " & cell.Text

wb.Close False
xl.Quit
Set xl = Nothing
