Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"

' Make a backup first
Set fso = CreateObject("Scripting.FileSystemObject")
backupPath = excelPath & ".backup"
If fso.FileExists(backupPath) Then
    fso.DeleteFile(backupPath)
End If
fso.CopyFile excelPath, backupPath
WScript.Echo "Backup created: " & backupPath

Set wb = xl.Workbooks.Open(excelPath, , False)  ' Not read-only
Set ws = wb.Sheets("Facturas")

' Find first REVISAR row
rowToUpdate = -1
For row = 2 To 100
    cat = ws.Cells(row, 17).Value
    If cat = "REVISAR" Then
        rowToUpdate = row
        WScript.Echo "Found REVISAR at row " & row
        WScript.Echo "Provider: " & ws.Cells(row, 4).Value
        Exit For
    End If
Next

If rowToUpdate > 0 Then
    WScript.Echo "Current value: " & ws.Cells(rowToUpdate, 17).Value

    ' Update it
    ws.Cells(rowToUpdate, 17).Value = "TEST_UPDATED_V2"
    WScript.Echo "Set to: " & ws.Cells(rowToUpdate, 17).Value

    ' Save using application SaveWorkbooks
    xl.Save
    WScript.Echo "Application saved"

    ' Also try workbook save
    On Error Resume Next
    wb.Save
    If Err.Number <> 0 Then
        WScript.Echo "Workbook save error: " & Err.Description
    Else
        WScript.Echo "Workbook saved"
    End If
    On Error Goto 0
End If

wb.Close True
xl.DisplayAlerts = True
xl.Quit
Set xl = Nothing

WScript.Sleep 2000
WScript.Echo "Checking result..."

' Reopen and verify
Set xl2 = CreateObject("Excel.Application")
xl2.Visible = False
xl2.DisplayAlerts = False
Set wb2 = xl2.Workbooks.Open(excelPath, , False)
Set ws2 = wb2.Sheets("Facturas")

If rowToUpdate > 0 Then
    WScript.Echo "Value after reopen: " & ws2.Cells(rowToUpdate, 17).Value
    If ws2.Cells(rowToUpdate, 17).Value = "TEST_UPDATED_V2" Then
        WScript.Echo "SUCCESS - Changes were saved!"
    Else
        WScript.Echo "FAILED - Changes were not saved"
    End If
End If

wb2.Close False
xl2.Quit
Set xl2 = Nothing
