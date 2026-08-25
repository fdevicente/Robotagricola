Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"

Set wb = xl.Workbooks.Open(excelPath)
Set ws = wb.Sheets("Facturas")

WScript.Echo "File opened"

' Find first REVISAR row
rowToUpdate = -1
For row = 2 To 100
    cat = ws.Cells(row, 17).Value
    If cat = "REVISAR" Then
        rowToUpdate = row
        WScript.Echo "Found REVISAR at row " & row
        Exit For
    End If
Next

If rowToUpdate > 0 Then
    WScript.Echo "Current value: '" & ws.Cells(rowToUpdate, 17).Value & "'"

    ' Update it
    ws.Cells(rowToUpdate, 17).Value = "UPDATED_V3"
    WScript.Echo "Set to: '" & ws.Cells(rowToUpdate, 17).Value & "'"

    ' Try SaveAs to force save
    On Error Resume Next
    tempPath = excelPath & ".tmp"
    wb.SaveAs tempPath
    If Err.Number <> 0 Then
        WScript.Echo "SaveAs error: " & Err.Number & " - " & Err.Description
    Else
        WScript.Echo "SaveAs succeeded to: " & tempPath
        ' Now close without saving and replace original
        wb.Close False
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile excelPath
        fso.MoveFile tempPath, excelPath
        WScript.Echo "File replaced"
    End If
    On Error Goto 0
Else
    wb.Close False
End If

xl.Quit
Set xl = Nothing

WScript.Sleep 1000

' Verify
Set xl2 = CreateObject("Excel.Application")
xl2.Visible = False
xl2.DisplayAlerts = False
Set wb2 = xl2.Workbooks.Open(excelPath)
Set ws2 = wb2.Sheets("Facturas")

If rowToUpdate > 0 Then
    WScript.Echo "After reopen: '" & ws2.Cells(rowToUpdate, 17).Value & "'"
End If

wb2.Close False
xl2.Quit
Set xl2 = Nothing
