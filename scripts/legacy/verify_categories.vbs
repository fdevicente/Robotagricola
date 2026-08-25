Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

excelPath = "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx"
Set wb = xl.Workbooks.Open(excelPath)
Set ws = wb.Sheets("Facturas")

lastRow = ws.UsedRange.Rows.Count

' Count categories
Set categories = CreateObject("Scripting.Dictionary")
totalCount = 0
totalAmount = 0

For row = 2 To lastRow
    cat = ws.Cells(row, 17).Value
    total = ws.Cells(row, 16).Value
    If Not IsNumeric(total) Then total = 0

    If cat <> "" And cat <> "REVISAR" Then
        If Not categories.Exists(cat) Then
            categories.Add cat, Array(0, 0)
        End If

        arr = categories(cat)
        arr(0) = arr(0) + 1
        arr(1) = arr(1) + total
        categories(cat) = arr

        totalCount = totalCount + 1
        totalAmount = totalAmount + total
    End If
Next

WScript.Echo "Categorized Invoices Distribution:" & vbCrLf

' Sort by amount
Dim catArray()
ReDim catArray(categories.Count - 1)
i = 0
For Each cat In categories
    Set catArray(i) = CreateObject("Scripting.Dictionary")
    catArray(i).Add "name", cat
    arr = categories(cat)
    catArray(i).Add "count", arr(0)
    catArray(i).Add "amount", arr(1)
    i = i + 1
Next

' Bubble sort by amount
For i = 0 To UBound(catArray)
    For j = 0 To UBound(catArray) - 1
        If catArray(j)("amount") < catArray(j+1)("amount") Then
            Set temp = catArray(j)
            Set catArray(j) = catArray(j+1)
            Set catArray(j+1) = temp
        End If
    Next
Next

For i = 0 To UBound(catArray)
    WScript.Echo catArray(i)("name") & ": " & catArray(i)("count") & " items, $" & catArray(i)("amount") & " CLP"
Next

WScript.Echo vbCrLf & "TOTAL: " & totalCount & " categorized invoices, $" & totalAmount & " CLP"

' Check for remaining REVISAR
revisar_count = 0
For row = 2 To lastRow
    cat = ws.Cells(row, 17).Value
    If cat = "REVISAR" Then
        revisar_count = revisar_count + 1
    End If
Next

WScript.Echo "Remaining REVISAR: " & revisar_count

wb.Close False
xl.Quit

Set xl = Nothing
