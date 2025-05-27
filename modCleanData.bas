Attribute VB_Name = "Module2"
Sub CleanSalesData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim qty As Variant, cost As Variant, price As Variant
    Dim statusCol As Long
    
    Set ws = ThisWorkbook.Sheets("SalesData")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Add Status column header at H1 if not exists
    statusCol = 8 ' Column H
    If ws.Cells(1, statusCol).Value <> "Status" Then
        ws.Cells(1, statusCol).Value = "Status"
        ws.Cells(1, statusCol).Font.Bold = True
    End If
    
    For i = 2 To lastRow
        ' Trim and standardize text fields
        ws.Cells(i, 1).Value = UCase(Trim(ws.Cells(i, 1).Value)) ' Product ID uppercase
        ws.Cells(i, 3).Value = Application.WorksheetFunction.Proper(Trim(ws.Cells(i, 3).Value)) ' Product Name proper case
        ws.Cells(i, 4).Value = Application.WorksheetFunction.Proper(Trim(ws.Cells(i, 4).Value)) ' Category proper case
        
        ' Validate numeric fields
        qty = ws.Cells(i, 5).Value
        cost = ws.Cells(i, 6).Value
        price = ws.Cells(i, 7).Value
        
        If IsNumeric(qty) And IsNumeric(cost) And IsNumeric(price) Then
            If qty > 0 And cost > 0 And price > 0 Then
                ' Validate sale date
                If IsDate(ws.Cells(i, 2).Value) Then
                    ws.Cells(i, statusCol).Value = "Valid"
                Else
                    ws.Cells(i, statusCol).Value = "Invalid: Date"
                End If
            Else
                ws.Cells(i, statusCol).Value = "Invalid: Negative or zero"
            End If
        Else
            ws.Cells(i, statusCol).Value = "Invalid: Non-numeric"
        End If
    Next i
    
    MsgBox "Data cleaning complete. Check 'Status' column for validation results.", vbInformation
End Sub

