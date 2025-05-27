Attribute VB_Name = "Module3"
Sub CalculateProfit()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim profitCol As Long
    Dim unitCost As Double, unitPrice As Double, qty As Double
    
    Set ws = ThisWorkbook.Sheets("SalesData")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    profitCol = 9 ' Column I
    
    ' Add Profit header if not exists
    If ws.Cells(1, profitCol).Value <> "Profit" Then
        ws.Cells(1, profitCol).Value = "Profit"
        ws.Cells(1, profitCol).Font.Bold = True
    End If
    
    For i = 2 To lastRow
        If ws.Cells(i, 8).Value = "Valid" Then ' Only for valid rows
            qty = ws.Cells(i, 5).Value
            unitCost = ws.Cells(i, 6).Value
            unitPrice = ws.Cells(i, 7).Value
            ws.Cells(i, profitCol).Value = Round((unitPrice - unitCost) * qty, 2)
        Else
            ws.Cells(i, profitCol).Value = ""
        End If
    Next i
    
    ws.Columns(profitCol).AutoFit
    MsgBox "Profit calculation complete.", vbInformation
End Sub

