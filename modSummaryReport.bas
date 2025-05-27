Attribute VB_Name = "Module4"
Sub GenerateSummaryReport()
    Dim wsData As Worksheet, wsSummary As Worksheet
    Dim lastRow As Long, i As Long
    Dim profitCol As Long, productCol As Long, statusCol As Long
    Dim totalProfit As Double
    Dim dict As Object
    Dim productName As String
    Dim profit As Double
    Dim summaryRow As Long
    
    Set wsData = ThisWorkbook.Sheets("SalesData")
    
    ' Delete existing Summary sheet if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsSummary = ThisWorkbook.Sheets.Add
    wsSummary.Name = "Summary"
    
    productCol = 3 ' Product Name column
    profitCol = 9  ' Profit column
    statusCol = 8  ' Status column
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    Set dict = CreateObject("Scripting.Dictionary")
    totalProfit = 0
    
    ' Accumulate profit per product
    For i = 2 To lastRow
        If wsData.Cells(i, statusCol).Value = "Valid" Then
            productName = wsData.Cells(i, productCol).Value
            profit = wsData.Cells(i, profitCol).Value
            
            totalProfit = totalProfit + profit
            
            If dict.Exists(productName) Then
                dict(productName) = dict(productName) + profit
            Else
                dict.Add productName, profit
            End If
        End If
    Next i
    
    With wsSummary
        .Range("A1").Value = "Summary Report"
        .Range("A1").Font.Bold = True
        
        .Range("A2").Value = "Product Name"
        .Range("B2").Value = "Total Profit"
        .Range("A2:B2").Font.Bold = True
        
        summaryRow = 3
        Dim key As Variant
        For Each key In dict.Keys
            .Cells(summaryRow, 1).Value = key
            .Cells(summaryRow, 2).Value = dict(key)
            .Cells(summaryRow, 2).NumberFormat = "#,##0.00"
            summaryRow = summaryRow + 1
        Next key
        
        ' Leave one blank row then total profit
        summaryRow = summaryRow + 1
        .Cells(summaryRow, 1).Value = "Total Profit:"
        .Cells(summaryRow, 1).Font.Bold = True
        .Cells(summaryRow, 2).Value = totalProfit
        .Cells(summaryRow, 2).NumberFormat = "#,##0.00"
        
        .Columns("A:B").AutoFit
    End With
    
    MsgBox "Summary report generated in 'Summary'"
End Sub
