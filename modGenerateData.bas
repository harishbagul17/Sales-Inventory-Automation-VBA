Attribute VB_Name = "Module1"
Sub GenerateSalesData()
    Dim ws As Worksheet
    Dim i As Long, n As Long
    Dim productIDs As Variant, productNames As Variant, categories As Variant
    Dim randProd As Long
    Dim startDate As Date, endDate As Date
    Dim saleDate As Date
    
    ' Create or clear SalesData sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SalesData")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "SalesData"
    Else
        ws.Cells.Clear
    End If
    
    ' Headers (including Category and Sale Date after Product ID)
    ws.Range("A1").Value = "Product ID"
    ws.Range("B1").Value = "Sale Date"
    ws.Range("C1").Value = "Product Name"
    ws.Range("D1").Value = "Category"
    ws.Range("E1").Value = "Quantity Sold"
    ws.Range("F1").Value = "Unit Cost"
    ws.Range("G1").Value = "Unit Sale Price"
    
    ' Bold headers
    ws.Range("A1:G1").Font.Bold = True
    
    ' Sample grocery items and categories
    productIDs = Array("PRD001", "PRD002", "PRD003", "PRD004", "PRD005", "PRD006", "PRD007", "PRD008", "PRD009", "PRD010")
    productNames = Array("Milk", "Bread", "Eggs", "Rice", "Sugar", "Salt", "Butter", "Tea", "Coffee", "Cheese")
    categories = Array("Dairy", "Bakery", "Poultry", "Grains", "Sweeteners", "Spices", "Dairy", "Beverages", "Beverages", "Dairy")
    
    ' Number of rows (random between 100 and 200)
    Randomize
    n = Int((200 - 100 + 1) * Rnd + 100)
    
    startDate = DateSerial(2024, 1, 1)
    endDate = DateSerial(2024, 4, 30)
    
    For i = 2 To n + 1
        randProd = Int((UBound(productIDs) - LBound(productIDs) + 1) * Rnd + LBound(productIDs))
        
        ws.Cells(i, 1).Value = productIDs(randProd)
        ' Random sale date between startDate and endDate
        saleDate = startDate + Int((endDate - startDate + 1) * Rnd)
        ws.Cells(i, 2).Value = saleDate
        ws.Cells(i, 2).NumberFormat = "yyyy-mm-dd"
        
        ws.Cells(i, 3).Value = productNames(randProd)
        ws.Cells(i, 4).Value = categories(randProd)
        ws.Cells(i, 5).Value = Int(10 * Rnd + 1) ' Quantity Sold between 1 and 10
        ws.Cells(i, 6).Value = Round(5 + 15 * Rnd, 2) ' Unit Cost between 5 and 20
        ws.Cells(i, 7).Value = Round(ws.Cells(i, 6).Value * (1.1 + 0.5 * Rnd), 2) ' Unit Sale Price 10%-60% higher than cost
    Next i
    
    ws.Columns("A:G").AutoFit
    MsgBox "Sales data generated with " & n & " rows.", vbInformation
End Sub



