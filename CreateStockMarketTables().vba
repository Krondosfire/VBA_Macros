Sub CreateStockMarketTables()
    Dim ws As Worksheet
    Dim i As Long
    Dim rng As Range
    
    ' Create a new worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "StockMarketData"
    
    ' Create StockInfo table
    ws.Range("A1").Value = "StockInfo"
    ws.Range("A2").Resize(1, 5).Value = Array("ID", "StockSymbol", "CompanyName", "Sector", "Industry")
    Set rng = ws.Range("A3:E102")
    rng.Value = GenerateStockInfo(100)
    ws.ListObjects.Add(xlSrcRange, rng.Offset(-1), , xlYes).Name = "StockInfo"
    
    ' Create DailyPrices table
    ws.Range("G1").Value = "DailyPrices"
    ws.Range("G2").Resize(1, 5).Value = Array("ID", "StockID", "Date", "OpenPrice", "ClosePrice")
    Set rng = ws.Range("G3:K1002")
    rng.Value = GenerateDailyPrices(1000)
    ws.ListObjects.Add(xlSrcRange, rng.Offset(-1), , xlYes).Name = "DailyPrices"
    
    ' Create FinancialMetrics table
    ws.Range("M1").Value = "FinancialMetrics"
    ws.Range("M2").Resize(1, 6).Value = Array("ID", "StockID", "Year", "Revenue", "NetIncome", "EPS")
    Set rng = ws.Range("M3:R402")
    rng.Value = GenerateFinancialMetrics(400)
    ws.ListObjects.Add(xlSrcRange, rng.Offset(-1), , xlYes).Name = "FinancialMetrics"
    
    ' Format tables
    ws.ListObjects("StockInfo").TableStyle = "TableStyleMedium2"
    ws.ListObjects("DailyPrices").TableStyle = "TableStyleMedium3"
    ws.ListObjects("FinancialMetrics").TableStyle = "TableStyleMedium4"
    
    ws.Columns.AutoFit
    
    MsgBox "Stock market tables created successfully!", vbInformation
End Sub

Function GenerateStockInfo(rows As Long) As Variant
    Dim data() As Variant
    Dim i As Long
    Dim sectors As Variant
    Dim industries As Variant
    
    ReDim data(1 To rows, 1 To 5)
    sectors = Array("Technology", "Healthcare", "Finance", "Consumer Goods", "Energy")
    industries = Array("Software", "Pharmaceuticals", "Banking", "Retail", "Oil & Gas")
    
    For i = 1 To rows
        data(i, 1) = i
        data(i, 2) = "STOCK" & Format(i, "000")
        data(i, 3) = "Company " & i
        data(i, 4) = sectors(i Mod 5)
        data(i, 5) = industries(i Mod 5)
    Next i
    
    GenerateStockInfo = data
End Function

Function GenerateDailyPrices(rows As Long) As Variant
    Dim data() As Variant
    Dim i As Long
    
    ReDim data(1 To rows, 1 To 5)
    
    For i = 1 To rows
        data(i, 1) = i
        data(i, 2) = WorksheetFunction.RandBetween(1, 100)
        data(i, 3) = DateSerial(2023, WorksheetFunction.RandBetween(1, 12), WorksheetFunction.RandBetween(1, 28))
        data(i, 4) = Round(WorksheetFunction.RandBetween(10, 1000) + Rnd(), 2)
        data(i, 5) = Round(data(i, 4) * (1 + (Rnd() - 0.5) / 10), 2)
    Next i
    
    GenerateDailyPrices = data
End Function

Function GenerateFinancialMetrics(rows As Long) As Variant
    Dim data() As Variant
    Dim i As Long
    
    ReDim data(1 To rows, 1 To 6)
    
    For i = 1 To rows
        data(i, 1) = i
        data(i, 2) = WorksheetFunction.RandBetween(1, 100)
        data(i, 3) = WorksheetFunction.RandBetween(2018, 2023)
        data(i, 4) = Round(WorksheetFunction.RandBetween(100000, 10000000) / 1000, 0) * 1000
        data(i, 5) = Round(data(i, 4) * WorksheetFunction.RandBetween(5, 20) / 100, 0)
        data(i, 6) = Round(data(i, 5) / WorksheetFunction.RandBetween(1000000, 10000000) * 1000, 2)
    Next i
    
    GenerateFinancialMetrics = data
End Function
