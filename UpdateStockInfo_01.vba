Sub UpdateStockInfo()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim stockSymbols As Variant
    Dim companyNames As Variant
    
    ' Set reference to the worksheet containing the StockInfo table
    Set ws = ThisWorkbook.Sheets("StockMarketData")
    
    ' Set reference to the StockInfo table
    Set tbl = ws.ListObjects("StockInfo")
    
    ' Define arrays of realistic stock symbols and company names
    stockSymbols = Array("AAPL", "MSFT", "AMZN", "GOOGL", "FB", "TSLA", "JPM", "JNJ", "V", "PG")
    companyNames = Array("Apple Inc.", "Microsoft Corporation", "Amazon.com Inc.", "Alphabet Inc.", "Meta Platforms Inc.", "Tesla Inc.", "JPMorgan Chase & Co.", "Johnson & Johnson", "Visa Inc.", "Procter & Gamble Company")
    
    ' Update the StockSymbol and CompanyName columns
    For i = 1 To WorksheetFunction.Min(tbl.ListRows.Count, UBound(stockSymbols) + 1)
        tbl.ListColumns("StockSymbol").DataBodyRange(i) = stockSymbols(i - 1)
        tbl.ListColumns("CompanyName").DataBodyRange(i) = companyNames(i - 1)
    Next i
    
    MsgBox "StockInfo table updated with realistic values!", vbInformation
End Sub
