Sub UpdateStockInfo()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim stockData As Variant
    
    ' Set reference to the worksheet containing the StockInfo table
    Set ws = ThisWorkbook.Sheets("StockMarketData")
    
    ' Set reference to the StockInfo table
    Set tbl = ws.ListObjects("StockInfo")
    
    ' Define array of realistic stock data (Symbol, Company, Sector, Industry)
    stockData = Array( _
        Array("AAPL", "Apple Inc.", "Technology", "Consumer Electronics"), _
        Array("MSFT", "Microsoft Corporation", "Technology", "Software"), _
        Array("AMZN", "Amazon.com Inc.", "Consumer Cyclical", "Internet Retail"), _
        Array("GOOGL", "Alphabet Inc.", "Communication Services", "Internet Content & Information"), _
        Array("FB", "Meta Platforms Inc.", "Communication Services", "Internet Content & Information"), _
        Array("TSLA", "Tesla Inc.", "Consumer Cyclical", "Auto Manufacturers"), _
        Array("JPM", "JPMorgan Chase & Co.", "Financial Services", "Banks"), _
        Array("JNJ", "Johnson & Johnson", "Healthcare", "Drug Manufacturers"), _
        Array("V", "Visa Inc.", "Financial Services", "Credit Services"), _
        Array("PG", "Procter & Gamble Company", "Consumer Defensive", "Household & Personal Products"))
    
    ' Update the StockSymbol, CompanyName, Sector, and Industry columns
    For i = 1 To tbl.ListRows.Count
        tbl.ListColumns("StockSymbol").DataBodyRange(i) = stockData((i - 1) Mod 10 + 1)(0)
        tbl.ListColumns("CompanyName").DataBodyRange(i) = stockData((i - 1) Mod 10 + 1)(1)
        tbl.ListColumns("Sector").DataBodyRange(i) = stockData((i - 1) Mod 10 + 1)(2)
        tbl.ListColumns("Industry").DataBodyRange(i) = stockData((i - 1) Mod 10 + 1)(3)
    Next i
    
    MsgBox "StockInfo table updated with realistic values!", vbInformation
End Sub
