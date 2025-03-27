Sub UpdateStockInfoAgain()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim stockData As Variant
    
    ' Set reference to the worksheet containing the StockInfo table
    Set ws = ThisWorkbook.Sheets("StockMarketData")
    
    ' Set reference to the StockInfo table
    Set tbl = ws.ListObjects("StockInfo")
    
    ' Debug: Print table name and column names
    Debug.Print "Table name: " & tbl.Name
    Debug.Print "Columns:"
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        Debug.Print col.Name
    Next col
    
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
        On Error Resume Next
        tbl.ListColumns("StockSymbol").DataBodyRange(i) = stockData((i - 1) Mod 10)(0)
        tbl.ListColumns("CompanyName").DataBodyRange(i) = stockData((i - 1) Mod 10)(1)
        tbl.ListColumns("Sector").DataBodyRange(i) = stockData((i - 1) Mod 10)(2)
        tbl.ListColumns("Industry").DataBodyRange(i) = stockData((i - 1) Mod 10)(3)
        If Err.Number <> 0 Then
            Debug.Print "Error on row " & i & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    
    MsgBox "StockInfo table updated with realistic values!", vbInformation
End Sub
