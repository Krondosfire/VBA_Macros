Sub GenerateFinanceTable()
    Dim ws As Worksheet
    Dim numRows As Integer
    Dim numCols As Integer
    Dim i As Integer, j As Integer
    
    ' Create a new worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Finance Data"
    
    ' Define the number of rows and columns
    numRows = 100 ' You can change this to generate more rows
    numCols = 42 ' Increased to 42 to include stock name and market abbreviation
    
    ' Define headers for the table
    Dim headers As Variant
    headers = Array("ID", "Stock Name", "Market", "Stock Price", "Market Cap", "Volume", "PE Ratio", "EPS", _
                    "Dividend Yield", "Beta", "52-Week High", "52-Week Low", _
                    "Open Price", "Close Price", "Day Change %", "Sector Growth %", _
                    "Industry Growth %", "Revenue Growth %", "Net Profit Margin %", _
                    "Operating Margin %", "Debt-to-Equity Ratio", "Return on Equity %", _
                    "Return on Assets %", "Cash Flow (in millions)", "Assets (in billions)", _
                    "Liabilities (in billions)", "Share Outstanding (in millions)", _
                    "Free Cash Flow (in millions)", "Earnings Growth %", _
                    "Price-to-Sales Ratio", "Price-to-Book Ratio", _
                    "Forward PE Ratio", "Trailing PE Ratio", _
                    "Short Interest %", "Insider Ownership %", _
                    "Institutional Ownership %", "Gross Margin %", _
                    "Operating Expenses (in millions)", "Net Income (in millions)", _
                    "EBITDA (in millions)", "% Change in Volume", "% Change in Stock Price")
    
    ' Add headers to the worksheet
    For j = 1 To numCols
        ws.Cells(1, j).Value = headers(j - 1)
    Next j
    
    ' Define sample stock names and market abbreviations
    Dim stockNames As Variant
    stockNames = Array("Apple", "Microsoft", "Amazon", "Google", "Facebook", "Tesla", "Nvidia", "JPMorgan Chase", "Johnson & Johnson", "Visa")
    
    Dim marketAbbreviations As Variant
    marketAbbreviations = Array("NASDAQ", "NYSE", "NASDAQ", "NASDAQ", "NASDAQ", "NASDAQ", "NASDAQ", "NYSE", "NYSE", "NYSE")
    
    ' Populate the table with random data
    For i = 2 To numRows + 1 ' Start from row 2 to leave row 1 for headers
        ws.Cells(i, 1).Value = i - 1 ' Set ID as primary key
        ws.Cells(i, 2).Value = stockNames((i - 2) Mod 10) ' Set stock name
        ws.Cells(i, 3).Value = marketAbbreviations((i - 2) Mod 10) ' Set market abbreviation
        
        ' Fill the rest of the columns with random numbers
        For j = 4 To numCols
            ws.Cells(i, j).Value = WorksheetFunction.RandBetween(1, 1000) / 10 ' Random numbers between 0.1 and 100.0
        Next j
    Next i
    
    ' Format the table
    With ws.Range(ws.Cells(1, 1), ws.Cells(numRows + 1, numCols))
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Arial"
        .Font.Size = 10
        .Columns.AutoFit
    End With
    
    ' Format the header row for better readability
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    MsgBox numRows & "-row table with finance-related data generated successfully!", vbInformation
End Sub
