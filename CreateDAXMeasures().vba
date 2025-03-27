Sub CreateDAXMeasures()
    Dim mdl As Model
    Set mdl = ThisWorkbook.Model
    
    ' Price to Earnings Growth (PEG) Ratio
    mdl.AddMeasure "PEG Ratio", "DIVIDE([PE Ratio], [Earnings Growth %])", "MarketData"
    
    ' Return on Investment (ROI)
    mdl.AddMeasure "ROI", "DIVIDE([Net Income (in millions)], [Assets (in billions)] * 1000)", "FinanceData"
    
    ' Debt to EBITDA Ratio
    mdl.AddMeasure "Debt to EBITDA", "DIVIDE([Liabilities (in billions)], [EBITDA (in millions)] / 1000)", "FinanceData"
    
    ' Year-over-Year Growth
    mdl.AddMeasure "YoY Growth", "([Stock Price] - CALCULATE([Stock Price], DATEADD('Date'[Date], -1, YEAR))) / CALCULATE([Stock Price], DATEADD('Date'[Date], -1, YEAR))", "MarketData"
    
    MsgBox "DAX measures created successfully!", vbInformation
End Sub
