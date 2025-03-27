Sub ExtractAndAggregateData()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set reference to source worksheet
    Set wsSource = ThisWorkbook.Sheets("StockMarketData")
    
    ' Create new target worksheet
    Set wsTarget = ThisWorkbook.Sheets.Add
    wsTarget.Name = "AggregatedData"
    
    ' Set up headers in the new sheet
    With wsTarget
        .Cells(1, 1).Value = "Stock ID"
        .Cells(1, 2).Value = "Stock Symbol"
        .Cells(1, 3).Value = "Company Name"
        .Cells(1, 4).Value = "Sector"
        .Cells(1, 5).Value = "Industry"
        .Cells(1, 6).Value = "Average Close Price"
        .Cells(1, 7).Value = "Latest Revenue"
        .Cells(1, 8).Value = "Latest Net Income"
        .Cells(1, 9).Value = "Latest EPS"
    End With
    
    ' Find last row in StockInfo table
    lastRow = wsSource.ListObjects("StockInfo").Range.Rows.Count
    
    ' Loop through StockInfo and extract data
    For i = 2 To lastRow ' Assuming headers in row 1
        With wsTarget
            .Cells(i, 1).Value = wsSource.Cells(i + 1, 1).Value ' Stock ID
            .Cells(i, 2).Value = wsSource.Cells(i + 1, 2).Value ' Stock Symbol
            .Cells(i, 3).Value = wsSource.Cells(i + 1, 3).Value ' Company Name
            .Cells(i, 4).Value = wsSource.Cells(i + 1, 4).Value ' Sector
            .Cells(i, 5).Value = wsSource.Cells(i + 1, 5).Value ' Industry
            
            ' Average Close Price (from DailyPrices)
            .Cells(i, 6).Formula = "=AVERAGEIF(DailyPrices[StockID],RC[-5],DailyPrices[ClosePrice])"
            
            ' Latest Financial Metrics
            .Cells(i, 7).Formula = "=MAXIFS(FinancialMetrics[Revenue],FinancialMetrics[StockID],RC[-6],FinancialMetrics[Year],MAX(FinancialMetrics[Year]))"
            .Cells(i, 8).Formula = "=MAXIFS(FinancialMetrics[NetIncome],FinancialMetrics[StockID],RC[-7],FinancialMetrics[Year],MAX(FinancialMetrics[Year]))"
            .Cells(i, 9).Formula = "=MAXIFS(FinancialMetrics[EPS],FinancialMetrics[StockID],RC[-8],FinancialMetrics[Year],MAX(FinancialMetrics[Year]))"
        End With
    Next i
    
    ' Format as table
    wsTarget.ListObjects.Add(xlSrcRange, wsTarget.Range("A1").CurrentRegion, , xlYes).Name = "AggregatedData"
    wsTarget.ListObjects("AggregatedData").TableStyle = "TableStyleMedium6"
    
    ' Autofit columns
    wsTarget.Columns.AutoFit
    
    MsgBox "Data extraction and aggregation complete!", vbInformation
End Sub
