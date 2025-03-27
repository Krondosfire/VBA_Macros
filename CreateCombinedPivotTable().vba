Sub CreateCombinedPivotTable()
    Dim wsPivot As Worksheet
    Dim pvtCache As PivotCache
    Dim pvtTable As PivotTable
    Dim connection As WorkbookConnection
    
    ' Create new worksheet for pivot table
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "CombinedAnalysis"
    
    ' Create PivotCache from Data Model
    Set pvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlExternal, _
        SourceData:=ThisWorkbook.Connections("ThisWorkbookDataModel"))
    
    ' Create PivotTable
    Set pvtTable = pvtCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="CombinedStockAnalysis")
    
    ' Add fields from multiple tables
    With pvtTable
        ' Row Fields
        .AddRowField .PivotFields("[StockInfo].[StockSymbol]")
        .AddRowField .PivotFields("[StockInfo].[CompanyName]")
        .AddRowField .PivotFields("[StockInfo].[Sector]")
        .AddRowField .PivotFields("[StockInfo].[Industry]")
        
        ' Column Fields
        .AddColumnField .PivotFields("[DailyPrices].[Date]")
        .AddColumnField .PivotFields("[FinancialMetrics].[Year]")
        
        ' Value Fields
        .AddDataField .PivotFields("[DailyPrices].[OpenPrice]"), "Avg Open Price", xlAverage
        .AddDataField .PivotFields("[DailyPrices].[ClosePrice]"), "Avg Close Price", xlAverage
        .AddDataField .PivotFields("[FinancialMetrics].[Revenue]"), "Total Revenue", xlSum
        .AddDataField .PivotFields("[FinancialMetrics].[NetIncome]"), "Total Net Income", xlSum
        .AddDataField .PivotFields("[FinancialMetrics].[EPS]"), "Avg EPS", xlAverage
        
        ' Filter Fields
        .PivotFields("[StockInfo].[Sector]").Orientation = xlPageField
        .PivotFields("[FinancialMetrics].[Year]").Orientation = xlPageField
    End With
    
    ' Format the PivotTable
    With pvtTable
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .ShowTableStyleRowStripes = True
        .TableStyle2 = "PivotStyleMedium9"
    End With
    
    ' Autofit columns for better visibility
    wsPivot.Columns.AutoFit
    
    MsgBox "Combined Pivot Table created successfully!", vbInformation
End Sub
