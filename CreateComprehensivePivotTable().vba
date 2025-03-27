Sub CreateComprehensivePivotTable()
    Dim wsPivot As Worksheet
    Dim pvtCache As PivotCache
    Dim pvtTable As PivotTable
    
    ' Create a new worksheet for the pivot table
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "ComprehensivePivot"
    
    ' Create PivotCache from the Data Model
    On Error Resume Next
    Set pvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlExternal, _
        SourceData:="ThisWorkbookDataModel")
    If Err.Number <> 0 Then
        MsgBox "Error creating PivotCache. Ensure tables are added to the Data Model.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Create PivotTable
    Set pvtTable = pvtCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="ComprehensiveStockAnalysis")
    
    ' Add fields to the PivotTable
    With pvtTable
        ' Row Fields (from StockInfo table)
        On Error Resume Next
        .PivotFields("[StockInfo].[StockSymbol]").Orientation = xlRowField
        .PivotFields("[StockInfo].[CompanyName]").Orientation = xlRowField
        .PivotFields("[StockInfo].[Sector]").Orientation = xlRowField
        .PivotFields("[StockInfo].[Industry]").Orientation = xlRowField
        
        If Err.Number <> 0 Then
            MsgBox "Error adding row fields. Ensure field names are correct and exist in the Data Model.", vbExclamation
            Exit Sub
        End If
        
        ' Column Fields (from DailyPrices and FinancialMetrics tables)
        On Error Resume Next
        .PivotFields("[DailyPrices].[Date]").Orientation = xlColumnField
        .PivotFields("[FinancialMetrics].[Year]").Orientation = xlColumnField
        
        If Err.Number <> 0 Then
            MsgBox "Error adding column fields. Ensure field names are correct and exist in the Data Model.", vbExclamation
            Exit Sub
        End If
        
        ' Value Fields (aggregations)
        On Error Resume Next
        .AddDataField .PivotFields("[DailyPrices].[OpenPrice]"), "Avg Open Price", xlAverage
        .AddDataField .PivotFields("[DailyPrices].[ClosePrice]"), "Avg Close Price", xlAverage
        .AddDataField .PivotFields("[FinancialMetrics].[Revenue]"), "Total Revenue", xlSum
        .AddDataField .PivotFields("[FinancialMetrics].[NetIncome]"), "Total Net Income", xlSum
        .AddDataField .PivotFields("[FinancialMetrics].[EPS]"), "Avg EPS", xlAverage
        
        If Err.Number <> 0 Then
            MsgBox "Error adding value fields. Ensure field names are correct and exist in the Data Model.", vbExclamation
            Exit Sub
        End If
        
        ' Filter Fields (optional)
        On Error Resume Next
        .PivotFields("[StockInfo].[Sector]").Orientation = xlPageField
        .PivotFields("[FinancialMetrics].[Year]").Orientation = xlPageField
        
        If Err.Number <> 0 Then
            MsgBox "Error adding filter fields. Ensure field names are correct and exist in the Data Model.", vbExclamation
            Exit Sub
        End If
        
    End With
    
    ' Format the PivotTable for better readability (optional)
    With pvtTable
        .RowAxisLayout xlTabularRow   ' Tabular layout for rows.
        .RepeatAllLabels xlRepeatLabels   ' Repeat row labels.
        .ShowTableStyleRowStripes = True   ' Add striped rows.
        .TableStyle2 = "PivotStyleMedium9"   ' Apply a medium pivot table style.
    End With
    
    ' Autofit columns for better visibility.
    wsPivot.Columns.AutoFit
    
    MsgBox "Comprehensive Pivot Table created successfully!", vbInformation

End Sub
