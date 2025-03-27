Sub CreatePivotTableFromDataModel()
    Dim pvtCache As PivotCache
    Dim pvtTable As PivotTable
    Dim ws As Worksheet
    
    ' Add a new worksheet for the PivotTable
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Market Analysis PivotTable"
    
    ' Create PivotCache from the Data Model
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:="ThisWorkbookDataModel")
    
    ' Create PivotTable
    Set pvtTable = pvtCache.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="MarketAnalysisPivot")
    
    ' Add fields to the PivotTable
    With pvtTable
        ' Rows
        .PivotFields("Market").Orientation = xlRowField
        .PivotFields("Stock Name").Orientation = xlRowField
        
        ' Columns
        .PivotFields("Market Cap").Orientation = xlColumnField
        
        ' Values
        .AddDataField .PivotFields("Stock Price"), "Average Stock Price", xlAverage
        .AddDataField .PivotFields("Volume"), "Total Volume", xlSum
        .AddDataField .PivotFields("PE Ratio"), "Average PE Ratio", xlAverage
        .AddDataField .PivotFields("Dividend Yield"), "Average Dividend Yield", xlAverage
        .AddDataField .PivotFields("Net Profit Margin %"), "Average Net Profit Margin", xlAverage
        
        ' Add a filter
        .PivotFields("Sector Growth %").Orientation = xlPageField
    End With
    
    ' Format the PivotTable
    With pvtTable
        .LayoutForm = xlTabular
        .RepeatAllLabels xlRepeatLabels
        .RowAxisLayout xlTabularRow
    End With
    
    MsgBox "PivotTable created successfully!", vbInformation
End Sub
