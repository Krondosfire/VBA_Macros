Sub CreatePivotTableWithComplexCalculations()
    ' First, create the DAX measures
    Call CreateDAXMeasures
    
    Dim pvtCache As PivotCache
    Dim pvtTable As PivotTable
    Dim ws As Worksheet
    
    ' Add a new worksheet for the PivotTable
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Complex Market Analysis"
    
    ' Create PivotCache from the Data Model
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:="ThisWorkbookDataModel")
    
    ' Create PivotTable
    Set pvtTable = pvtCache.CreatePivotTable(TableDestination:=ws.Range("A3"), TableName:="ComplexMarketAnalysisPivot")
    
    ' Add fields to the PivotTable
    With pvtTable
        ' Rows
        .PivotFields("Stock Name").Orientation = xlRowField
        .PivotFields("Market").Orientation = xlRowField
        
        ' Columns
        .PivotFields("Sector Growth %").Orientation = xlColumnField
        
        ' Values (including our new measures)
        .AddDataField .PivotFields("Stock Price"), "Average Stock Price", xlAverage
        .AddDataField .PivotFields("PEG Ratio"), "Average PEG Ratio", xlAverage
        .AddDataField .PivotFields("ROI"), "Average ROI", xlAverage
        .AddDataField .PivotFields("Debt to EBITDA"), "Average Debt to EBITDA", xlAverage
        .AddDataField .PivotFields("YoY Growth"), "Average YoY Growth", xlAverage
        
        ' Add filters
        .PivotFields("Industry Growth %").Orientation = xlPageField
        .PivotFields("Market Cap").Orientation = xlPageField
    End With
    
    ' Format the PivotTable
    With pvtTable
        .LayoutForm = xlTabular
        .RepeatAllLabels xlRepeatLabels
        .RowAxisLayout xlTabularRow
    End With
    
    MsgBox "PivotTable with complex calculations created successfully!", vbInformation
End Sub
