Sub CreatePivotTableFromTrackedData()
    Dim wsSource As Worksheet
    Dim wsPivot As Worksheet
    Dim pvtCache As PivotCache
    Dim pvtTable As PivotTable
    Dim sourceTbl As ListObject
    
    ' Set reference to the source worksheet
    Set wsSource = ThisWorkbook.Sheets("TrackedDataVLOOKUP")
    
    ' Set reference to the source table
    Set sourceTbl = wsSource.ListObjects("TrackedDataVLOOKUPTable")
    
    ' Create a new worksheet for the pivot table
    Set wsPivot = ThisWorkbook.Sheets.Add
    wsPivot.Name = "StockPivotTable"
    
    ' Create PivotCache from the source data
    Set pvtCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=sourceTbl.Range)
    
    ' Create PivotTable
    Set pvtTable = pvtCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="StockPerformancePivot")
    
    ' Add fields to the PivotTable
    With pvtTable
        .PivotFields("Stock Symbol").Orientation = xlRowField
        .PivotFields("Date").Orientation = xlColumnField
        .AddDataField .PivotFields("Open Price"), "Avg Open Price", xlAverage
        .AddDataField .PivotFields("Close Price"), "Avg Close Price", xlAverage
    End With
    
    ' Format the PivotTable
    pvtTable.ShowTableStyleRowStripes = True
    pvtTable.TableStyle2 = "PivotStyleMedium9"
    
    MsgBox "Pivot table created successfully!", vbInformation
End Sub
