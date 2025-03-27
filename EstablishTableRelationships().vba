Sub EstablishTableRelationships()
    Dim ws As Worksheet
    Dim mdl As Model
    Dim tblStockInfo As ListObject
    Dim tblDailyPrices As ListObject
    Dim tblFinancialMetrics As ListObject
    
    ' Set reference to the worksheet containing the tables
    Set ws = ThisWorkbook.Sheets("StockMarketData")
    
    ' Set reference to the workbook's data model
    Set mdl = ThisWorkbook.Model
    
    ' Set references to the tables
    Set tblStockInfo = ws.ListObjects("StockInfo")
    Set tblDailyPrices = ws.ListObjects("DailyPrices")
    Set tblFinancialMetrics = ws.ListObjects("FinancialMetrics")
    
    ' Add tables to the data model if they're not already there
    On Error Resume Next
    mdl.AddTable tblStockInfo
    mdl.AddTable tblDailyPrices
    mdl.AddTable tblFinancialMetrics
    On Error GoTo 0
    
    ' Create relationships
    On Error Resume Next
    
    ' Relationship between StockInfo and DailyPrices
    mdl.AddRelationship _
        tblStockInfo.ListColumns("ID").DataBodyRange, _
        tblDailyPrices.ListColumns("StockID").DataBodyRange
    
    ' Relationship between StockInfo and FinancialMetrics
    mdl.AddRelationship _
        tblStockInfo.ListColumns("ID").DataBodyRange, _
        tblFinancialMetrics.ListColumns("StockID").DataBodyRange
    
    If Err.Number <> 0 Then
        MsgBox "Error creating relationships. They may already exist or there might be an issue with the data.", vbExclamation
    Else
        MsgBox "Relationships established successfully!", vbInformation
    End If
    
    On Error GoTo 0
End Sub
