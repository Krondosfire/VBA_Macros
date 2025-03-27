Sub RefreshAggregatedData()
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets("AggregatedData")
    Set tbl = ws.ListObjects("AggregatedData")
    
    ' Your custom code to update the table data goes here
    ' For example:
    ' tbl.DataBodyRange.Formula = "=SomeCalculation()"
    
    ' Alternatively, loop through the table and update cell values
    ' For Each cell In tbl.DataBodyRange
    '     cell.Value = SomeFunction(cell)
    ' Next cell
    
    MsgBox "AggregatedData table updated successfully!", vbInformation
End Sub
