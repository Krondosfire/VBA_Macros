Sub ExtractDataWithVLOOKUP()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long, i As Long
    
    ' Set references to worksheets
    Set ws1 = ThisWorkbook.Sheets("Sheet1") ' Sheet with lookup values
    Set ws2 = ThisWorkbook.Sheets("Sheet2") ' Sheet with data table
    
    ' Find last row in Sheet1
    lastRow = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through rows and perform VLOOKUP
    For i = 2 To lastRow ' Assuming headers in row 1
        ws1.Cells(i, "B").Value = Application.WorksheetFunction.VLookup( _
            ws1.Cells(i, "A").Value, _
            ws2.Range("A2:C100"), _
            3, _
            False)
    Next i
    
    MsgBox "Data extraction complete!", vbInformation
End Sub
