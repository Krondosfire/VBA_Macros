Sub ExtractDataWithVLOOKUP()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set reference to source worksheet
    Set wsSource = ThisWorkbook.Sheets("StockMarketData")
    
    ' Create new target worksheet
    Set wsTarget = ThisWorkbook.Sheets.Add
    wsTarget.Name = "VLOOKUPData"
    
    ' Set up headers in the new sheet
    With wsTarget
        .Cells(1, 1).Value = "Stock ID"
        .Cells(1, 2).Value = "Stock Symbol"
        .Cells(1, 3).Value = "Company Name"
        .Cells(1, 4).Value = "Latest Close Price"
        .Cells(1, 5).Value = "Latest Revenue"
    End With
    
    ' Find last row in StockInfo table
    lastRow = wsSource.ListObjects("StockInfo").Range.Rows.Count
    
    ' Loop through StockInfo and extract data
    For i = 2 To lastRow
        With wsTarget
            .Cells(i, 1).Value = wsSource.Cells(i + 1, 1).Value ' Stock ID
            
            ' Use VLOOKUP for Stock Symbol and Company Name
            .Cells(i, 2).Formula = "=VLOOKUP(A" & i & ",StockInfo[#All],2,FALSE)"
            .Cells(i, 3).Formula = "=VLOOKUP(A" & i & ",StockInfo[#All],3,FALSE)"
            
            ' Use VLOOKUP for Latest Close Price
            .Cells(i, 4).Formula = "=VLOOKUP(A" & i & ",DailyPrices[#All],5,FALSE)"
            
            ' Use VLOOKUP for Latest Revenue
            .Cells(i, 5).Formula = "=VLOOKUP(A" & i & ",FinancialMetrics[#All],4,FALSE)"
        End With
    Next i
    
    ' Format as table
    wsTarget.ListObjects.Add(xlSrcRange, wsTarget.Range("A1").CurrentRegion, , xlYes).Name = "VLOOKUPData"
    wsTarget.ListObjects("VLOOKUPData").TableStyle = "TableStyleMedium6"
    
    ' Autofit columns
    wsTarget.Columns.AutoFit
    
    MsgBox "Data extraction with VLOOKUP complete!", vbInformation
End Sub
