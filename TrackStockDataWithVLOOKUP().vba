Sub TrackStockDataWithVLOOKUP()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim dailyPricesTbl As ListObject
    Dim stockInfoTbl As ListObject
    Dim lastRow As Long
    Dim i As Long
    
    ' Set reference to the source worksheet containing tables
    Set wsSource = ThisWorkbook.Sheets("StockMarketData")
    
    ' Set references to the DailyPrices and StockInfo tables
    On Error Resume Next
    Set dailyPricesTbl = wsSource.ListObjects("DailyPrices")
    Set stockInfoTbl = wsSource.ListObjects("StockInfo")
    On Error GoTo 0
    
    If dailyPricesTbl Is Nothing Or stockInfoTbl Is Nothing Then
        MsgBox "DailyPrices or StockInfo table not found on StockMarketData sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Create a new worksheet for the tracked data
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("TrackedDataVLOOKUP")
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Sheets.Add
        wsTarget.Name = "TrackedDataVLOOKUP"
    Else
        wsTarget.Cells.Clear ' Clear existing data if sheet already exists
    End If
    On Error GoTo 0
    
    ' Set up headers in the new sheet
    With wsTarget
        .Cells(1, 1).Value = "Stock ID"
        .Cells(1, 2).Value = "Stock Symbol"
        .Cells(1, 3).Value = "Date"
        .Cells(1, 4).Value = "Open Price"
        .Cells(1, 5).Value = "Close Price"
    End With
    
    ' Loop through DailyPrices table and extract relevant data using VLOOKUP
    lastRow = dailyPricesTbl.ListRows.Count
    For i = 1 To lastRow
        With wsTarget
            .Cells(i + 1, 1).Value = dailyPricesTbl.DataBodyRange(i, 2).Value ' Stock ID
            
            ' Use VLOOKUP to find Stock Symbol from StockInfo table based on Stock ID
            .Cells(i + 1, 2).Formula = "=VLOOKUP(A" & (i + 1) & ",StockInfo[#All],2,FALSE)"
            
            ' Extract Date, Open Price, and Close Price from DailyPrices table
            .Cells(i + 1, 3).Value = dailyPricesTbl.DataBodyRange(i, 3).Value ' Date
            .Cells(i + 1, 4).Value = dailyPricesTbl.DataBodyRange(i, 4).Value ' Open Price
            .Cells(i + 1, 5).Value = dailyPricesTbl.DataBodyRange(i, 5).Value ' Close Price
        End With
    Next i
    
    ' Format as table for better readability
    wsTarget.ListObjects.Add(xlSrcRange, wsTarget.Range("A1").CurrentRegion, , xlYes).Name = "TrackedDataVLOOKUPTable"
    wsTarget.Columns.AutoFit
    
    MsgBox "Stock data has been successfully tracked and created in the TrackedDataVLOOKUP sheet!", vbInformation
End Sub
