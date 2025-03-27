Sub TrackStockDataByDate()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim dailyPricesTbl As ListObject
    Dim stockInfoTbl As ListObject
    Dim trackedDate As Date
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
    
    ' Prompt user to enter a date to track data
    trackedDate = Application.InputBox("Enter the date (YYYY-MM-DD) to track stock data:", "Track Stock Data", Date, , , , , 1)
    
    If trackedDate = False Then Exit Sub ' User canceled
    
    ' Create a new worksheet for the tracked data
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("TrackedData")
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Sheets.Add
        wsTarget.Name = "TrackedData"
    Else
        wsTarget.Cells.Clear ' Clear existing data if sheet already exists
    End If
    On Error GoTo 0
    
    ' Set up headers in the new sheet
    With wsTarget
        .Cells(1, 1).Value = "Date"
        .Cells(1, 2).Value = "Stock ID"
        .Cells(1, 3).Value = "Stock Symbol"
        .Cells(1, 4).Value = "Open Price"
        .Cells(1, 5).Value = "Close Price"
        .Cells(2, 1).Resize(dailyPricesTbl.ListRows.Count).Value = trackedDate ' Fill the date column with the tracked date
    End With
    
    ' Loop through DailyPrices table and extract relevant data using VLOOKUP
    For i = 1 To dailyPricesTbl.ListRows.Count
        With wsTarget
            .Cells(i + 1, 2).Value = dailyPricesTbl.DataBodyRange(i, 2).Value ' Stock ID
            
            ' Use VLOOKUP to find Stock Symbol from StockInfo table based on Stock ID
            .Cells(i + 1, 3).Formula = "=VLOOKUP(B" & (i + 1) & ",StockInfo[#All],2,FALSE)"
            
            ' Use VLOOKUP to find Open Price and Close Price from DailyPrices table based on Date and Stock ID
            If dailyPricesTbl.DataBodyRange(i, 3).Value = trackedDate Then
                .Cells(i + 1, 4).Value = dailyPricesTbl.DataBodyRange(i, 4).Value ' Open Price
                .Cells(i + 1, 5).Value = dailyPricesTbl.DataBodyRange(i, 5).Value ' Close Price
            End If
        End With
    Next i
    
    ' Format as table for better readability (Optional)
    wsTarget.ListObjects.Add(xlSrcRange, wsTarget.Range("A1").CurrentRegion, , xlYes).Name = "TrackedDataTable"
    
    MsgBox "Tracked data for " & trackedDate & " has been successfully created in the TrackedData sheet!", vbInformation
End Sub
