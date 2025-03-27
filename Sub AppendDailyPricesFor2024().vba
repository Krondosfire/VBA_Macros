Sub AppendDailyPricesFor2024()
    Dim ws As Worksheet
    Dim tblDailyPrices As ListObject
    Dim lastRow As Long
    Dim newDataRange As Range
    Dim i As Long, j As Long
    Dim startDate As Date
    Dim stockCount As Long
    
    ' Set reference to the worksheet containing the DailyPrices table
    Set ws = ThisWorkbook.Sheets("StockMarketData")
    
    ' Set reference to the DailyPrices table
    Set tblDailyPrices = ws.ListObjects("DailyPrices")
    
    ' Get the last row of the table
    lastRow = tblDailyPrices.Range.Rows.Count
    
    ' Set the start date for 2024
    startDate = DateSerial(2024, 1, 1)
    
    ' Determine the number of unique stocks
    stockCount = Application.WorksheetFunction.Max(tblDailyPrices.ListColumns("StockID").DataBodyRange)
    
    ' Set the range for new data (assuming 252 trading days per year)
    Set newDataRange = tblDailyPrices.Range.Resize(lastRow + (252 * stockCount))
    
    ' Resize the table to include new rows
    tblDailyPrices.Resize newDataRange
    
    ' Generate and append new data for 2024
    For i = 1 To stockCount
        For j = 0 To 251 ' 252 trading days
            lastRow = lastRow + 1
            With tblDailyPrices.DataBodyRange
                .Cells(lastRow, 1).Value = lastRow ' ID
                .Cells(lastRow, 2).Value = i ' StockID
                .Cells(lastRow, 3).Value = WorksheetFunction.WorkDay(startDate, j) ' Date
                .Cells(lastRow, 4).Value = Round(Rnd() * 100 + 50, 2) ' OpenPrice
                .Cells(lastRow, 5).Value = Round(.Cells(lastRow, 4).Value * (1 + (Rnd() - 0.5) / 10), 2) ' ClosePrice
            End With
        Next j
    Next i
    
    ' Sort the table by StockID and Date
    tblDailyPrices.Sort.SortFields.Clear
    tblDailyPrices.Sort.SortFields.Add Key:=tblDailyPrices.ListColumns("StockID").Range, SortOn:=xlSortOnValues, Order:=xlAscending
    tblDailyPrices.Sort.SortFields.Add Key:=tblDailyPrices.ListColumns("Date").Range, SortOn:=xlSortOnValues, Order:=xlAscending
    tblDailyPrices.Sort.Apply
    
    MsgBox "Daily prices for 2024 have been appended successfully!", vbInformation
End Sub
