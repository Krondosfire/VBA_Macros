Sub GenerateStateAbbreviationsTable()
    Dim ws As Worksheet
    Dim states As Variant
    Dim i As Integer
    
    ' Create a new worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "US States and Abbreviations"
    
    ' Define the states and abbreviations array
    states = Array( _
        Array("Alabama", "AL"), Array("Alaska", "AK"), Array("Arizona", "AZ"), _
        Array("Arkansas", "AR"), Array("California", "CA"), Array("Colorado", "CO"), _
        Array("Connecticut", "CT"), Array("Delaware", "DE"), Array("Florida", "FL"), _
        Array("Georgia", "GA"), Array("Hawaii", "HI"), Array("Idaho", "ID"), _
        Array("Illinois", "IL"), Array("Indiana", "IN"), Array("Iowa", "IA"), _
        Array("Kansas", "KS"), Array("Kentucky", "KY"), Array("Louisiana", "LA"), _
        Array("Maine", "ME"), Array("Maryland", "MD"), Array("Massachusetts", "MA"), _
        Array("Michigan", "MI"), Array("Minnesota", "MN"), Array("Mississippi", "MS"), _
        Array("Missouri", "MO"), Array("Montana", "MT"), Array("Nebraska", "NE"), _
        Array("Nevada", "NV"), Array("New Hampshire", "NH"), Array("New Jersey", "NJ"), _
        Array("New Mexico", "NM"), Array("New York", "NY"), Array("North Carolina", "NC"), _
        Array("North Dakota", "ND"), Array("Ohio", "OH"), Array("Oklahoma", "OK"), _
        Array("Oregon", "OR"), Array("Pennsylvania", "PA"), Array("Rhode Island", "RI"), _
        Array("South Carolina", "SC"), Array("South Dakota", "SD"), Array("Tennessee", "TN"), _
        Array("Texas", "TX"), Array("Utah", "UT"), Array("Vermont", "VT"), _
        Array("Virginia", "VA"), Array("Washington", "WA"), Array("West Virginia", "WV"), _
        Array("Wisconsin", "WI"), Array("Wyoming", "WY") _
    )
    
    ' Add headers
    ws.Cells(1, 1).Value = "State"
    ws.Cells(1, 2).Value = "Abbreviation"
    
    ' Populate the table
    For i = 0 To UBound(states)
        ws.Cells(i + 2, 1).Value = states(i)(0)
        ws.Cells(i + 2, 2).Value = states(i)(1)
    Next i
    
    ' Format the table
    With ws.Range("A1:B" & UBound(states) + 2)
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Arial"
        .Font.Size = 11
        .Columns.AutoFit
    End With
    
    ' Format the header row
    With ws.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    MsgBox "Table generated successfully!", vbInformation
End Sub
