Sub GenerateStateAbbreviationsAndZipCodesTable()
    Dim ws As Worksheet
    Dim states As Variant
    Dim i As Integer
    
    ' Create a new worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "US States, Abbreviations, and ZIP Codes"
    
    ' Define the states, abbreviations, and ZIP code ranges array
    states = Array( _
        Array("Alabama", "AL", "35004-36925"), Array("Alaska", "AK", "99501-99950"), _
        Array("Arizona", "AZ", "85001-86556"), Array("Arkansas", "AR", "71601-72959"), _
        Array("California", "CA", "90001-96162"), Array("Colorado", "CO", "80001-81658"), _
        Array("Connecticut", "CT", "06001-06928"), Array("Delaware", "DE", "19701-19980"), _
        Array("Florida", "FL", "32004-34997"), Array("Georgia", "GA", "30001-39901"), _
        Array("Hawaii", "HI", "96701-96898"), Array("Idaho", "ID", "83201-83876"), _
        Array("Illinois", "IL", "60001-62999"), Array("Indiana", "IN", "46001-47997"), _
        Array("Iowa", "IA", "50001-52809"), Array("Kansas", "KS", "66002-67954"), _
        Array("Kentucky", "KY", "40003-42788"), Array("Louisiana", "LA", "70001-71497"), _
        Array("Maine", "ME", "03901-04992"), Array("Maryland", "MD", "20331-21930"), _
        Array("Massachusetts", "MA", "01001-05544"), Array("Michigan", "MI", "48001-49971"), _
        Array("Minnesota", "MN", "55001-56763"), Array("Mississippi", "MS", "38601-39776"), _
        Array("Missouri", "MO", "63001-65899"), Array("Montana", "MT", "59001-59937"), _
        Array("Nebraska", "NE", "68001-69367"), Array("Nevada", "NV", "88901-89883"), _
        Array("New Hampshire", "NH", "03031-03897"), Array("New Jersey", "NJ", "07001-08989"), _
        Array("New Mexico", "NM", "87001-88439"), Array("New York", "NY", "00501-14975"), _
        Array("North Carolina", "NC", "27006-28909"), Array("North Dakota", "ND", "58001-58856"), _
        Array("Ohio", "OH", "43001-45999"), Array("Oklahoma", "OK", "73001-74966"), _
        Array("Oregon", "OR", "97001-97920"), Array("Pennsylvania", "PA", "15001-19640"), _
        Array("Rhode Island", "RI", "02801-02940"), Array("South Carolina", "SC", "29001-29945"), _
        Array("South Dakota", "SD", "57001-57799"), Array("Tennessee", "TN", "37010-38589"), _
        Array("Texas", "TX", "73301-79999"), Array("Utah", "UT", "84001-84784"), _
        Array("Vermont", "VT", "05001-05907"), Array("Virginia", "VA", "20101-24658"), _
        Array("Washington", "WA", "98001-99403"), Array("West Virginia", "WV", "24701-26886"), _
        Array("Wisconsin", "WI", "53001-54990"), Array("Wyoming", "WY", "82001-83414") _
    )
    
    ' Add headers
    ws.Cells(1, 1).Value = "State"
    ws.Cells(1, 2).Value = "Abbreviation"
    ws.Cells(1, 3).Value = "ZIP Code Range"
    
    ' Populate the table
    For i = 0 To UBound(states)
        ws.Cells(i + 2, 1).Value = states(i)(0)
        ws.Cells(i + 2, 2).Value = states(i)(1)
        ws.Cells(i + 2, 3).Value = states(i)(2)
    Next i
    
    ' Format the table
    With ws.Range("A1:C" & UBound(states) + 2)
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Arial"
        .Font.Size = 11
        .Columns.AutoFit
    End With
    
    ' Format the header row
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    MsgBox "Table generated successfully!", vbInformation
End Sub
