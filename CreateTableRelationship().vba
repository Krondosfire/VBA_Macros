Sub CreateTableRelationship()
    Dim wb As Workbook
    Dim mdl As Model
    Dim rel As ModelRelationship
    Dim i As Long
    
    Set wb = ThisWorkbook
    Set mdl = wb.Model
    
    ' Check if tables exist using indexed loop
    If Not TableExists(mdl, "MarketData") Or Not TableExists(mdl, "FinanceData") Then
        MsgBox "Tables not found in data model. Add them first!", vbExclamation
        Exit Sub
    End If
    
    ' Create relationship
    Set rel = mdl.Relationships.Add( _
        ForeignKeyColumn:=mdl.Tables("FinanceData").Columns("ID"), _
        PrimaryKeyColumn:=mdl.Tables("MarketData").Columns("ID") _
    )
    
    MsgBox "Relationship created successfully!", vbInformation
End Sub

' Helper function using indexed loop instead of For Each
Function TableExists(mdl As Model, tblName As String) As Boolean
    TableExists = False
    For i = 1 To mdl.Tables.Count
        If mdl.Tables(i).Name = tblName Then
            TableExists = True
            Exit Function
        End If
    Next i
End Function
