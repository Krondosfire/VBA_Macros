Sub CreateTableRelationship()
    Dim wb As Workbook
    Dim mdl As Model
    Dim rel As ModelRelationship
    
    Set wb = ThisWorkbook
    Set mdl = wb.Model
    
    ' Check if tables exist using helper function
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

' Helper function to check if a table exists in the data model
Function TableExists(mdl As Model, tblName As String) As Boolean
    Dim tbl As ModelTable
    TableExists = False
    For Each tbl In mdl.Tables
        If tbl.Name = tblName Then
            TableExists = True
            Exit Function
        End If
    Next tbl
End Function
