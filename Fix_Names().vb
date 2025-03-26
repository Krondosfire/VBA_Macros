Sub Fix_Names()
'
' Fix_Names Macro
'

'
    Cells.Replace What:="Oranje", Replacement:="Orange", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
        Cells.Replace What:="Charry", Replacement:="Cherry", LookAt:=xlPart
        Dim myBoolean As Boolean
        
        myBoolean = Cells.Replace("Chirry", "Cherry")
        
        With Cells
            .Replace What:="Aple", Replacement:="Apple"
        End With
        Call Cells.Replace("Bannana", "Banana")
        Call Cells.Replace("Grpe", "Grape")
        
End Sub
