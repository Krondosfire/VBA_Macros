Sub Boolean_Examples()

Dim my_name_is_correct As Boolean
Dim first_name As String
Dim last_name As String

Dim dog_is_correct As Boolean
Dim dog_name As String

Dim combined_is_correct As Boolean

my_name_is_correct = (first_name = "Javor") And (last_name = "Mladenoff")

dog_is_correct = (dog_name = "Sharo") Or (dog_name = "Gruh")

combined_is_correct = ((first_name = "Javor") And (last_name = "Mladenoff")) And ((dog_name = "Sharo") Or (dog_name = "Gruh"))

combined_is_correct = my_name_is_correct And dog_is_correct

If first_name = "Javor" And Len(first_name) <> 12 Or dog_name = "Kvik" Or dog_name = "Salam" Then
ElseIf first_name = "Pepa" Then
ElseIf first_name = "Yoan" Then
ElseIf first_name = "Biserka" Then
Else
End If

Select Case first_name

Case "Javor"
Case "Pepa"
Case "Yoan"
Case "Biserka"
Case Else
End Select

Select Case Len(first_name)
Case 10, 9, 8, 7 ' this is equal to: If Len(first_name) = 10 OR Len(first_name) = 9 OR Len(first_name) = 8 OR Len(first_name) = 7 Then
Case 11
Case Is >= 12
Case Else
End Select




End Sub
