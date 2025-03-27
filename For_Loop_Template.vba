Sub For_Loop_Lesson()
Dim x As Integer

For x = 1 To 10
      ' Do code 10 times
    x = x + 1
Next x

For x = 2 To 10 Step 2
      ' Do code 5 times
    x = x + 2
Next x

For x = 10 To 2 Step -2
      ' Do code 5 times
   x = x - 2
Next x
Dim whatreplArray(6, 2) As String
Dim i As Integer
Dim myWrkSheet As Worksheet

i = 1 ' i =1 after

'Define what column of Array
whatreplArray(1, 1) = "Aple"
whatreplArray(2, 1) = "Bannana"
whatreplArray(3, 1) = "Charry"
whatreplArray(4, 1) = "Chirry"
whatreplArray(5, 1) = "Grpe"
whatreplArray(6, 1) = "Oranje"

' Define replace column of Array
whatreplArray(1, 2) = "Apple"
whatreplArray(2, 2) = "Banana"
whatreplArray(3, 2) = "Cherry"
whatreplArray(4, 2) = "Cherry"
whatreplArray(5, 2) = "Grape"
whatreplArray(6, 2) = "Orange"

For Each myWrkSheet In Sheets
myWrkSheet.Activate
For i = LBound(whatreplArray, 1) To UBound(whatreplArray, 1)
With Cells
      .Replace What:=whatreplArray(i, 1), Replacement:=whatreplArray(i, 2)
End With

Next i

Next myWrkSheet
End Sub