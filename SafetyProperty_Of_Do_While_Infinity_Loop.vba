Sub Do_While_Forever_Fixed()

Dim x As Integer
Dim y As String
Dim emergency_break As Long

x = 1
y = "Apple"
emergency_break = 0

Do While y = "Apple" And emergency_break <= 1000
x = x + 1
emergency_break = emergency_break + 1
Loop
End Sub