Sub TestTwo()

Dim myVariableName As Integer

Dim x As Integer

Dim y(5) As Integer 'Y is an array --> Y(0) = First Entry
                                     ' Y(1) = Second Entry
                                     ' Y(2) = Third Entry
                                     ' Y(3) = Fourth Entry
                                     ' Y(4) = Fifth Entry
                                     ' Y(5) = Sixth Entry


y(3) = 5
y(1) = 12
y(5) = 18

Dim W(5, 3) As Integer 'W(5,3) will have 15 slots:
                            ' W(1,1) W(1,2) W(1,3)
                            ' W(2,1) W(2,2) W(2,3)
                            ' W(3,1) W(3,2) W(3,3)
                            ' W(4,1) W(4,2) W(4,3)
                            ' W(4,1) W(4,2) W(4,3)

W(3, 2) = 4
W(1, 3) = 17
W(5, 3) = 38

Dim variableName As String


variableName = "Apple"

Dim myWorksheet As Worksheet

Dim tempApplication As Application

x = 1

Set myWorksheet = Sheets(1)
Set tempApplication = New Application


myVariableName = 1
Dim myTestName As String

myTestName = (myVariableName = 1)

myVariableName = myVariableName + 1




End Sub