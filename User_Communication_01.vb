Sub What_Kind_Of_Day()
Dim myInput As Range
Dim offHelp As VbMsgBoxResult

Set myInput = Application.InputBox("How is your day?", "Today", , , , , , 8)

If myInput.Value = "I am doing well today." Then
Call MsgBox("Great! I'm glad to hear that.")
ElseIf myInput.Value = "Today is a bad day." Then

offHelp = MsgBox("I am sorry, is there anything that I can do?", vbYesNo, "Empathy")
If offHelp = vbYes Then
Call MsgBox("Yes, please do my homework for me.")
Else
Call MsgBox("No, I have an exam.")
End If

Else
Call MsgBox("Fuck it!")
End If

End Sub