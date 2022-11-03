Attribute VB_Name = "text2speech"
Sub SayThisString()

Dim SayThis As String

SayThis = "I love Microsoft Excel. is week. i am so based."

Application.Speech.Speak (SayThis)

End Sub
