Attribute VB_Name = "ChromeBrowser"
Sub test544()

  Dim chromePath As String

  chromePath = """C:\Program Files\Google\Chrome\Application\chrome.exe"""

  Shell (chromePath & " -url http:google.ca")

End Sub
