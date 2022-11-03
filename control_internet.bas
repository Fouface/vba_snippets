Attribute VB_Name = "control_internet"
Option Explicit

Sub internetcontrols()
Dim ie As Object, url As String
Set ie = CreateObject("internetexplorer.application")

url = "https://www.google.com"

ie.navigate url




ie.Quit
ie.Close

End Sub
