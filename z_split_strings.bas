Attribute VB_Name = "z_split_strings"
Sub SplitText()
 Dim strT As String
 Dim strArray() As String
 Dim Name As Variant

'populate the string with names
 strT = "John,Mary,Jack,Fred,Melanie,Steven,Paul,Robert"

'populate the array and indicate the delmiter
 strArray = Split(strT, ",")

'loop through each name and display in immediate window
 For Each Name In strArray
   Debug.Print Name
 Next
End Sub


Sub ExtractData()
 Dim StrData As String
 Dim strLeft As String
 Dim strRight As String
 Dim strMid As String

'populate the string
 StrData = "C:\Data\TestFile.xls"

'break down the name
 strLeft = Left(StrData, 7)
 strMid = Mid(StrData, 9, 8)
 strRight = Right(StrData, 3)

'return the results
 MsgBox "The path is " & strLeft & ", the File name is " & strMid & " and the extension is " & strRight
End Sub



Sub ExtractData2()
  Dim StrData As String
  StrData = "John""Mary""Jack""Fred""Melanie""Steven""Paul""Robert"""
  StrData = Replace(StrData, """", ",")
  MsgBox StrData
End Sub


