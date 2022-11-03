Attribute VB_Name = "IE_pull_html_data"
Sub Pull_Data_from_Website()

'Website_Address = "https://exceldemy.com"
Website_Address = "https://www.automateexcel.com/vba/charts-graphs/#Creating_an_Embedded_Chart_Using_VBA"
HTML_Tag = "div"

Dim Browser As New InternetExplorer
Dim Doc As New HTMLDocument
Dim Data As Object

Browser.Visible = True
Browser.navigate Website_Address

Do
DoEvents
Loop Until Browser.readyState = READYSTATE_COMPLETE
Set Doc = Browser.document
Set Data = Doc.getElementsByTagName(HTML_Tag)

MsgBox Data(0).innerHTML
Cells(1, 1) = Data(0).innerHTML

End Sub






Function PullData(Website_Address, HTML_Tag)
'the same as above but in a function format
Dim Browser As New InternetExplorer
Dim Doc As New HTMLDocument
Dim Data As Object

Browser.navigate Website_Address

Do
DoEvents
Loop Until Browser.readyState = READYSTATE_COMPLETE
Set Doc = Browser.document
Set Data = Doc.getElementsByTagName(HTML_Tag)

Dim Output As Variant
ReDim Output(Data.Length - 1)

For i = LBound(Output) To UBound(Output)
    Output(i) = Data(i).innerHTML
Next i

PullData = Output

End Function
