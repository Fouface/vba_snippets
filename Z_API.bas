Attribute VB_Name = "Z_API"
'HttpRequest SetCredentials flags.

Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0


Private Sub ListSubs()
'requires microsoft winHTTP service reference
Dim MyRequest As New WinHttpRequest


    MyRequest.Open "GET", _
    "https://www.automateexcel.com/vba/winhttprequest-with-login/"


    'Set credentials
'only needed if api key is required
    'MyRequest.SetCredentials "USERNAME", "PASSWORD", _
    'HTTPREQUEST_SETCREDENTIALS_FOR_SERVER


    ' Send Request.
    MyRequest.send

    'And we get this response
    MsgBox MyRequest.responseText

End Sub

