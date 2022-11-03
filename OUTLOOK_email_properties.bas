Attribute VB_Name = "OUTLOOK_email_properties"
Sub vdsv()
Dim wb As Object
Set wb = ThisWorkbook

 For Each Property In ThisWorkbook.BuiltinDocumentProperties
           Debug.Print vbTab & Property.Name & " = " & Property.Value
           Next Property
End Sub

Sub EnumerateItemProperties()
 
 Dim oM As Outlook.MailItem
 
 Dim i As Integer
 
 Set oM = Application.ActiveInspector.CurrentItem
 
 If Not (oM Is Nothing) Then
 
 For i = 0 To oM.ItemProperties.Count - 1
 
 Debug.Print oM.ItemProperties(i).Name
 
 Next
 
 End If
 
End Sub
