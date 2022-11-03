Attribute VB_Name = "Macros_2_rightCLIKmenu"
Public Sub mymacro1()

MsgBox "Macro1 from a right click menu"

End Sub



Public Sub mymacro2()

MsgBox "Macro2 from a right click menu"

End Sub


Private Sub Workbook_Open()

Dim MyMenu As Object
    
Set MyMenu = Application.ShortcutMenus(xlWorksheetCell) _
    .MenuItems.AddMenu("This is my Custom Menu", 1)
    
With MyMenu.MenuItems
    .Add "MyMacro1", "MyMacro1", , 1, , ""
    .Add "MyMacro2", "MyMacro2", , 2, , ""
End With
  
Set MyMenu = Nothing

End Sub

