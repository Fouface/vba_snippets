Attribute VB_Name = "z_join"
Sub JoinArray_Example1()
    Dim strFullPath As String
    Dim arrPath(0 To 2) As String
    arrPath(0) = "C:"
    arrPath(1) = "Users"
    arrPath(2) = "Public"
    strFullPath = Join(arrPath, "\")
    Debug.Print strFullPath
End Sub

Sub JoinArray_Example2()
    MsgBox Join(Array("A", "B", "C", "C", "C"), ", ")
End Sub




Sub UnionExample()
'create a union of a range of cells and select group
Dim Rng1, Rng2, Rng3 As Range

Set Rng1 = Range("A1,A3,A5,A7,A9,A11,A13,A15,A17,A19,A21")
Set Rng2 = Range("C1,C3,C5,C7,C9,C11,C13,C15,C17,C19,C21")
Set Rng3 = Range("E1,E3,E5,E7,E9,E11,E13,E15,E17,E19,E21")

Union(Rng1, Rng2, Rng3).Select

End Sub
