Attribute VB_Name = "Formula_cell_display"
Option Explicit

Function Show_Cell_Formula(Cell As Range) As String
'this shows the formula in selected cell.
'use this in a spreadsheet
    Show_Cell_Formula = "Cell " & Cell.Address & " has the formulae: " & Cell.Formula & " '"
End Function


