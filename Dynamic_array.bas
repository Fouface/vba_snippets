Attribute VB_Name = "Dynamic_array"
Option Explicit

Sub Dynamic_array()

'declaring array with no element.
'---------------------
Dim iNames() As String
'---------------------

'declaring variables to store counter _
'and elements from the range.
'----------------------
Dim iCount As Integer
Dim iElement As Integer
'----------------------

'get the last row number to decide the _
'number of elements for the array.
'------------------------------------
iCount = Range("A1").End(xlDown).Row
'------------------------------------

're-defining the elements for the array.
'-------------------
ReDim iNames(iCount)
'-------------------

'using a for loop to add elements in the array
'from the range starting from cell A1
'--------------------------------------------------
For iElement = 1 To iCount
    iNames(iElement - 1) = Cells(iElement, 1).Value
Next iElement
'--------------------------------------------------

'print all the elements from the
'to the immediate window
'--------------------
Debug.Print iNames(0)
Debug.Print iNames(1)
Debug.Print iNames(2)
'--------------------

End Sub
