Attribute VB_Name = "arrays"
Option Explicit

Private Sub Constant_demo_Click()
   Dim arr(2, 3) As Variant ' Which has 3 rows and 4 columns
   arr(0, 0) = "Apple"
   arr(0, 1) = "Orange"
   arr(0, 2) = "Grapes"
   arr(0, 3) = "pineapple"
   arr(1, 0) = "cucumber"
   arr(1, 1) = "beans"
   arr(1, 2) = "carrot"
   arr(1, 3) = "tomato"
   arr(2, 0) = "potato"
   arr(2, 1) = "sandwitch"
   arr(2, 2) = "coffee"
   arr(2, 3) = "nuts"
           
   MsgBox ("Value in Array index 0,1 : " & arr(0, 1))
   MsgBox ("Value in Array index 2,2 : " & arr(2, 2))
End Sub



Private Sub AAConstant_demo_Click()
   Dim a() As Variant, i As Integer
   i = 0
   ReDim a(5)
   a(0) = "XYZ"
   a(1) = 41.25
   a(2) = 22
  
   ReDim Preserve a(7)
   For i = 3 To 7
   a(i) = i
   Next
  
   'to Fetch the output
   For i = 0 To UBound(a)
      Debug.Print a(i)
   Next
End Sub



































Sub TestArrayValuesMultiple()
'Declare the array as a variant array
   Dim arRng() As Variant

'Declare the integer to store the number of rows
   Dim iRw As Integer

'Declare the integer to store the number of columns
   Dim iCol As Integer

'Assign range to a the array variable
   arRng = Range("A1:c10")

'loop through the rows - 1 to 10
   For iRw = 1 To UBound(arRng, 1)

'now - while in row 1, loop through the 3 columns
      For iCol = 1 To UBound(arRng, 1)

'show the result in the immediate window
         Debug.Print arRng(iRw, iCol)
      Next iCol
   Next iRw
End Sub
