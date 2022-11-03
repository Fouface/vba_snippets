Attribute VB_Name = "z_array"
Sub testrange()
Dim rng, tt
rng = Range("a1:a5")

a = RangeToArray(Range("a1:a5"))

Debug.Print Join(a, " ") ' print complete iarray as a join
End Sub
Public Function RangeToArray(rng As Range) As Variant
    Dim i As Long, r As Range
    ReDim arr(1 To rng.Count)

    i = 1
    For Each r In rng
        arr(i) = r.Value
        i = i + 1
     Next r

    RangeToArray = arr
End Function








Sub WriteArray()
    ' Read the data to an array
    Data = Range("A1:E10").Value
    ' Resize the output range and write the array to another sheet
    Worksheets("sheet7").Range("A1").Resize(UBound(Data, 1), UBound(Data, 2)).Value = Data
End Sub







Public Sub TestLoopArray2()
'declare the array
   Dim rnArray() As Variant
'Declare the integer to store the number of rows
   Dim iRw As Integer
'Assign range to a the array variable
   rnArray = Range("A1:A10")
'loop through the rows - 1 to 10
   For iRw = 1 To UBound(rnArray)
'output to the immediate window
     Debug.Print rnArray(iRw, 1)
   Next iRw
End Sub



Public Sub TestOutputTranspose()
'declare the array
   Dim rnArray() As Variant
'populate it with the range
   rnArray = Range("A1:A38")
'transpose the data
   Range(Cells(1, 3), Cells(1, 40)).Value = Application.Transpose(rnArray)
    Range("C2:m2").Value = Application.Transpose(rnArray) 'does the same
    

End Sub


Public Sub TestLoopArray()
'declare the array
   Dim rnArray() As Variant
'Declare the integer to store the number of rows
   Dim iRw As Integer
'Assign range to a the array variable
   rnArray = Range("A1:A10")
'loop through the values in the array
   For iRw = LBound(rnArray) To UBound(rnArray)
'populate a different range with the data
      Cells(iRw, 9).Value = rnArray(iRw, 1)
   Next iRw
End Sub






Public Sub TestOutput()
'declare the array
   Dim rnArray() As Variant
'populate the array with the range
   rnArray = Range("A1:b5")
'output the array to a different range of cells
   Range("d1:e5") = rnArray()
   
   
   Dim i As Integer
   For i = 1 To 5
        Debug.Print rnArray(i, 1)
   Next i
   
   Dim ii As Variant
    For ii = LBound(rnArray) To UBound(rnArray)
        Debug.Print rnArray(ii, 1)
    Next ii
   
End Sub



Sub LoopForNextStatic()
'declare a variant array
   Dim strNames(1 To 4) As String
'populate the array
   strNames(1) = "Bob"
   strNames(2) = "Peter"
   strNames(3) = "Keith"
   strNames(4) = "Sam"
'declare an integer
   Dim i As Integer
'loop from position 2 to position 3 of the array
   For i = 2 To 3
'show the name in the immediate window
      Debug.Print strNames(i)
   Next i
End Sub




Sub LoopForNextDynamic()

'Dim strNames(1 To 4)
'declare a variant array
  ' Dim strNames() As String
'initialize the array
   ReDim strNames(1 To 5)
'populate the array
   strNames(1) = "Bob"
   strNames(2) = "Peter"
   strNames(3) = "Keith"
   strNames(4) = "Sam"
   strNames(4) = "Sasdam"
'declare an integer
   Dim i As Variant
'loop from the lower bound of the array to the upper bound of the array - the entire array
   For i = LBound(strNames) To UBound(strNames)
'show the name in the immediate window
      Debug.Print strNames(i)
   Next i
End Sub


Sub LoopForArrayStatic()
   'declare a variant array
   Dim strNames(1 To 4) As String

   'populate the array
   strNames(1) = "Bob"
   strNames(2) = "Peter"
   strNames(3) = "Keith"
   strNames(4) = "Sam"

   'declare a variant to hold the array element
   Dim item As Variant

   'loop through the entire array
   For Each item In strNames
      'show the element in the debug window.
      Debug.Print item
   Next item
End Sub
