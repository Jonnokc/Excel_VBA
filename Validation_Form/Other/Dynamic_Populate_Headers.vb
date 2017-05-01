Sub Dynamic_Header_Creator()
'
' This program will loop through an array and populate each column header
' by the coresponding variable within the array.
'

Dim Header_Array As Variant

Header_Array = Array("ColA", "ColB", "ColC")

Off_Count = 0
For i = 0 To UBound(Header_Array)
    ' Uses cell "A1" as the starting point for the header row.
    Range("A1").Offset(0, Off_Count).Value = Header_Array(i) ' places next array value within the next column
    Off_Count = Off_Count + 1 'Increases the offset count on each loop
Next i

End Sub
