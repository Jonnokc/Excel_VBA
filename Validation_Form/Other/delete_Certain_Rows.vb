LastRow = sht.Cells(sht.Rows.Count, "C").End(xlUp).Row


' Loops through all the rows and deletes ones with a count too low
For i = LastRow To 1 Step -1
  If (Cells(i, Count_Col_Loc).Value) < 1000 Then
    Cells(i, "A").EntireRow.Delete
  End If
Next i
