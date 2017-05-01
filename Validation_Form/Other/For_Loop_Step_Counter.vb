Sub ShadeEveryThirdRow()

' Loops through a range and shades the background of every 3rd row grey.

Dim i As Long

For i = 1 To 100 Step 3
  Rows(i).Interior.Color = RGB(200, 200, 200)
  Next i

End Sub
