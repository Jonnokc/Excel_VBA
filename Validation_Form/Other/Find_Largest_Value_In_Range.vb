Sub ShowMax()

  ' This code finds the largest value within column "A"

Dim TheMax As Double

TheMax = WorksheetFunction.MAX(Range("A:A"))
MsgBox(TheMax)

End Sub
