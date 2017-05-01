Sub TurnAutoFilterOn()
'check for filter, turn on if none exists
  If Not ActiveSheet.AutoFilterMode Then
    ActiveSheet.Range("A2").AutoFilter
  End If
End Sub
