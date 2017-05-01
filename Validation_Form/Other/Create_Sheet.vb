
Sub Create_Sheet()

With ThisWorkbook
  .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Combined Registry Measures"
End With

End Sub
