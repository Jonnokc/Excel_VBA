Sub Code_System_Sheets()

    For Each code In Range("Code_ID_List")

			With ThisWorkbook
			  .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = code
			End With

    Next code

End Sub
