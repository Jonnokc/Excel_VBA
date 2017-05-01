Sub CreateIndexSheet()

    Dim Sheet As Worksheet


    ActiveWorkbook.Sheets.Add(Before:=Worksheets(1)).Name = "Index Sheet"    'Call whatever you like

    Range("A1").Select
    Selection.Value = "Index Sheet"
    ActiveCell.Offset(1, 0).Select    'Moves down a row

    For Each Sheet In Worksheets
        If Sheet.Name <> "Index Sheet" Then
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'" & Sheet.Name & "'" & "!A1", TextToDisplay:=Sheet.Name
            ActiveCell.Offset(1, 0).Select    'Moves down a row
        End If
    Next Sheet


    Range("A1").EntireColumn.AutoFit
    Range("A1").EntireRow.Delete    'Remove content Sheet from content list

End Sub
