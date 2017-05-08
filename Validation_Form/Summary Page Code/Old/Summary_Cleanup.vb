Private Sub Summary_Cleanup()

' Delets extra sheets and columns after summary sheet process has completed


Dim sheet as Worksheet
Dim header_location As Variant

Sheets("Summary View").Select

Range("A1:K1").Name = "Header_Row"

For each header in Range("Header_Row")
  ' Deletes the Concat column
  If header = "Concat" Then
    header_location = Mid(header.Address, 2, 1)
    Columns(header_location & ":" & header_location).Select
    Selection.Delete Shift:=xlToLeft
  End If
Next header

' Delete the extra sheets
Application.DisplayAlerts = False

For Each sheet In Worksheets

    If sheet.Name = "Potential_Summary_Pivot" _
       Or sheet.Name = "Clinical_Summary_Pivot" _
       Or sheet.Name = "Unmapped_Summary_Pivot" _
       Or sheet.Name = "Combined Registry Measures" _
       Then
        sheet.Delete
    End If
Next sheet

    Application.DisplayAlerts = True

End Sub
