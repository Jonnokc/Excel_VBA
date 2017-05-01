Sub WorksheetLoop()

   Dim Sheet As Worksheet

Application.DisplayAlerts = True

   For Each Sheet In Worksheets
     If Sheet.Name = "Clinical_Summary_Pivot" _
       Or Sheet.Name = "Validated_Summary_Pivot" _
       Or Sheet.Name = "Unmapped_Summary_Pivot" _
       Or Sheet.Name = "Combined Registry Measures" _
     Then
        Sheet.Delete
     End If
   Next Sheet

Application.DisplayAlerts = True

End Sub
