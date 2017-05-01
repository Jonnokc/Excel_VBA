Sub Delete_Extra_Sheets()

    Dim sheet As Worksheet

    Application.DisplayAlerts = False

    For Each sheet In Worksheets

        If sheet.Name = "Unmapped Codes" _
           Or sheet.Name = "Health Maintenance Summary" _
           Or sheet.Name = "Clinical Documentation" _
           Or sheet.Name = "Source_Code_Systems" _
           Or sheet.Name = "Sheet1" _
           Then
            sheet.Delete
        End If
    Next sheet

    Application.DisplayAlerts = True

End Sub
