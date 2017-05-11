Sub Copy_Marked_Rows()

' Macro removes filters and then copies all rows which are flagged as "Yes" from the Raw_Concept_To_Measure sheet onto the additional tabs for proper distribution.

Dim sht As Worksheet

Application.ScreenUpdating = False

    answer = MsgBox("You are about to copy the flagged rows and populate the Workbook. Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

    If answer = vbYes Then

        'Application.ScreenUpdating = False

        Sheets("Pivot").Select
        ActiveSheet.Outline.ShowLevels RowLevels:=2
        Sheets("Raw_Concept_To_Measure").Select

        'Selects the Registry, Measure, and Concept columns with "Yes" filtered
        ActiveSheet.ListObjects("Raw_Table_Main").Range.AutoFilter Field:=4, _
                Criteria1:="Yes"
        Range("Raw_Table_Main[[#Headers],[Registry Friendly Name]:[Concept Alias]]"). _
                Select

        'Copies the selected cells
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy Sheets("Clinical Documentation").Range("A3")
        Selection.Copy Sheets("Unmapped Codes").Range("A3")
        Selection.Copy Sheets("Potential Mapping Issues").Range("A3")

        'Pastes the selected cells on the clinical Documentation Tab
        Set sht = Worksheets("Clinical Documentation")
        With sht
          sht.Range("3:3").Delete Shift:=xlUp
        End With

        'Pastes cells on Unmapped Codes tab
        Set sht = Worksheets("Unmapped Codes")
        With sht
          sht.Range("3:3").Delete Shift:=xlUp
        End With

        'Pastes cells on Potential Mapping issues tab
        Set sht = Worksheets("Potential Mapping Issues")
        With sht
          sht.Range("3:3").Delete Shift:=xlUp
        End With


    Else

        'do nothing

    End If

    Sheets("Unmapped Codes").Select
    Application.ScreenUpdating = True

End Sub
