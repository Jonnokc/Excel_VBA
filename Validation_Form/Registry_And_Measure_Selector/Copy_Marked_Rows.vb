Sub Copy_Marked_Rows()

' Macro removes filters and then copies all rows which are flagged as "Yes" from the Raw_Concept_To_Measure sheet onto the additional tabs for proper distribution.

    answer = MsgBox("You are about to copy the flagged rows and populate the Workbook. Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

    If answer = vbYes Then

        Application.ScreenUpdating = False

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
        Selection.Copy

        'Pastes the selected cells on the clinical Documentation Tab
        Sheets("Clinical Documentation").Select
        Range("A3").Select
        ActiveSheet.Paste
        Rows("3:3").Select
        Selection.Delete Shift:=xlUp

        'Pastes cells on Unmapped Codes tab
        Sheets("Unmapped Codes").Select
        Range("A3").Select
        ActiveSheet.Paste
        Rows("3:3").Select
        Selection.Delete Shift:=xlUp

        'Pastes cells on Potential Mapping issues tab
        Sheets("Potential Mapping Issues").Select
        Range("A3").Select
        ActiveSheet.Paste
        Rows("3:3").Select
        Application.CutCopyMode = False
        Selection.Delete Shift:=xlUp


    Else

        'do nothing

    End If

    Application.ScreenUpdating = True

End Sub
