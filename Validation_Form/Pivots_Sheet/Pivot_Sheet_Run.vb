Sub Pivot_Sheet_Setup()

    answer = MsgBox("This will run the Pivot Table Sheet Setup. Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

    If answer = vbYes Then

        Application.ScreenUpdating = False

        Call Pivots_Unmapped
        Call Pivots_Clin_Doc_EV_Code
        Call Pivots_Clin_Doc_NOMID_Code
        Call Remove_Table_Format

        Application.ScreenUpdating = True

    Else
        'Do Nothing

    End If

End Sub
