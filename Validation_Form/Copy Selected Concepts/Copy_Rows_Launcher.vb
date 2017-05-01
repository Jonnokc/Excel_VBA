Sub Startup_Populate_Sheets()

' Startup script which prompts the user to confirm they intended to launch and then calls scripts in correct order to set up document.


    answer = MsgBox("This will populate the selected concepts on the other sheets. Are you read?", vbYesNo + vbQuestion, "Empty Sheet")

    If answer = vbYes Then

        Application.ScreenUpdating = False

        Call Copy_Marked_Rows
        Call Clin_Docum_Sheet_Setup
        Call Potential_Issues_Sheet_Setup
        Call Unmapped_Codes_Sheet_Setup
        Call Remove_Table_Format


    Else

        'do nothing

    End If
    Application.ScreenUpdating = True

    Sheets("Summary View").Select

End Sub
