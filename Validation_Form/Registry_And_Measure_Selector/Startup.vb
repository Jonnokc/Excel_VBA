Sub Startup_Create_Selector()

' Startup script which prompts the user to confirm they intended to launch and then calls scripts in correct order to set up document.


    answer = MsgBox("You are about to launch the startup script for the unmapped codes. Leave computer alone until completed. Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

    If answer = vbYes Then

        Application.ScreenUpdating = False

        Call Insert_Pivot_1
        Call Insert_Pivot_2
        Call set_rngList
        Call Set_Pivot_X_Y_Range
        Call Set_Raw_Table_Range
        Call Remove_Duplicates
        Call Apply_Format_Main_Raw_Table
        Call Apply_Format_Hidden_Table
        Call Apply_Pivot_Format_As_Table
        Call Apply_Additional_Columns_And_Formulas
        Call GroupCells
        Call ungroup_first_row
        Call Apply_Dropdown
        Call Remove_Table_Format
        Call setupCompleted

        Application.ScreenUpdating = True

    Else

        'do nothing

    End If

    Sheets("Pivot").Select

End Sub
