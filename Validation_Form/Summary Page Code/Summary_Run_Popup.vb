Sub Summary_Sheet_Setup()

'Prompts user to confirm they have reviewed the data in the validation form BEFORE running this.
    Confirm_Scrubbed = MsgBox("You have initiated the program to initalize the Summary Sheet. Please click ""Ok"" to run or ""Cancel"" to close the program", vbOKCancel + vbQuestion, "Empty Sheet")

    'If user hits cancel then close program.
    If Confirm_Scrubbed = vbCancel Then
        MsgBox ("Program is canceling per user action.")
        Exit Sub

    End If


    Call Summary_Create_Lookup_Sheet
    Call Summary_Combined_Lookup_Sheet
    Call Summary_Sheet_Initial_Setup
    Call Summary_Pop_Dots
    Call Summary_Cleanup
    Call Remove_Table_Format

    Sheets("Summary View").Select

    MsgBox ("Program Completed")

End Sub
