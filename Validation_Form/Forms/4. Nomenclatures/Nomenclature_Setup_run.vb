Sub Nomenclature_Auto()

    Dim Confirm_Run As Integer

    Confirm_Run = MsgBox("The program is about to run. This will take roughly 2 minutes." & vbNewLine & vbNewLine & "Please verify before running that you have entered all the needed data per the automation Instructions. If you have not please click Cancel. Else click Ok to run. ", vbOKCancel + vbQuestion, "Empty Sheet")

    'If user hits cancel then close program.
    If Confirm_Run = vbCancel Then
        MsgBox ("Program is canceling per user action.")
        Exit Sub
    End If

    Call Nomenclature_Row_Finder
    Call Nomenclature_Notes

    MsgBox ("Program Completed")

End Sub
