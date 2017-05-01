Private Sub Apply_Dropdown()

' Applies the dropdown data validation to the Y/N column on the main pivot table.

    Sheets("Pivot").Select
    Range("E2").Select
    Application.Goto Reference:="Pivot_Y_N_Range"
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:="=Yes_No!$A$3:$A$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("A2").Select

End Sub
