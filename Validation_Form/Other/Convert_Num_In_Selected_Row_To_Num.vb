Sub ConvertTextToNumber()

'TO USE: Click a cell within the column you wish to run the convert to numbers on.
'
'What this does: Loops through all cells in the column selected and if the cell is a number,
'then convert the format of that cell to a number.
'

    Dim Rng As Range
    Dim cell As Range

    MsgBox ("The number converter is about to run. Please leave your computer alone until the completed popup window.")

    'Helps improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    'Selects all cells in column
    Range(Cells(1, ActiveCell.Column), Cells(Rows.Count, ActiveCell.Column).End(xlUp)).Select
    Selection.Name = "Number_Check"    'Names Range
    Set Rng = Range("Number_Check")

    'If cell is a number, then convert format to number
    For Each cell In Rng
        If IsNumeric(cell) Then
            cell.Value = Val(cell.Value)
            cell.NumberFormat = "0"
        End If
    Next cell

    're-enables updates
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    'Notify of completion
    Range("A1").Select
    MsgBox ("Completed")
End Sub
