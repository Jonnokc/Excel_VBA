Sub Check_Number()

    Dim Temp_Cell As Variant
    Dim Rng As Range

    Sheets("Validated Codes").Select    'Selects Sheet

    Range("D1:D" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select    'Selects all cells not empty in column
    Selection.Name = "Codes"    'Names Range

    Set Rng = Range("Codes")    'Assigns range to variable

    For Each cell In Rng    'Loops through cells in range

        If IsNumeric(cell) Then    'If cell contains numbers then X
            cell.Select    'Select the cell
            With Selection    'With the selected cell convert cell format to number without any decimal places
                Selection.NumberFormat = "0"
                .Value = .Value
            End With

        End If
    Next cell


End Sub
