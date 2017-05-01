Private Sub Set_Pivot_X_Y_Range()

' Creates a named range to be used to populate the Yes/No data validation dropdown

    Sheets("Pivot").Select

    Range("E2:E" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select
    Selection.Name = "Pivot_Y_N_Range"
End Sub
