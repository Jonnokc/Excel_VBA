Private Sub set_rngList()

' Creates a named range for column A of the main pivot table for the group macro

    Sheets("Pivot").Select

    Range("A1:A" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

    ActiveWorkbook.Names.Add Name:="rngList", RefersToR1C1:="=Pivot!C1"
    ActiveWorkbook.Names("rngList").Comment = ""

End Sub
