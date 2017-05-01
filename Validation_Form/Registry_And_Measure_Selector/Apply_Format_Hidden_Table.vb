Private Sub Apply_Format_Hidden_Table()

' Converts the second pivot table into a formatted table to allow columns to autopopulate additional rows with formulas.

    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Pivot")

    Sheets("Pivot").Select
    Range("AA1:AF" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

    Ws.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "HiddenTable"
    Ws.ListObjects("HiddenTable").TableStyle = "TableStyleLight9"

End Sub
