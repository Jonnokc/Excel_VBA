Private Sub Apply_Pivot_Format_As_Table()

' Formats Pivot table 1 into a table to improve readability and allow for filtering and autopopulation of rows.


    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Pivot")

    Sheets("Pivot").Select
    Range("A1:C" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

    Ws.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "PivotMain"
    Ws.ListObjects("PivotMain").TableStyle = "TableStyleLight9"

End Sub
