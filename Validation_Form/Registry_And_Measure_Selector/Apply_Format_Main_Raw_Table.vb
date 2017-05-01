Private Sub Apply_Format_Main_Raw_Table()

' Formats main raw table as a formatted table to allow for filtering later on.

    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Raw_Concept_To_Measure")

    Sheets("Raw_Concept_To_Measure").Select
    Range("A1:K" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

    Ws.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "Raw_Table_Main"
    Ws.ListObjects("Raw_Table_Main").TableStyle = "TableStyleLight9"

End Sub
