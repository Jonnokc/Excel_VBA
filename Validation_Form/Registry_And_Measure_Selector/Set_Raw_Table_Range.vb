Private Sub Set_Raw_Table_Range()

' Creates a named raw table range that is used to apply table formatting


    Set Ws = ThisWorkbook.Sheets("Raw_Concept_To_Measure")
    Sheets("Raw_Concept_To_Measure").Select

    Range("A1:E" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select
    Selection.Name = "Raw_Table_Range"

End Sub
