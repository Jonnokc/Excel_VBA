Range("C2:C" & Cells.SpecialCells(xlCellTypeLastCell).Row).Select


Dim LR As Long
LR = Range("A" & Rows.Count).End(xlUp).Row
Range("E2" & ":E" & LR).SpecialCells(xlCellTypeVisible).Select
