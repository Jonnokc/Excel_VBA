Private Sub Format_Unmapped_Sht_As_Tbl()
'
' Creates unmapped pivot table which is used for our summary color sheet
'
    Dim lastrow As Long
    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim LastColumn As Long
    Dim StartCell As Range
    Dim rList As Range

    Sheets("Unmapped Codes").Select

    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist                           ' convert the table back to a range
        End With

        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
        End With

    End If

    Set sht = Worksheets("Unmapped Codes")
    Set StartCell = Range("A2")

    'Refresh UsedRange
    Worksheets("Unmapped Codes").UsedRange

    'Find Last Row and Column
    lastrow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(lastrow, LastColumn)).Select

    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Unmapped_Table"
    tbl.TableStyle = "TableStyleLight12"


End Sub
