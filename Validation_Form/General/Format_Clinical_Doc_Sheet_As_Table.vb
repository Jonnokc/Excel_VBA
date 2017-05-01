Sub Format_Clin_Doc_Sht_As_Tbl()
'
' Creates the clinical doc pivot table on the Pivots Sheet within the validation form

    Dim lastrow As Long
    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim LastColumn As Long
    Dim StartCell As Range
    Dim rList As Range

    Sheets("Clinical Documentation").Select    'Selects the clinical doc sheet

    'If AutoFilters are on turn them off
    If ActiveSheet.AutoFilterMode = True Then
        ActiveSheet.AutoFilterMode = False
    End If

    'Checks the current sheet. If it is in table format, convert it to standard format.
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

    Set sht = Worksheets("Clinical Documentation")
    Set StartCell = Range("A2")

    'Refresh UsedRange
    Worksheets("Clinical Documentation").UsedRange

    'Find Last Row and Column
    lastrow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(lastrow, LastColumn)).Select

    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Clinical_Table"
    tbl.TableStyle = "TableStyleLight12"


End Sub
