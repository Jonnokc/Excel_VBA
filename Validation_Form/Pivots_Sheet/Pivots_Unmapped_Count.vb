Private Sub Pivots_Unmapped()
'
' Creates the clinical doc pivot table on the Pivots Sheet within the validation form

    Dim lastrow As Long
    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim LastColumn As Long
    Dim StartCell As Range
    Dim rList As Range

    Sheets("Unmapped Codes").Select    'Selects the clinical doc sheet

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

    Sheets("Unmapped Codes").Select
    ActiveWorkbook.Worksheets("Unmapped_Summary_Pivot").PivotTables( _
            "Unmapped_Pivot").PivotCache.CreatePivotTable TableDestination:="Pivots!R8C8" _
                                                                            , TableName:="Pivot_Unmapped", DefaultVersion:=6
    Sheets("Pivots").Select

    With ActiveSheet.PivotTables("Pivot_Unmapped").PivotFields("Measure")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Pivot_Unmapped").PivotFields("Raw Display")
        .Orientation = xlRowField
        .Position = 2
        .PivotItems("(blank)").Visible = False
    End With

    ActiveSheet.PivotTables("Pivot_Unmapped").AddDataField ActiveSheet.PivotTables( _
                                                           "Pivot_Unmapped").PivotFields("Raw Code"), "Count of Raw Code", xlCount
    ActiveWorkbook.ShowPivotTableFieldList = False

    Range("A2").Select
End Sub
