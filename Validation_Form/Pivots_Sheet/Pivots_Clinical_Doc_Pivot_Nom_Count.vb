Private Sub Pivots_Clin_Doc_NOMID_Code()
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

    Sheets("Clinical Documentation").Select

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                                      "Clinical_Table", Version:=6).CreatePivotTable TableDestination:= _
                                      "Pivots!R8C5", TableName:="Pivot_Clinical_Nom", DefaultVersion:=6

    Sheets("Pivots").Select

    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("Pivot_Clinical_Nom").PivotFields("DocumentType")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Pivot_Clinical_Nom").PivotFields("Notes")
        .Orientation = xlPageField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("Pivot_Clinical_Nom").PivotFields("EventDisplay")
        .Orientation = xlRowField
        .Position = 1
    End With

    ActiveSheet.PivotTables("Pivot_Clinical_Nom").AddDataField ActiveSheet.PivotTables( _
                                                               "Pivot_Clinical_Nom").PivotFields("NomenclatureID"), "Count of NomenclatureID", xlCount

    With ActiveSheet.PivotTables("Pivot_Clinical_Nom").PivotFields("EventDisplay")
    End With
    Range("F2").Select
End Sub
