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


Private Sub Pivots_Clin_Doc_EV_Code()
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
                                      "Pivots!R8C2", TableName:="Pivot_Clinical_Doc", DefaultVersion:=6

    Sheets("Pivots").Select

    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("Pivot_Clinical_Doc").PivotFields("DocumentType")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Pivot_Clinical_Doc").PivotFields("EventDisplay")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Pivot_Clinical_Doc").AddDataField ActiveSheet.PivotTables( _
                                                               "Pivot_Clinical_Doc").PivotFields("EventCode"), "Count of EventCode", xlCount
    With ActiveSheet.PivotTables("Pivot_Clinical_Doc").PivotFields("EventDisplay")
        .PivotItems("(blank)").Visible = False
    End With
    Range("F2").Select
End Sub


Private Sub Remove_Table_Format()

    Dim rList As Range

    Sheets("Potential Mapping Issues").Select

    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist                           ' convert the table back to a range
        End With

    End If

    If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
        ActiveSheet.Range("A2").AutoFilter
    End If

    Sheets("Unmapped Codes").Select

    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist                           ' convert the table back to a range
        End With

    End If

    If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
        ActiveSheet.Range("A2").AutoFilter
    End If

    Sheets("Clinical Documentation").Select

    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist                           ' convert the table back to a range
        End With

    End If

    If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
        ActiveSheet.Range("A2").AutoFilter
    End If

    Range("A2").Select


End Sub


Sub Pivot_Sheet_Setup()

    answer = MsgBox("This will run the Pivot Table Sheet Setup. Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

    If answer = vbYes Then

        Application.ScreenUpdating = False

        Call Pivots_Unmapped
        Call Pivots_Clin_Doc_EV_Code
        Call Pivots_Clin_Doc_NOMID_Code
        Call Remove_Table_Format

        Application.ScreenUpdating = True

    Else
        'Do Nothing

    End If

End Sub
