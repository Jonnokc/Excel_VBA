Private Sub Insert_Pivot_1()

' Inserts clean pivot table 1 which is used to cleanly view and flaw registries, measures and concepts.

    Dim PT As PivotTable

    Columns("A:C").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Raw_Concept_To_Measure!R1C1:R1048576C3", Version:=6).CreatePivotTable _
            TableDestination:="Raw_Pivot!R1C1", TableName:="Concept_Pivot_Table", _
            DefaultVersion:=6

    Sheets("Raw_Pivot").Select
    Cells(1, 1).Select

    ActiveWorkbook.ShowPivotTableFieldList = Truez
    With ActiveSheet.PivotTables("Concept_Pivot_Table").PivotFields( _
            "Registry Friendly Name")
        .Orientation = xlRowField
        .Position = 1
    End With

    With ActiveSheet.PivotTables("Concept_Pivot_Table").PivotFields( _
            "Measure Friendly Name")
        .Orientation = xlRowField
        .Position = 2
    End With

    With ActiveSheet.PivotTables("Concept_Pivot_Table").PivotFields("Concept Alias")
        .Orientation = xlRowField
        .Position = 3
    End With

    Range("A3").Select
    ActiveSheet.PivotTables("Concept_Pivot_Table").RowAxisLayout xlOutlineRow
    ActiveSheet.PivotTables("Concept_Pivot_Table").PivotFields("Registry Friendly Name"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
    ActiveSheet.PivotTables("Concept_Pivot_Table").PivotFields("Measure Friendly Name"). _
            Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
            False, False)
    ActiveSheet.PivotTables("Concept_Pivot_Table").PivotFields("Concept Alias").Subtotals _
            = Array(False, False, False, False, False, False, False, False, False, False, False, False _
            )

    With ActiveSheet.PivotTables("Concept_Pivot_Table")
        .ColumnGrand = False
        .RowGrand = False
    End With

    Set PT = ActiveSheet.PivotTables(1)
    Range("A1").Select
    ActiveWorkbook.ShowPivotTableFieldList = False
    PT.TableRange1.Select
    Selection.Copy
    Sheets("Pivot").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Range("A1").Select
End Sub
