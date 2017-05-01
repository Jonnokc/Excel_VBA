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
    ActiveWorkbook.ShowPivotTableFieldList = True

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


Private Sub Insert_Pivot_2()

' Inserts second pivot table which is used to create CONCATENATE fields for matching to auto populate child rows.

    Dim PT As PivotTable

    Sheets("Raw_Pivot").Select

    Range("D1").Select
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveSheet.PivotTables("Concept_Pivot_Table").RepeatAllLabels xlRepeatLabels
    Set PT = ActiveSheet.PivotTables(1)
    PT.TableRange1.Select
    Selection.Copy
    Sheets("Pivot").Select
    Application.Goto Reference:="R1C27"
    Range("AA1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Range("A1").Select
End Sub

Private Sub set_rngList()

' Creates a named range for column A of the main pivot table for the group macro

    Sheets("Pivot").Select

    Range("A1:A" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

    ActiveWorkbook.Names.Add Name:="rngList", RefersToR1C1:="=Pivot!C1"
    ActiveWorkbook.Names("rngList").Comment = ""

End Sub


Private Sub Set_Pivot_X_Y_Range()

' Creates a named range to be used to populate the Yes/No data validation dropdown

    Sheets("Pivot").Select

    Range("E2:E" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select
    Selection.Name = "Pivot_Y_N_Range"
End Sub


Private Sub Set_Raw_Table_Range()

' Creates a named raw table range that is used to apply table formatting


    Set Ws = ThisWorkbook.Sheets("Raw_Concept_To_Measure")
    Sheets("Raw_Concept_To_Measure").Select

    Range("A1:E" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select
    Selection.Name = "Raw_Table_Range"

End Sub


Private Sub Remove_Duplicates()

' A duplicate checker. It runs against the Raw Concept To Measure file looking for duplicates across col A, B, C.

    Sheets("Raw_Concept_To_Measure").Select
    Application.Goto Reference:="Raw_Table_Range"
    Selection.RemoveDuplicates Columns:=Array(1, 2, 3), _
            Header:=xlYes
End Sub

Private Sub Apply_Format_Hidden_Table()

' Converts the second pivot table into a formatted table to allow columns to autopopulate additional rows with formulas.

    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Pivot")

    Sheets("Pivot").Select
    Range("AA1:AF" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

    Ws.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "HiddenTable"
    Ws.ListObjects("HiddenTable").TableStyle = "TableStyleLight9"

End Sub



Private Sub Apply_Format_Main_Raw_Table()

' Formats main raw table as a formatted table to allow for filtering later on.

    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Raw_Concept_To_Measure")

    Sheets("Raw_Concept_To_Measure").Select
    Range("A1:K" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

    Ws.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "Raw_Table_Main"
    Ws.ListObjects("Raw_Table_Main").TableStyle = "TableStyleLight9"

End Sub


Private Sub Apply_Pivot_Format_As_Table()

' Formats Pivot table 1 into a table to improve readability and allow for filtering and autopopulation of rows.


    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Pivot")

    Sheets("Pivot").Select
    Range("A1:C" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

    Ws.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "PivotMain"
    Ws.ListObjects("PivotMain").TableStyle = "TableStyleLight9"

End Sub


Private Sub Apply_Additional_Columns_And_Formulas()

' Overall applies column headers and formulas to additional rows

    Sheets("Raw_Concept_To_Measure").Select

    ' Creates the column headers for the raw concet to measure sheet

    Range("D1").Select
    ActiveCell.Formula = "Flagged_Result"
    Range("E1").Select
    ActiveCell.Formula = "Registry Check"
    Range("F1").Select
    ActiveCell.Formula = "Measure Check"
    Range("G1").Select
    ActiveCell.Formula = "Concept Check"
    Range("H1").Select
    ActiveCell.Formula = "Registry and Measure"
    Range("J1").Select
    ActiveCell.Formula = "Combined - All"
    Range("K1").Select
    ActiveCell.Formula = "Flagged_Result2"

    ' Applies the formulas to the raw concept to measure additional columns

    Range("E2").Select
    ActiveCell.Formula = _
            "=INDEX(Pivot!AF:AF,MATCH(Raw_Concept_To_Measure!A2,Pivot!$AA:$AA,0))"

    Range("F2").Select
    ActiveCell.Formula = _
            "=INDEX(Pivot!AF:AF,MATCH(Raw_Concept_To_Measure!H2,Pivot!AD:AD,0))"

    Range("G2").Select
    ActiveCell.Formula = _
            "=INDEX(Pivot!AF:AF,MATCH(Raw_Concept_To_Measure!J2,Pivot!AE:AE,0))"

    Range("H2").Select
    ActiveCell.Formula = "=CONCATENATE(A2,""|"",B2)"

    Range("J2").Select
    ActiveCell.Formula = "=CONCATENATE(A2,""|"",B2,""|"",C2)"

    Range("K2").Select
    ActiveCell.Formula = "=D2"

    Range("D2").Select
    ActiveCell.Formula = _
            "=IF(G2=""Yes"", ""Yes"",IF(AND(E2=""Yes"",F2<>""No"",G2<>""No""),""Yes"",IF(AND(F2=""Yes"",G2<>""No""),""Yes"",""No"")))"

    ' Creates the column headers for the pivot table

    Sheets("Pivot").Select
    Range("D1").Select
    ActiveCell.Formula = "Flagged"
    Range("E1").Select
    ActiveCell.Formula = "Y/N"
    Range("AD1").Select
    ActiveCell.Formula = "Registry and Measure"
    Range("AE1").Select
    ActiveCell.Formula = "All Combined"
    Range("AG1").Select
    ActiveCell.Formula = "Result"

    ' Applies the formulas for the pivot table new columns

    Range("AD2").Select
    ActiveCell.Formula = "=CONCATENATE(AA2,""|"",AB2)"
    Range("AE2").Select
    ActiveCell.Formula = "=CONCATENATE(AA2,""|"",AB2,""|"",AC2)"
    Range("AF2").Select
    ActiveCell.Formula = "=E2"
    Range("D2").Select
    ActiveCell.Formula = _
            "=IFERROR(VLOOKUP(AE2,Raw_Concept_To_Measure!J:K,2,""FALSE""),"""")"

    Range("A1").Select

End Sub


Private Sub GroupCells()

    Dim myRange As Range
    Dim rowCount As Integer
    Dim currentRow As Integer
    Dim firstBlankRow As Integer
    Dim lastBlankRow As Integer
    Dim currentRowValue As String
    Dim neighborColumnValue As String

    'select range based on given named range
    Set myRange = Range("rngList")
    rowCount = Cells(Rows.Count, myRange.Column).End(xlUp).Row

    firstBlankRow = 0
    lastBlankRow = 0
    'for every row in the range
    For currentRow = 1 To rowCount
        currentRowValue = Cells(currentRow, myRange.Column).Value

        If (IsEmpty(currentRowValue) Or currentRowValue = "") Then
            'if cell is blank and firstBlankRow hasn't been assigned yet
            If firstBlankRow = 0 Then
                firstBlankRow = currentRow
            End If
        ElseIf Not (IsEmpty(currentRowValue) Or currentRowValue = "") Then
            'if the cell is not blank,
            'and firstBlankRow hasn't been assigned, then this is the firstBlankRow
            'to consider for grouping
            If firstBlankRow = 0 Then
                firstBlankRow = currentRow
            ElseIf firstBlankRow <> 0 And currentRowValue <> "" Then
                'if firstBlankRow is assigned and this row has a value then the cell one row above this one is to be considered
                'the lastBlankRow to include in the grouping
                lastBlankRow = currentRow - 1
            End If
        End If

        'if first AND last blank rows have been assigned, then create a group
        'then reset the first/lastBlankRow values to 0 and begin searching for next
        'grouping
        If firstBlankRow <> 0 And lastBlankRow <> 0 Then
            Range(Cells(firstBlankRow, myRange.Column), Cells(lastBlankRow, myRange.Column)).EntireRow.Select
            Selection.Group
            firstBlankRow = 0
            lastBlankRow = 0
        End If
    Next
End Sub


Private Sub ungroup_first_row()

' Group Macro creates a one line group with the header file. This macro deletes that grouping to 'keep things clean.

    Sheets("Pivot").Select
    Range("A1").Select
    Selection.Rows.Ungroup
End Sub


Private Sub Apply_Dropdown()

' Applies the dropdown data validation to the Y/N column on the main pivot table.

    Sheets("Pivot").Select
    Range("E2").Select
    Application.Goto Reference:="Pivot_Y_N_Range"
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=Yes_No!$A$3:$A$5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("A2").Select

End Sub


Private Sub setupCompleted()

'Informs the user that setup has been completed.

    MsgBox ("Setup is completed. Flag away!")

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
        ActiveSheet.Range("2:2").AutoFilter
    End If

    Sheets("Unmapped Codes").Select

    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist                           ' convert the table back to a range
        End With

    End If

    If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
        ActiveSheet.Range("2:2").AutoFilter
    End If

    Sheets("Clinical Documentation").Select

    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist                           ' convert the table back to a range
        End With

    End If

    If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
        ActiveSheet.Range("2:2").AutoFilter
    End If

End Sub


Sub Startup_Create_Selector()

' Startup script which prompts the user to confirm they intended to launch and then calls scripts in correct order to set up document.


    answer = MsgBox("You are about to launch the startup script for the unmapped codes. Leave computer alone until completed. Are you ready?", vbYesNo + vbQuestion, "Empty Sheet")

    If answer = vbYes Then

        Application.ScreenUpdating = False

        Call Insert_Pivot_1
        Call Insert_Pivot_2
        Call set_rngList
        Call Set_Pivot_X_Y_Range
        Call Set_Raw_Table_Range
        Call Remove_Duplicates
        Call Apply_Format_Main_Raw_Table
        Call Apply_Format_Hidden_Table
        Call Apply_Pivot_Format_As_Table
        Call Apply_Additional_Columns_And_Formulas
        Call GroupCells
        Call ungroup_first_row
        Call Apply_Dropdown
        Call Remove_Table_Format
        Call setupCompleted

    Else

        'do nothing

    End If
    Application.ScreenUpdating = True

    Sheets("Pivot").Select

End Sub
