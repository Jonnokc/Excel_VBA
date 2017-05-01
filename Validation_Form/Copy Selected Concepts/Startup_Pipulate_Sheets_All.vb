Private Sub Copy_Marked_Rows()

' Macro removes filters and then copies all rows which are flagged as "Yes" from the Raw_Concept_To_Measure sheet onto the additional tabs for proper distribution.


    Sheets("Pivot").Select
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    Sheets("Raw_Concept_To_Measure").Select

    ActiveSheet.ListObjects("Raw_Table_Main").Range.AutoFilter Field:=4, _
                                                               Criteria1:="Yes"
    Range("Raw_Table_Main[[#Headers],[Registry Friendly Name]:[Concept Alias]]"). _
            Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Clinical Documentation").Select
    Range("B3").Select
    ActiveSheet.Paste

    Sheets("Unmapped Codes").Select
    Range("B3").Select
    ActiveSheet.Paste

    Sheets("Potential Mapping Issues").Select
    Range("B3").Select
    ActiveSheet.Paste

    'Deletes the header row which was included in the paste
    Rows("3:3").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp

    Sheets("Unmapped Codes").Select
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp

    Sheets("Clinical Documentation").Select
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp

End Sub

Private Sub Clin_Docum_Sheet_Setup()

    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range

    Application.ScreenUpdating = False

    Sheets("Clinical Documentation").Select

    Set sht = Worksheets("Clinical Documentation")
    Set StartCell = Range("A2")

    'Refresh UsedRange
    Worksheets("Clinical Documentation").UsedRange

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Clinical_Table"
    tbl.TableStyle = "TableStyleLight9"

    Range("A2").Select

End Sub


Private Sub Potential_Issues_Sheet_Setup()

    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range

    Application.ScreenUpdating = False

    Sheets("Potential Mapping Issues").Select

    Set sht = Worksheets("Potential Mapping Issues")
    Set StartCell = Range("A2")

    'Refresh UsedRange
    Worksheets("Potential Mapping Issues").UsedRange

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Potential_Issus_Table"
    tbl.TableStyle = "TableStyleLight9"

    Range("A2").Select

End Sub


Private Sub Unmapped_Codes_Sheet_Setup()

    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range

    Application.ScreenUpdating = False

    Sheets("Unmapped Codes").Select

    Set sht = Worksheets("Unmapped Codes")
    Set StartCell = Range("A2")

    'Refresh UsedRange
    Worksheets("Unmapped Codes").UsedRange

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Unmapped_Table"
    tbl.TableStyle = "TableStyleLight9"

    Range("A2").Select

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

    Range("A2").Select

End Sub


Sub Startup_Populate_Sheets()

' Startup script which prompts the user to confirm they intended to launch and then calls scripts in correct order to set up document.


    answer = MsgBox("This will populate the selected concepts on the other sheets. Are you read?", vbYesNo + vbQuestion, "Empty Sheet")

    If answer = vbYes Then

        Application.ScreenUpdating = False

        Call Copy_Marked_Rows
        Call Clin_Docum_Sheet_Setup
        Call Potential_Issues_Sheet_Setup
        Call Unmapped_Codes_Sheet_Setup
        Call Remove_Table_Format


    Else

        'do nothing

    End If
    Application.ScreenUpdating = True

    Sheets("Summary View").Select

End Sub
