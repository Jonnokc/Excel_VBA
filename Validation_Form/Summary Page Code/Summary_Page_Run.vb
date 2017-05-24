Private Sub Summary_Cleanup()

' Delets extra sheets and columns after summary sheet process has completed


Dim Sheet As Worksheet
Dim header_location As Variant

Sheets("Summary View").Select

Range("A1:K1").Name = "Header_Row"

For Each Header In Range("Header_Row")
  ' Deletes the Concat column
  If Header = "Concat" Then
    header_location = Mid(Header.Address, 2, 1)
    Columns(header_location & ":" & header_location).Select
    Selection.Delete Shift:=xlToLeft
  End If
Next Header

' Delete the extra sheets
Application.DisplayAlerts = False

For Each Sheet In Worksheets

    If Sheet.Name = "Potential_Summary_Pivot" _
       Or Sheet.Name = "Clinical_Summary_Pivot" _
       Or Sheet.Name = "Unmapped_Summary_Pivot" _
       Or Sheet.Name = "Combined Registry Measures" _
       Then
        Sheet.Delete
    End If
Next Sheet

    Application.DisplayAlerts = True

End Sub



Private Sub Summary_Combined_Lookup_Sheet()
'
' Takes the Registries, Measures and Concepts from the Unmapped and Validated Sheets and combinds them into one sheet.
' Then creates a CONCATENATE column for lookup.
'
    Dim WkNames As Variant
    Dim HeaderNames As Variant
    Dim DataRange As Variant
    Dim Next_Blank_Row As Long
    Dim counter As Long
    Dim tbl As ListObject
    Dim Sheet As Worksheet


    'This disables settings to improve macro performance.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    WkNames = Array("Potential Mapping Issues", "Unmapped Codes")
    HeaderNames = Array("Registry", "Measure", "Concept", "Concat", "Potential_Lookup", "Unmapped_Lookup", "Clinical_Lookup")

    'Loops through all worksheets and checks the worksheet names for a match against the array.
    For i = 0 To UBound(WkNames)
        WkNamesCheck = False

        For Each Sheet In Worksheets
            If Sheet.Name = WkNames(i) Then
                WkNamesCheck = True
                Exit For
            End If
        Next Sheet

        'If the worksheet does not exist tell the user to fix the issue then end the program
        If WkNamesCheck = False Then
            MsgBox ("Program can not find worksheet - " & WkNames(1) & vbNewLine & vbNewLine & "This worksheet is required for the program to run. Please alter the program and/or the worksheet name then re-run the program.")
            End
        End If

    Next i

    'Deletes the Sheets if they already exist to allow user to re-run program
    Application.DisplayAlerts = False

    For Each Sheet In Worksheets
        If Sheet.Name = "Combined Registry Measures" Then
            Sheet.Delete
        End If
    Next Sheet

    Application.DisplayAlerts = True

    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Combined Registry Measures"
    End With

    'Populates headers on the new Worksheet
    Sheets("Combined Registry Measures").Select
    Range("A1:G1").Select
    Selection.Name = "Header_Range"


    counter = 0
    'Populates the header row
    For Each cell In Range("Header_Range")
        cell.Value = HeaderNames(counter)
        counter = counter + 1

    Next cell



    For i = 0 To UBound(WkNames)

        CurrentWk = WkNames(i)

        Sheets(CurrentWk).Select
        Range("A3:B3").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy

        Sheets("Combined Registry Measures").Select
        Next_Blank_Row = Range("A" & Rows.Count).End(xlUp).Row + 1
        Range("A" & Next_Blank_Row).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

    Next i

    'Creates a named table from selected range
    Range("A1:G" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select

    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "combined_lookup_range"
    tbl.TableStyle = "TableStyleLight12"


    ActiveSheet.Range("combined_lookup_range[#All]").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes

    Range("D2").Select
    ActiveCell.Formula = "=CONCATENATE(A2,""|"",B2)"

    Range("E2").Select
    ActiveCell.Formula = _
            "=IFERROR(INDEX(Potential_Summary_Pivot!C:C,MATCH(D2,Potential_Summary_Pivot!D:D,0)),0)"

    Range("F2").Select
    ActiveCell.Formula = _
            "=IFERROR(INDEX(Unmapped_Summary_Pivot!C:C,MATCH(D2,Unmapped_Summary_Pivot!D:D,0)),0)"

    Range("G2").Select
    ActiveCell.Formula = _
            "=IFERROR(INDEX(Clinical_Summary_Pivot!C:C,MATCH(D2,Clinical_Summary_Pivot!D:D,0)),0)"


    'Re-enables previously disabled settings after all code has run.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True


End Sub

Private Sub Summary_Create_Lookup_Sheet()


    Dim wb As Workbook
    Dim Sheet As Worksheet
    Dim Table_Obj As ListObject
    Dim StartCell As Range
    Dim WkNames As Variant
    Dim TblNames As Variant
    Dim PivotNames As Variant
    Dim PivotSheetNames As Variant
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim rList As Range
    Dim WkExistCheck As Variant


    'DEBUG

    'This disables settings to improve macro performance.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False


    WkNames = Array("Potential Mapping Issues", "Unmapped Codes", "Clinical Documentation")
    WkExistCheck = Array(False, False, False)
    TblNames = Array("Potential_Table", "Unmapped_Table", "Clinical_Table")
    PivotNames = Array("Potential_Pivot", "Unmapped_Pivot", "Clinical_Pivot")
    PivotSheetNames = Array("Potential_Summary_Pivot", "Unmapped_Summary_Pivot", "Clinical_Summary_Pivot")

    ' Unhides the needed Worksheets

    For Each Sheet In Worksheets
      For i = 0 To UBound(WkNames)
        If WkNames(i) = Sheet.Name Then
          Sheet.Visible = xlSheetVisible
        End If
      Next i
    Next Sheet


    'Deletes the Sheets if they already exist to allow user to re-run program
    Application.DisplayAlerts = False

    For Each Sheet In Worksheets
        If Sheet.Name = "Clinical_Summary_Pivot" _
                Or Sheet.Name = "Potential_Summary_Pivot" _
                Or Sheet.Name = "Unmapped_Summary_Pivot" _
                Or Sheet.Name = "Combined Registry Measures" _
                Then
            Sheet.Delete
        End If

    Next Sheet

    Application.DisplayAlerts = True

    'Checks if Wk Exists

    For i = 0 To UBound(WkNames)
        On Error GoTo NoSheet

        Sheets(WkNames(i)).Select
        WkExistCheck(i) = True

NoSheet:
        Resume ClearError

ClearError:

    Next i


    ' Loop through each of the worksheets needed and format them in a standardized way
    ' That is used later on with different programs
    For i = 0 To UBound(WkNames)

        CurrentExistCheck = WkExistCheck(i)
        CurrentWkName = WkNames(i)
        CurrentTblName = TblNames(i)
        CurrentPivotName = PivotNames(i)
        CurrentPivotSheetName = PivotSheetNames(i)

        If CurrentExistCheck = True Then

            Sheets(WkNames(i)).Select

            If ActiveSheet.AutoFilterMode = True Then
                ActiveSheet.AutoFilterMode = False
            End If

            'Checks the current sheet. If it is in table format, convert it to range.
            If ActiveSheet.ListObjects.Count > 0 Then
                With ActiveSheet.ListObjects(1)
                    Set rList = .Range
                    .Unlist
                End With
                'Reverts the color of the range back to standard.
                With rList
                    .Interior.ColorIndex = xlColorIndexNone
                    .Font.ColorIndex = xlColorIndexAutomatic
                    .Borders.LineStyle = xlLineStyleNone
                End With
            End If

            Set sht = Worksheets(WkNames(i))    'Sets value
            Set StartCell = Range("A2")    'Start cell used to determine where to begin creating the table range

            'Find Last Row and Column
            LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
            LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column
            Sheet_Name = WkNames(i)    'Assigns sheet name to a variable as a string

            'Select Range
            sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

            'Creates the table
            Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
            tbl.Name = TblNames(i)    'Names the table
            tbl.TableStyle = "TableStyleLight12"    'Sets table color theme

            Rows("2:2").Select
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With


            'Creates a new sheet which will house the validated codes pivot table
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = CurrentPivotSheetName
            End With

            Sheets(CurrentWkName).Select
            Range(CurrentTblName).Select
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                    CurrentTblName, Version:=6).CreatePivotTable TableDestination:= _
                    CurrentPivotSheetName & "!R1C1", TableName:=CurrentPivotName, DefaultVersion:=6

            Sheets(CurrentPivotSheetName).Select
            Cells(1, 1).Select


            ActiveSheet.PivotTables(CurrentPivotName).AddDataField ActiveSheet.PivotTables( _
                    CurrentPivotName).PivotFields("Source"), "Count of Source", xlCount

            With ActiveSheet.PivotTables(CurrentPivotName).PivotFields("Registry")
                .Orientation = xlRowField
                .Position = 1
            End With


            With ActiveSheet.PivotTables(CurrentPivotName).PivotFields("Measure")
                .Orientation = xlRowField
                .Position = 2
            End With

            'Sets pivot table layout to OUTLINE
            ActiveSheet.PivotTables(CurrentPivotName).RowAxisLayout xlOutlineRow

            'Turns on repeat blank lines
            ActiveSheet.PivotTables(CurrentPivotName).RepeatAllLabels xlRepeatLabels

            'Sets empty values to 0 which helps in a couple places! but also allows the below autofill to have a range reference'
            ActiveSheet.PivotTables(CurrentPivotName).NullString = "0"

            Range("D1").Select

            LastRow = ActiveSheet.Range("C2").End(xlDown).Row

            Sheets(CurrentPivotSheetName).Select
            Range("D2").Select
            ActiveCell.Formula = "=IF(B2 <>"""",CONCATENATE(A2,""|"",B2),"""")"

            With ActiveSheet.Range("D2")
                .AutoFill Destination:=Range("D2:D" & LastRow&)
            End With

        End If

    Next i

    'Re-enables previously disabled settings after all code has run.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Private Sub Summary_Pop_Dots()
'
' Copies the temporary values from the lookup columns and pastes the VALUES into the appropriate columns.

    Dim Sheet_Headers As Variant
    Dim Find_Header As Range
    Dim rngHeaders As Range
    Dim ColHeaders As Variant
    Dim Validated_Col As Variant
    Dim Unmapped_Col As Variant
    Dim Clinical_Col As Variant
    Dim Health_Col As Variant
    Dim WkNames As Variant
    Dim PivotNames As Variant
    Dim CombinedCopyColumns As Variant
    Dim SummaryColumns As Variant
    Dim HyperLinkSheets As Variant
    Dim HeaderNames As Variant
    Dim SummaryHeaderChecker As Variant
    Dim Header_Missing As Integer
    Dim FirstSum As Variant
    Dim EndSum As Variant


    ' This disables settings to improve macro performance.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Sheets("Summary View").Select
    Columns("E:H").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    ' Refreshes pivot table data
    WkNames = Array("Potential_Summary_Pivot", "Clinical_Summary_Pivot", "Unmapped_Summary_Pivot")
    PivotNames = Array("Potential_Pivot", "Clinical_Pivot", "Unmapped_Pivot")
    CombinedCopyColumns = Array("E2", "F2", "G2")
    SummaryColumns = Array(False, False, False)
    HyperLinkSheets = Array("'Potential Mapping Issues'", "'Clinical Documentation'", "'Unmapped Codes'")
    HeaderNames = Array("Potential Mapping Issues", "Unmapped Codes", "Clinical Documentation")
    SummaryHeaderChecker = Array(False, False, False)

    For i = 0 To UBound(WkNames)
        CurrentWk = WkNames(i)
        Sheets(CurrentWk).Select
        ActiveSheet.PivotTables(PivotNames(i)).PivotCache.Refresh
    Next i

    '
    ' PRIMARY - finds and stores summary header columns
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Summary View").Select
    Range("B1:J1").Select
    Selection.Name = "Header_Row"


    ' finds column letter for each of the colums we care about
    For Each cell In Range("Header_Row")

        If cell = "Potential Mapping Issues" Then
            SummaryColumns(0) = Mid(cell.Address, 2, 1)
            SummaryHeaderChecker(0) = True

        ElseIf cell = "Unmapped Codes" Then
            SummaryColumns(1) = Mid(cell.Address, 2, 1)
            SummaryHeaderChecker(1) = True

        ElseIf cell = "Clinical Documentation" Then
            SummaryColumns(2) = Mid(cell.Address, 2, 1)
            SummaryHeaderChecker(2) = True

            'Elseif cell = "Health Maintenance" Then
            'SummaryColumns(3) = Mid(cell.Address, 2, 1)
            'SummaryHeaderChecker(3) = True
        End If

    Next cell

    ' Prompts user to confirm they have reviewed the data in the validation form BEFORE running this.
    For i = 0 To UBound(SummaryHeaderChecker)

        If SummaryHeaderChecker(i) = False Then
            Header_Missing = MsgBox("It looks like a column is missing or has a different header name." & vbNewLine & vbNewLine & "Unable to find header " & HeaderNames(i) & vbNewLine & vbNewLine & "If the header column exists or should exist please click Cancel and update the column header accordingly then rerun. If the column is not needed and thus was deleted or hidden on purpose please click Ok to continue running the program", vbOKCancel + vbQuestion, "Empty Sheet")
        End If

        ' If user hits cancel then close program.
        If Header_Missing = vbCancel Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            MsgBox ("Program is canceling per user action.")
            Exit Sub
        End If
    Next i

    For i = 0 To UBound(CombinedCopyColumns)
        CurrentWk = WkNames(i)
        CurrentCopyCol = CombinedCopyColumns(i)
        CurrentSumCol = SummaryColumns(i)

        ' Confirms the column exists. If the column does not exist then skip it.
        If CurrentSumCol <> False Then

            Sheets("Combined Registry Measures").Select
            Range(CurrentCopyCol).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy

            Sheets("Summary View").Select
            Range(CurrentSumCol & "2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False

        End If

    Next i

    Range("B2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats

    ' Autofit for all cells on screen.
    Cells.Select
    Cells.EntireColumn.AutoFit

    ' SUB - If column exists then copy the data to the coresponding column on the summary sheet
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 0 To UBound(SummaryColumns)

        If SummaryColumns(i) <> False Then

        ' IMPORTANT - Removing conditional formatting for the time being per new file format
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Range(SummaryColumns(i) & "2").Select
            ' Range(Selection, Selection.End(xlDown)).Select
            ' Application.CutCopyMode = False
            ' Selection.FormatConditions.AddIconSetCondition
            ' Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            ' With Selection.FormatConditions(1)
            '     .ReverseOrder = True
            '     .ShowIconOnly = True
            '     .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
            ' End With
            '
            ' With Selection.FormatConditions(1).IconCriteria(2)
            '     .Type = xlConditionValueNumber
            '     .Value = 1
            '     .Operator = 7
            ' End With
            '
            ' With Selection.FormatConditions(1).IconCriteria(3)
            '     .Type = xlConditionValueNumber
            '     .Value = 4
            '     .Operator = 7
            ' End With

            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With

        End If
    Next i


    ' Adds the hyperlink address to the street lights
    For i = 0 To UBound(SummaryColumns)
        CurrentWk = WkNames(i)
        CurrentCopyCol = CombinedCopyColumns(i)
        CurrentSumCol = SummaryColumns(i)
        CurrentHyperSht = HyperLinkSheets(i)

        ' Confirms the column exists. If the column does not exist then skip it.
        If CurrentSumCol <> False Then

            Range(CurrentSumCol & "2").Select
            Range(Selection, Selection.End(xlDown)).Select
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
                    CurrentHyperSht & "!A1"

        End If

    Next i


    ' Formats the angle for the header row of Summary Sheet
    Rows("1:1").Select

    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 45
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    ' Autofit for all cells on screen.
    Cells.Select
    Cells.EntireColumn.AutoFit

    ' Cleans up selected cells on sheet.
    Range("A1").Select

    ' Re-enables previously disabled settings after all code has run.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub


Private Sub Summary_Sheet_Initial_Setup()
'
' Sets up the Summary Sheet. Copies the Registries and Measures and creates the concat column
'

    Dim tbl As ListObject
    Dim HeaderNames As Variant
    Dim HeaderLocations As Variant
    Dim rList As Range

    'This disables settings to improve macro performance.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    HeaderNames = Array("Registry", "Measure", "Concat")
    SummaryColumns = Array("Reg", "Meas", "Concat", "Key")

    Sheets("Summary View").Select

    ActiveSheet.AutoFilterMode = False

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

    'Clears formats
    Range("B1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ClearFormats

    'Clear cell values if there are analysis
    Sheets("Summary View").Select
    Range("A1:K1").Select
    Selection.Name = "Summary_Headers"

    For Each cell In Range("Summary_Headers")

        CurrentHeader = cell
        IsInHeaderArray = Not IsError(Application.Match(CurrentHeader, HeaderNames, 0))

        If IsInHeaderArray = True Then
            CurrentAddress = Mid(cell.Address, 2, 1)
            Range(CurrentAddress & "2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Clear
        End If

    Next cell

    Sheets("Summary View").Select
    Range("A1:M1").Select
    Selection.Name = "Header_Row"

    'finds column letter for each of the colums we care about
    For Each cell In Range("Header_Row")

        If cell = "Registry" Then
            SummaryColumns(0) = Mid(cell.Address, 2, 1)

        ElseIf cell = "Measure" Then
            SummaryColumns(1) = Mid(cell.Address, 2, 1)

        ElseIf cell = "Concat" Then
            SummaryColumns(2) = Mid(cell.Address, 2, 1)

        ElseIf cell = "Key" Then
            SummaryColumns(3) = Mid(cell.Address, 2, 1)

            'Elseif cell = "Health Maintenance" Then
            '  SummaryColumns(3) = Mid(cell.Address, 2, 1)
        End If

    Next cell

    'If Concat column has already been deleted. Re-Add the column
    If SummaryColumns(2) = "Concat" Then
        KeyCol = SummaryColumns(3)
        Columns(KeyCol & ":" & KeyCol).Select
        Selection.Insert Shift:=xlToRight
        ActiveCell = "Concat"
        ActiveCell(1).Select

        SummaryColumns(2) = Mid(ActiveCell.Address, 2, 1)
    End If


    'Copies the Registry and measure columns to the summary view sheet
    Sheets("Combined Registry Measures").Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Summary View").Select

    'Uses the location of the Registry column to paste the data
    Range(SummaryColumns(0) & "2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


    Range("B1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select


    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Summary_Table"
    tbl.TableStyle = "TableStyleLight13"

    'Changes header font back to white
    Rows("1:1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    'Uses the location of the concat column
    Range(SummaryColumns(2) & "2").Select
    ActiveCell.Formula = "=CONCATENATE(B2,""|"",C2)"


    'Re-enables previously disabled settings after all code has run.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True


End Sub


Private Sub Remove_Table_Format()

    Dim rList As Range
    Dim WkNames As Variant


    WkNames = Array("Potential Mapping Issues", "Unmapped Codes", "Clinidal Documentation")

    For i = 0 To UBound(WkNames)

        On Error GoTo NoSheet
        Sheets(WkNames(i)).Select

        If ActiveSheet.ListObjects.Count > 0 Then

            With ActiveSheet.ListObjects(1)
                Set rList = .Range
                .Unlist                           ' convert the table back to a range
            End With

            ' Changes header row font color to white
            Rows("2:2").Select
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With

        End If

        If Not ActiveSheet.AutoFilterMode Then  'Adds the filter buttons to the sheet
            ActiveSheet.Range("2:2").AutoFilter
        End If

        Range("A2").Select

        'Error handling incase sheet does not exist
NoSheet:
        'MsgBox("No Code for " & EventCode)
        Resume ClearError

ClearError:
        'Clears variables for next loop

    Next i

End Sub

Sub Summary_Sheet_Setup()

'Prompts user to confirm they have reviewed the data in the validation form BEFORE running this.
    Confirm_Scrubbed = MsgBox("You have initiated the program to initalize the Summary Sheet. Please click ""Ok"" to run or ""Cancel"" to close the program", vbOKCancel + vbQuestion, "Empty Sheet")

    'If user hits cancel then close program.
    If Confirm_Scrubbed = vbCancel Then
        MsgBox ("Program is canceling per user action.")
        Exit Sub

    End If


    Call Summary_Create_Lookup_Sheet
    Call Summary_Combined_Lookup_Sheet
    Call Summary_Sheet_Initial_Setup
    Call Summary_Pop_Dots
    Call Summary_Cleanup
    Call Remove_Table_Format

    Sheets("Summary View").Select

    MsgBox ("Program Completed")

End Sub
