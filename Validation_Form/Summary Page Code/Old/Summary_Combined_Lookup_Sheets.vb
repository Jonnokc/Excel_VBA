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
            Msgbox("Program can not find worksheet - " & WkNames(1) & vbNewLine & vbNewLine & "This worksheet is required for the program to run. Please alter the program and/or the worksheet name then re-run the program.")
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
