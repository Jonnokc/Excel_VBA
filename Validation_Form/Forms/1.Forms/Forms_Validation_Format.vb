Sub Validation_Format()
'
'Loops through the table and deletes all the extra header rows



    Dim RowToTest As Long
    Dim tbl As ListObject
    Dim Rng As Range
    Dim rList As Range

    'Disables settings to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    MsgBox ("Program is about to run. Please leave computer alone until completed")

    Sheets("Validated Mappings").Select    'Selects Sheet
    ActiveSheet.AutoFilterMode = False

    'If table exists on this sheet, then convert to range
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

    Range("D1:D" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row).Select    'Selects all cells not empty in column
    Selection.Name = "Validated_Code_ID"    'Names Range

    Set Rng = Range("Validated_Code_ID")    'Assigns range to variable

    For Each cell In Rng
        If IsNumeric(cell) Then
            cell.Value = Val(cell.Value)
            cell.NumberFormat = "0"
        End If
    Next cell


    'Removes extra formatting of cells and standardizes all cells in the same format then formats 'range as table.

    Sheets("Validated Results").Select
    Cells.Select
    Selection.Style = "Normal"
    ActiveSheet.AutoFilterMode = False    'Turns off autofilter

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

    Range("I1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Forms_Val"
    tbl.TableStyle = "TableStyleLight12"

    'changes font color of header row to white
    Rows("1:1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With


    For RowToTest = Cells(Rows.Count, 2).End(xlUp).Row To 2 Step -1

        With Cells(RowToTest, 2)
            If .Value = "FORM_DEFINITION" Then
                Rows(RowToTest).EntireRow.Delete
            End If
        End With
    Next RowToTest


    Range("A2").Select
    ActiveCell = _
    "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(E2,'Validated Mappings'!D:D,0)),0)"
    Selection.AutoFill Destination:=Range("Forms_Val[Mapped?]")

    'Re-enables calculations
    Application.Calculation = xlCalculationAutomatic

    Columns("A:A").Select
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
    Range("A2").Select

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox ("Conpleted")


End Sub
