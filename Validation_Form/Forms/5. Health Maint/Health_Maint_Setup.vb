Sub Health_Maint_Setup()


    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim StartCell As Range
    Dim rList As Range

    MsgBox ("Program is about to run. This will take about a minute or more to complete. Please leave computer alone until completed")

    'Disables settings to improve performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Sheets("Validated Mappings").Select

    ActiveSheet.AutoFilterMode = False    'Disables autoFilter

    'If table exists on sheet then convert to range
    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist
        End With

        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
        End With

    End If

    Set sht = Worksheets("Validated Mappings")
    Set StartCell = Range("A1")

    'Refresh UsedRange
    Worksheets("Validated Mappings").UsedRange

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Validated_tbl"
    tbl.TableStyle = "TableStyleLight12"

    'changes font color of header row to white
    Rows("1:1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    Application.Goto Reference:="Validated_tbl"
    Set Rng = Selection

    For Each cell In Rng    'Loops through cells in range
        If IsNumeric(cell) Then
            cell.Value = Val(cell.Value)
            cell.NumberFormat = "0"
        End If
    Next cell

    Rows("1:1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With


    Sheets("Health Maint Clinical Events").Select

    ActiveSheet.AutoFilterMode = False    'Disables autoFilter

    'If table exists on sheet then convert to range
    If ActiveSheet.ListObjects.Count > 0 Then

        With ActiveSheet.ListObjects(1)
            Set rList = .Range
            .Unlist
        End With

        With rList
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Borders.LineStyle = xlLineStyleNone
        End With

    End If

    Set sht = Worksheets("Health Maint Clinical Events")
    Set StartCell = Range("A1")

    'Refresh UsedRange
    Worksheets("Health Maint Clinical Events").UsedRange

    'Find Last Row and Column
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
    LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

    'Select Range
    sht.Range(StartCell, sht.Cells(LastRow, LastColumn)).Select

    'Turn selected Range Into Table
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    tbl.Name = "Health_Maint_Tbl"
    tbl.TableStyle = "TableStyleLight12"

    Application.Goto Reference:="Health_Maint_Tbl"
    Set Rng = Selection

    For Each cell In Rng    'Loops through cells in range
        If IsNumeric(cell) Then
            cell.Value = Val(cell.Value)
            cell.NumberFormat = "0"
        End If
    Next cell

    Range("A2").Select
    ActiveCell = "=IFERROR(INDEX('Validated Mappings'!I:I,MATCH(BI2,'Validated Mappings'!D:D,0)),0)"
    Selection.AutoFill Destination:=Range("Health_Maint_Tbl[Mapped?]")

    'Re-enables Auto-calculate for forumlas
    Application.Calculation = xlCalculationAutomatic

    Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Rows("1:1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    'Filters to remove Validated
    ActiveSheet.ListObjects("Health_Maint_Tbl").Range.AutoFilter Field:=1, _
                                                                 Criteria1:="0"

    'Filters to remove lines which do not have an event code
    ActiveSheet.ListObjects("Health_Maint_Tbl").Range.AutoFilter Field:=61, _
                                                                 Criteria1:="<>0", Operator:=xlAnd

    Range("A2").Select

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox ("Program Completed")

End Sub
